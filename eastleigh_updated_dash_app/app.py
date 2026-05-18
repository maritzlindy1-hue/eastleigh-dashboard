
from pathlib import Path
from collections import defaultdict

import pandas as pd
import dash
from dash import html, dash_table
import dash_cytoscape as cyto


EXCEL_FILE = Path(__file__).with_name("Eastleigh_Input.xlsx")
SHEET_NAME = 0


def clean_meter(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_kwh(value):
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def node_id(name, meter):
    name = clean_text(name)
    meter = clean_meter(meter)

    if meter and meter.lower() != "virtual meter":
        return meter

    return name.replace(" ", "_").replace("/", "_").replace("&", "and")


def load_data():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    required = [
        "Parent",
        "Child",
        "Parent Meter Serial Nr",
        "Child Meter Serial Nr",
        "Level",
        "kWh",
    ]

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns in Excel: {missing}")

    nodes = {}
    edges = []
    children_map = defaultdict(list)

    # Root / parent rows
    for _, row in df.iterrows():
        parent = clean_text(row["Parent"])
        child = clean_text(row["Child"])
        parent_meter = clean_meter(row["Parent Meter Serial Nr"])
        child_meter = clean_meter(row["Child Meter Serial Nr"])
        kwh = clean_kwh(row["kWh"])

        if not parent:
            continue

        parent_id = node_id(parent, parent_meter)

        if parent_id not in nodes:
            nodes[parent_id] = {
                "id": parent_id,
                "name": parent,
                "meter": parent_meter,
                "kwh": 0.0,
                "level": 1,
                "kind": "virtual" if parent_meter.lower() == "virtual meter" else "meter",
            }

        # Row with no child is the root/node own value
        if not child:
            nodes[parent_id]["kwh"] = kwh
            nodes[parent_id]["kind"] = "root"
            continue

        child_id = node_id(child, child_meter)

        level_number = 1
        level_text = clean_text(row["Level"]).lower().replace("leve", "level")
        for token in level_text.split():
            if token.isdigit():
                level_number = int(token)
                break

        kind = "meter"
        if child_meter.lower() == "virtual meter":
            kind = "virtual"

        nodes[child_id] = {
            "id": child_id,
            "name": child,
            "meter": child_meter,
            "kwh": kwh,
            "level": level_number,
            "kind": kind,
        }

        edges.append((parent_id, child_id))
        children_map[parent_id].append(child_id)

    return nodes, edges, children_map


def calculate_positions(nodes, children_map):
    # Find root
    all_children = {child for children in children_map.values() for child in children}
    roots = [node_id for node_id in nodes if node_id not in all_children]
    root = roots[0] if roots else list(nodes.keys())[0]

    # Preferred top-level order
    if root in children_map:
        children = children_map[root]
        children_map[root] = sorted(
            children,
            key=lambda x: 0 if nodes[x]["name"].upper() == "TX2" else 1
        )

    x_spacing = 190
    y_spacing = 175
    leaf_counter = 0
    positions = {}

    def dfs(nid, depth=0):
        nonlocal leaf_counter

        children = children_map.get(nid, [])

        if not children:
            x = leaf_counter * x_spacing + 80
            leaf_counter += 1
        else:
            child_xs = [dfs(child, depth + 1) for child in children]
            x = sum(child_xs) / len(child_xs)

        y = depth * y_spacing + 40
        positions[nid] = {"x": x, "y": y}
        return x

    dfs(root)

    return positions


def downstream_total(node_id_value, nodes, children_map):
    return sum(nodes[child]["kwh"] for child in children_map.get(node_id_value, []))


def status_colour(parent_kwh, child_kwh):
    if not parent_kwh:
        return "OK"
    diff_pct = abs((parent_kwh - child_kwh) / parent_kwh)
    if diff_pct > 0.10:
        return "ALERT"
    if diff_pct > 0.05:
        return "CHECK"
    return "OK"


nodes, edges, children_map = load_data()
positions = calculate_positions(nodes, children_map)

cy_elements = []

for nid, n in nodes.items():
    meter_text = n["meter"]
    if meter_text.lower() == "virtual meter":
        meter_text = "Virtual Meter"

    label = f"{n['name']}\nMeter: {meter_text}\n{n['kwh']:,.2f} kWh"

    node_class = n["kind"]
    if n["name"].lower() in ["lv 2", "unit 18", "panel a", "panel b", "main incomer a", "main incomer b"]:
        node_class = "main"
    if n["name"].lower() in ["common area", "unit 5&6", "unit 5 & 6"]:
        node_class = "submain"

    cy_elements.append({
        "data": {
            "id": nid,
            "label": label,
            "name": n["name"],
            "meter": n["meter"],
            "kwh": n["kwh"],
        },
        "position": positions[nid],
        "classes": node_class,
    })

for source, target in edges:
    cy_elements.append({
        "data": {
            "source": source,
            "target": target,
        }
    })


balance_rows = []
for nid, n in nodes.items():
    if nid not in children_map:
        continue

    parent_kwh = n["kwh"]
    child_kwh = downstream_total(nid, nodes, children_map)
    diff = parent_kwh - child_kwh
    diff_pct = (diff / parent_kwh) if parent_kwh else 0
    status = status_colour(parent_kwh, child_kwh)

    balance_rows.append({
        "Parent": n["name"],
        "Meter": n["meter"],
        "Parent kWh": round(parent_kwh, 2),
        "Downstream kWh": round(child_kwh, 2),
        "Difference": round(diff, 2),
        "Difference %": f"{diff_pct:.1%}",
        "Status": status,
    })


app = dash.Dash(__name__)

app.layout = html.Div(
    style={"fontFamily": "Arial", "backgroundColor": "#eef3f8", "minHeight": "100vh"},
    children=[
        html.Div(
            style={
                "backgroundColor": "white",
                "padding": "14px 22px",
                "boxShadow": "0 2px 10px rgba(0,0,0,0.10)",
                "display": "flex",
                "justifyContent": "space-between",
                "alignItems": "center",
            },
            children=[
                html.Div([
                    html.H2("Eastleigh Energy Management Dashboard", style={"margin": "0"}),
                    html.Div(
                        "Built directly from Eastleigh_Input.xlsx — zoom, pan, drag, and refresh after Excel changes.",
                        style={"color": "#64748b", "fontSize": "13px"},
                    ),
                ]),
                html.Div("Local Demo", style={
                    "fontWeight": "bold",
                    "color": "#2563eb",
                    "border": "1px solid #cbd5e1",
                    "padding": "8px 12px",
                    "borderRadius": "10px",
                }),
            ],
        ),

        html.Div(
            style={
                "display": "grid",
                "gridTemplateColumns": "1fr 430px",
                "gap": "12px",
                "padding": "12px",
            },
            children=[
                html.Div(
                    style={
                        "backgroundColor": "white",
                        "borderRadius": "14px",
                        "boxShadow": "0 3px 14px rgba(0,0,0,0.08)",
                        "overflow": "hidden",
                    },
                    children=[
                        cyto.Cytoscape(
                            id="eastleigh-network",
                            elements=cy_elements,
                            layout={"name": "preset", "fit": True, "padding": 70},
                            userZoomingEnabled=True,
                            userPanningEnabled=True,
                            boxSelectionEnabled=True,
                            style={"width": "100%", "height": "78vh"},
                            stylesheet=[
                                {
                                    "selector": "node",
                                    "style": {
                                        "label": "data(label)",
                                        "text-wrap": "wrap",
                                        "text-max-width": "125px",
                                        "font-size": "10.5px",
                                        "font-weight": "bold",
                                        "text-valign": "center",
                                        "text-halign": "center",
                                        "background-color": "#ffffff",
                                        "border-width": 2,
                                        "border-color": "#f59e0b",
                                        "width": "145px",
                                        "height": "76px",
                                        "shape": "round-rectangle",
                                        "color": "#111827",
                                    },
                                },
                                {
                                    "selector": ".root, .virtual",
                                    "style": {
                                        "border-color": "#7c3aed",
                                        "background-color": "#f5f3ff",
                                    },
                                },
                                {
                                    "selector": ".main",
                                    "style": {
                                        "border-width": 3,
                                        "border-color": "#2563eb",
                                        "background-color": "#eff6ff",
                                    },
                                },
                                {
                                    "selector": ".submain",
                                    "style": {
                                        "border-width": 3,
                                        "border-color": "#16a34a",
                                        "background-color": "#f0fdf4",
                                    },
                                },
                                {
                                    "selector": "edge",
                                    "style": {
                                        "width": 2,
                                        "line-color": "#334155",
                                        "target-arrow-color": "#334155",
                                        "target-arrow-shape": "triangle",
                                        "curve-style": "taxi",
                                        "taxi-direction": "downward",
                                        "taxi-turn": "50%",
                                    },
                                },
                            ],
                        )
                    ],
                ),

                html.Div(
                    style={
                        "backgroundColor": "white",
                        "borderRadius": "14px",
                        "boxShadow": "0 3px 14px rgba(0,0,0,0.08)",
                        "padding": "12px",
                        "height": "78vh",
                        "overflowY": "auto",
                    },
                    children=[
                        html.H3("Parent vs Downstream Check", style={"marginTop": "0"}),
                        dash_table.DataTable(
                            data=balance_rows,
                            columns=[{"name": c, "id": c} for c in balance_rows[0].keys()] if balance_rows else [],
                            style_table={"overflowX": "auto"},
                            style_cell={
                                "fontSize": "11px",
                                "padding": "6px",
                                "textAlign": "left",
                                "whiteSpace": "normal",
                                "height": "auto",
                            },
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#f8fafc",
                            },
                            style_data_conditional=[
                                {"if": {"filter_query": "{Status} = 'ALERT'"}, "backgroundColor": "#fee2e2"},
                                {"if": {"filter_query": "{Status} = 'CHECK'"}, "backgroundColor": "#fef3c7"},
                                {"if": {"filter_query": "{Status} = 'OK'"}, "backgroundColor": "#dcfce7"},
                            ],
                        ),
                    ],
                ),
            ],
        ),
    ],
)


if __name__ == "__main__":
    print("Dashboard starting... open http://127.0.0.1:8050")
    import os

port = int(os.environ.get("PORT", 8050))

app.run(
    host="0.0.0.0",
    port=port,
    debug=False
)
