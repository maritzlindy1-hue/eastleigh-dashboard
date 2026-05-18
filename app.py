
from pathlib import Path
from collections import defaultdict, deque

import pandas as pd
import dash
from dash import html, dash_table, dcc, Input, Output
import dash_cytoscape as cyto
import plotly.express as px


EXCEL_FILE = Path(__file__).with_name("Eastleigh_Masterlist.xlsx")
SHEET_NAME = 0

BALANCE_OK_PERCENT = 5       # 0% to 5% = balanced / green
BALANCE_WARN_PERCENT = 10    # 5% to 10% = warning / amber
# Above 10% = mismatch / red


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_meter(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def clean_kwh(value):
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def parse_level(value):
    text = clean_text(value).lower().replace("leve", "level")
    for token in text.replace("-", " ").split():
        if token.isdigit():
            return int(token)
    return 1


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
    parent_map = {}

    for _, row in df.iterrows():
        parent_name = clean_text(row["Parent"])
        child_name = clean_text(row["Child"])
        parent_meter = clean_meter(row["Parent Meter Serial Nr"])
        child_meter = clean_meter(row["Child Meter Serial Nr"])
        level = parse_level(row["Level"])
        kwh = clean_kwh(row["kWh"])

        if not parent_name:
            continue

        parent_id = node_id(parent_name, parent_meter)

        if parent_id not in nodes:
            nodes[parent_id] = {
                "id": parent_id,
                "name": parent_name,
                "meter": parent_meter,
                "kwh": 0.0,
                "level": max(level - 1, 1),
                "kind": "virtual" if parent_meter.lower() == "virtual meter" else "meter",
            }

        # If no child, this row is the top/root value
        if not child_name:
            nodes[parent_id]["kwh"] = kwh
            nodes[parent_id]["kind"] = "root"
            nodes[parent_id]["level"] = level
            continue

        child_id = node_id(child_name, child_meter)
        child_kind = "virtual" if child_meter.lower() == "virtual meter" else "meter"

        nodes[child_id] = {
            "id": child_id,
            "name": child_name,
            "meter": child_meter,
            "kwh": kwh,
            "level": level,
            "kind": child_kind,
        }

        edges.append((parent_id, child_id))
        children_map[parent_id].append(child_id)
        parent_map[child_id] = parent_id

    return nodes, edges, children_map, parent_map


def calc_balance(nid, nodes, children_map):
    parent_kwh = nodes[nid]["kwh"]
    child_total = sum(nodes[c]["kwh"] for c in children_map.get(nid, []))
    diff = parent_kwh - child_total
    diff_pct = (diff / parent_kwh * 100) if parent_kwh else 0.0
    abs_pct = abs(diff_pct)

    if nid not in children_map:
        status = "NO CHILDREN"
    elif abs_pct <= BALANCE_OK_PERCENT:
        status = "BALANCED"
    elif abs_pct <= BALANCE_WARN_PERCENT:
        status = "CHECK"
    else:
        status = "MISMATCH"

    return child_total, diff, diff_pct, status


def find_root(nodes, parent_map):
    roots = [nid for nid in nodes if nid not in parent_map]
    # Prefer HV Supply if found
    for root in roots:
        if "hv" in nodes[root]["name"].lower():
            return root
    return roots[0] if roots else list(nodes.keys())[0]


def order_children(children, nodes):
    preferred = {
        "tx2": 0,
        "tx1": 1,
        "lv 2": 0,
        "unit 18": 1,
        "panel a": 0,
        "panel b": 1,
        "common area": 50,
        "unit 5&6": 50,
        "unit 5 & 6": 50,
    }
    return sorted(children, key=lambda c: (preferred.get(nodes[c]["name"].lower(), 20), nodes[c]["name"]))


def calculate_positions(nodes, children_map, parent_map):
    root = find_root(nodes, parent_map)

    # Use tree layout with fixed coordinates so it resembles a single line diagram.
    x_spacing = 175
    y_spacing = 165
    positions = {}
    leaf_counter = 0

    def dfs(nid, depth=0):
        nonlocal leaf_counter
        children = order_children(children_map.get(nid, []), nodes)

        if not children:
            x = leaf_counter * x_spacing + 70
            leaf_counter += 1
        else:
            child_xs = [dfs(child, depth + 1) for child in children]
            x = sum(child_xs) / len(child_xs)

        positions[nid] = {"x": x, "y": depth * y_spacing + 35}
        return x

    dfs(root)

    return positions


def class_for_node(nid, n, children_map, balance_status):
    name = n["name"].lower()

    if n["kind"] in ["root", "virtual"] or "tx" in name or "transformer" in name:
        base = "virtual"
    elif nid in children_map:
        base = "parent"
    else:
        base = "child"

    if name in ["common area", "unit 5&6", "unit 5 & 6"]:
        base = "subparent"

    if balance_status == "MISMATCH":
        return f"{base} mismatch"
    if balance_status == "CHECK":
        return f"{base} check"
    if balance_status == "BALANCED":
        return f"{base} balanced"

    return base


nodes, edges, children_map, parent_map = load_data()
positions = calculate_positions(nodes, children_map, parent_map)

cy_elements = []
balance_rows = []

for nid, n in nodes.items():
    child_total, diff, diff_pct, status = calc_balance(nid, nodes, children_map)

    meter_text = n["meter"]
    if meter_text.lower() == "virtual meter":
        meter_text = "Virtual Meter"

    if nid in children_map:
        label = (
            f"{n['name']}\n"
            f"Meter: {meter_text}\n"
            f"{n['kwh']:,.2f} kWh\n"
            f"Child: {child_total:,.2f} kWh\n"
            f"Diff: {diff_pct:+.1f}%"
        )
    else:
        label = (
            f"{n['name']}\n"
            f"Meter: {meter_text}\n"
            f"{n['kwh']:,.2f} kWh"
        )

    cy_elements.append({
        "data": {
            "id": nid,
            "label": label,
            "name": n["name"],
            "meter": meter_text,
            "kwh": n["kwh"],
            "child_total": child_total,
            "diff": diff,
            "diff_pct": round(diff_pct, 2),
            "status": status,
        },
        "position": positions.get(nid, {"x": 0, "y": 0}),
        "classes": class_for_node(nid, n, children_map, status),
    })

    if nid in children_map:
        balance_rows.append({
            "Parent": n["name"],
            "Meter": meter_text,
            "Parent kWh": round(n["kwh"], 2),
            "Child Total kWh": round(child_total, 2),
            "Difference kWh": round(diff, 2),
            "Difference %": f"{diff_pct:+.1f}%",
            "Status": status,
        })

for source, target in edges:
    cy_elements.append({"data": {"source": source, "target": target}})


parent_options = [
    {"label": f"{nodes[nid]['name']} ({nodes[nid]['meter']})", "value": nid}
    for nid in nodes
    if nid in children_map
]

default_parent = parent_options[0]["value"] if parent_options else None


def make_pie(parent_id):
    if not parent_id or parent_id not in children_map:
        return px.pie(title="No parent selected")

    child_ids = children_map[parent_id]
    pie_df = pd.DataFrame([
        {
            "Meter": nodes[c]["name"],
            "Meter No": nodes[c]["meter"],
            "kWh": nodes[c]["kwh"],
        }
        for c in child_ids
    ])

    fig = px.pie(
        pie_df,
        names="Meter",
        values="kWh",
        title=f"Contribution below: {nodes[parent_id]['name']}",
        hole=0.35,
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    fig.update_layout(
        margin=dict(l=10, r=10, t=55, b=10),
        legend=dict(orientation="h", y=-0.15),
        height=390,
    )
    return fig


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
                        "Single-line balancing view + simple contribution view for non-technical users.",
                        style={"color": "#64748b", "fontSize": "13px"},
                    ),
                ]),
                html.Div("Excel-driven EMS Demo", style={
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
                "gridTemplateColumns": "1fr 460px",
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
                        html.Div(
                            style={"padding": "10px 14px", "fontSize": "13px", "color": "#475569"},
                            children=[
                                html.B("Colour rule: "),
                                "Green = balanced within ±5%, Amber = check ±5–10%, Red = mismatch above ±10%.",
                            ],
                        ),
                        cyto.Cytoscape(
                            id="eastleigh-network",
                            elements=cy_elements,
                            layout={"name": "preset", "fit": True, "padding": 70},
                            userZoomingEnabled=True,
                            userPanningEnabled=True,
                            boxSelectionEnabled=True,
                            style={"width": "100%", "height": "72vh"},
                            stylesheet=[
                                {
                                    "selector": "node",
                                    "style": {
                                        "label": "data(label)",
                                        "text-wrap": "wrap",
                                        "text-max-width": "132px",
                                        "font-size": "10px",
                                        "font-weight": "bold",
                                        "text-valign": "center",
                                        "text-halign": "center",
                                        "background-color": "#ffffff",
                                        "border-width": 2,
                                        "border-color": "#f59e0b",
                                        "width": "150px",
                                        "height": "92px",
                                        "shape": "round-rectangle",
                                        "color": "#111827",
                                    },
                                },
                                {
                                    "selector": ".virtual",
                                    "style": {
                                        "border-width": 3,
                                        "border-color": "#7c3aed",
                                        "background-color": "#f5f3ff",
                                    },
                                },
                                {
                                    "selector": ".parent",
                                    "style": {
                                        "border-width": 3,
                                        "border-color": "#2563eb",
                                        "background-color": "#eff6ff",
                                    },
                                },
                                {
                                    "selector": ".subparent",
                                    "style": {
                                        "border-width": 3,
                                        "border-color": "#16a34a",
                                        "background-color": "#f0fdf4",
                                    },
                                },
                                {
                                    "selector": ".balanced",
                                    "style": {
                                        "background-color": "#dcfce7",
                                        "border-color": "#16a34a",
                                    },
                                },
                                {
                                    "selector": ".check",
                                    "style": {
                                        "background-color": "#fef3c7",
                                        "border-color": "#f59e0b",
                                        "border-width": 4,
                                    },
                                },
                                {
                                    "selector": ".mismatch",
                                    "style": {
                                        "background-color": "#fee2e2",
                                        "border-color": "#dc2626",
                                        "border-width": 4,
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
                        ),
                    ],
                ),

                html.Div(
                    style={
                        "display": "grid",
                        "gridTemplateRows": "auto 1fr",
                        "gap": "12px",
                    },
                    children=[
                        html.Div(
                            style={
                                "backgroundColor": "white",
                                "borderRadius": "14px",
                                "boxShadow": "0 3px 14px rgba(0,0,0,0.08)",
                                "padding": "12px",
                            },
                            children=[
                                html.H3("Contribution Chart", style={"marginTop": 0}),
                                html.Div("Choose a parent meter:", style={"fontSize": "12px", "color": "#64748b"}),
                                dcc.Dropdown(
                                    id="parent-dropdown",
                                    options=parent_options,
                                    value=default_parent,
                                    clearable=False,
                                    style={"marginTop": "6px"},
                                ),
                                dcc.Graph(id="contribution-pie", figure=make_pie(default_parent)),
                            ],
                        ),

                        html.Div(
                            style={
                                "backgroundColor": "white",
                                "borderRadius": "14px",
                                "boxShadow": "0 3px 14px rgba(0,0,0,0.08)",
                                "padding": "12px",
                                "overflowY": "auto",
                            },
                            children=[
                                html.H3("Balance Check", style={"marginTop": 0}),
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
                                        {"if": {"filter_query": "{Status} = 'MISMATCH'"}, "backgroundColor": "#fee2e2", "color": "#991b1b", "fontWeight": "bold"},
                                        {"if": {"filter_query": "{Status} = 'CHECK'"}, "backgroundColor": "#fef3c7", "color": "#92400e"},
                                        {"if": {"filter_query": "{Status} = 'BALANCED'"}, "backgroundColor": "#dcfce7", "color": "#166534"},
                                    ],
                                    page_size=12,
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        ),
    ],
)


@app.callback(
    Output("contribution-pie", "figure"),
    Input("parent-dropdown", "value"),
)
def update_pie(parent_id):
    return make_pie(parent_id)


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 8050))
    print(f"Dashboard starting... open http://127.0.0.1:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
