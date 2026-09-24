"""
03_dashboard_builder.py
=======================
Stage 3: Standalone Interactive HTML Dashboard Compiler.
Renders Plotly visualizations, dynamic KPI metric cards, and responsive tables
into a single self-contained HTML file using Jinja2.
Runs completely offline in any web browser without requiring a server.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any, Dict, List
import jinja2
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

INPUT_FILE = Path("output") / "carr_comparative_analysed.xlsx"
TEMPLATE_DIR = Path("templates")
TEMPLATE_FILE = "dashboard_template.html"
OUTPUT_HTML = Path("output") / "carr_comparative_dashboard.html"

COLOR_PALETTE = {
    "Navy": "#1e3a8a",
    "NavyDark": "#0f172a",
    "Gold": "#f59e0b",
    "Green": "#16a34a",
    "Red": "#dc2626",
    "Cyan": "#0891b2",
    "Purple": "#7c3aed",
    "Gray": "#64748b",
}


def log_msg(level: str, msg: str) -> None:
    """Print structured log messages with timestamps."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    icons = {"INFO": "[INFO]", "WARN": "[WARN]", "ERROR": "[FAIL]", "SUCCESS": "[ OK ]"}
    print(f"[{ts}] {icons.get(level, '[INFO]')} {msg}")


def format_inr(val: Any) -> str:
    """Format floating point numbers into Indian currency style with the INR prefix."""
    if val is None or pd.isna(val):
        return "₹ 0.00"
    try:
        val_float = float(val)
    except (ValueError, TypeError):
        return str(val)

    is_neg = val_float < 0
    val_abs = abs(val_float)

    s, sep, dec = f"{val_abs:.2f}".partition(".")
    if len(s) > 3:
        leading, last_three = s[:-3], s[-3:]
        pairs = []
        while len(leading) > 2:
            pairs.insert(0, leading[-2:])
            leading = leading[:-2]
        if leading:
            pairs.insert(0, leading)
        formatted = ",".join(pairs) + "," + last_three + "." + dec
    else:
        formatted = f"{val_abs:.2f}"

    return f"-₹ {formatted}" if is_neg else f"₹ {formatted}"


def build_macro_trend_chart(df_macro: pd.DataFrame) -> str:
    """Grouped bar and dual-axis line chart for receipts, expenditure, and utilization."""
    fig = go.Figure()
    years = df_macro["Financial_Year"].tolist()

    fig.add_trace(go.Bar(
        x=years,
        y=df_macro["Total_Receipts"],
        name="Available Receipts",
        marker_color=COLOR_PALETTE["Navy"]
    ))
    fig.add_trace(go.Bar(
        x=years,
        y=df_macro["Total_Expenditure"],
        name="Audited Expenditure",
        marker_color=COLOR_PALETTE["Green"]
    ))
    fig.add_trace(go.Scatter(
        x=years,
        y=df_macro["Fund_Utilization_Pct"],
        name="Utilization %",
        yaxis="y2",
        mode="lines+markers+text",
        text=[f"{v:.1f}%" for v in df_macro["Fund_Utilization_Pct"]],
        textposition="top center",
        line=dict(color=COLOR_PALETTE["Gold"], width=3)
    ))

    fig.update_layout(
        template="plotly_white",
        barmode="group",
        title="Multi-Year Macro Growth & Utilization Trajectory",
        font=dict(family="Segoe UI, Arial, sans-serif"),
        yaxis=dict(title="Amount (₹)"),
        yaxis2=dict(title="Utilization %", overlaying="y", side="right", range=[0, 100]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=40, r=40, t=60, b=40)
    )
    return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})


def build_scheme_multiline(df_ledger: pd.DataFrame) -> str:
    """Multi-line trajectory of normalized scheme expenditures across audit cycles."""
    df_scheme = df_ledger.groupby(
        ["Financial_Year", "Normalized_Bucket"]
    )["Payment_Expenditure"].sum().reset_index()

    fig = px.line(
        df_scheme,
        x="Financial_Year",
        y="Payment_Expenditure",
        color="Normalized_Bucket",
        markers=True,
        title="Scheme-Wise Expenditure Trajectories Across Audit Cycles",
        color_discrete_sequence=px.colors.qualitative.Safe
    )
    fig.update_layout(
        template="plotly_white",
        font=dict(family="Segoe UI, Arial, sans-serif"),
        margin=dict(l=40, r=40, t=50, b=40)
    )
    return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})


def build_parking_bubble(df_parking: pd.DataFrame) -> str:
    """Bubble scatter chart identifying severe fund parking outliers."""
    sample = df_parking.head(80).copy()
    fig = px.scatter(
        sample,
        x="Total_Receipt",
        y="Closing_Balance",
        size="Total_Receipt",
        color="Financial_Year",
        hover_data=["Block_Administrative_Centre", "Entity_Name", "Raw_Scheme", "Utilization_Pct"],
        title="Severe Fund Parking Outliers (>= ₹20 Lakhs, <= 20% Utilized)",
        color_discrete_sequence=px.colors.qualitative.Dark24
    )
    fig.update_layout(
        template="plotly_white",
        font=dict(family="Segoe UI, Arial, sans-serif"),
        margin=dict(l=40, r=40, t=50, b=40)
    )
    return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})


def main() -> None:
    log_msg("INFO", "Initializing Stage 3: Interactive Dashboard Compilation...")
    if not INPUT_FILE.exists():
        log_msg("ERROR", f"File '{INPUT_FILE}' not found. Run Stage 2 first.")
        return

    excel_data = pd.read_excel(INPUT_FILE, sheet_name=None)
    df_macro = excel_data.get("Macro_MultiYear_Summary", pd.DataFrame())
    df_yoy = excel_data.get("YoY_Expenditure_Matrix", pd.DataFrame())
    df_parking = excel_data.get("Severe_Fund_Parking", pd.DataFrame())
    df_ledger = excel_data.get("Enriched_Ledger", pd.DataFrame())

    log_msg("INFO", "Rendering Plotly charts...")
    chart_macro = build_macro_trend_chart(df_macro)
    chart_scheme = build_scheme_multiline(df_ledger)
    chart_parking = build_parking_bubble(df_parking)

    latest = df_macro.iloc[-1]
    kpis = {
        "years_span": f"{df_macro['Financial_Year'].iloc[0]} to {latest['Financial_Year']}",
        "latest_receipts": format_inr(latest["Total_Receipts"]),
        "latest_expenditure": format_inr(latest["Total_Expenditure"]),
        "latest_cb": format_inr(latest["Closing_Unspent_Balance"]),
        "latest_util_pct": f"{latest['Fund_Utilization_Pct']:.1f}",
        "parking_count": len(df_parking),
    }

    macro_display = df_macro.to_dict("records")
    for r in macro_display:
        r["Total_Receipts"] = format_inr(r["Total_Receipts"])
        r["Total_Expenditure"] = format_inr(r["Total_Expenditure"])
        r["Closing_Unspent_Balance"] = format_inr(r["Closing_Unspent_Balance"])

    parking_display = df_parking.head(100).to_dict("records")
    for r in parking_display:
        r["Total_Receipt"] = format_inr(r["Total_Receipt"])
        r["Payment_Expenditure"] = format_inr(r["Payment_Expenditure"])
        r["Closing_Balance"] = format_inr(r["Closing_Balance"])
        r["Utilization_Pct"] = f"{r['Utilization_Pct']:.1f}"

    yoy_sample = df_yoy.head(100).copy()
    yoy_columns = list(yoy_sample.columns)
    yoy_rows = yoy_sample.to_dict("records")
    for r in yoy_rows:
        for c in yoy_columns:
            if isinstance(r[c], (int, float)):
                if "Pct" in c or "CAGR" in c or "%" in c:
                    r[c] = f"{r[c]:.1f}%" if not pd.isna(r[c]) else "-"
                else:
                    r[c] = format_inr(r[c])

    env = jinja2.Environment(
        loader=jinja2.FileSystemLoader(str(TEMPLATE_DIR)),
        autoescape=jinja2.select_autoescape(["html", "xml"])
    )
    template = env.get_template(TEMPLATE_FILE)

    rendered_html = template.render(
        kpis=kpis,
        chart_macro=chart_macro,
        chart_scheme=chart_scheme,
        chart_parking=chart_parking,
        macro_table=macro_display,
        parking_table=parking_display,
        yoy_columns=yoy_columns,
        yoy_rows=yoy_rows,
        generated_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    )

    OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(rendered_html)

    log_msg("SUCCESS", f"Dashboard HTML generated successfully: {OUTPUT_HTML.resolve()}")


if __name__ == "__main__":
    main()