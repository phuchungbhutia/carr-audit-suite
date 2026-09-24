"""
02_comparative_analytics.py
===========================
Stage 2: Cross-Year Comparative Analytics & Anomaly Detection.
"""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any, Dict, List
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

INPUT_MASTER = Path("output") / "carr_multiyear_master.xlsx"
OUTPUT_ANALYSED = Path("output") / "carr_comparative_analysed.xlsx"


def log_msg(level: str, msg: str) -> None:
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    icons = {"INFO": "[INFO]", "WARN": "[WARN]", "ERROR": "[FAIL]", "SUCCESS": "[ OK ]"}
    print(f"[{ts}] {icons.get(level, '[INFO]')} {msg}")


def compute_cagr(start_val: float, end_val: float, periods: int) -> float:
    if periods <= 0 or start_val <= 0 or end_val <= 0:
        return np.nan
    try:
        return (end_val / start_val) ** (1.0 / periods) - 1.0
    except (ZeroDivisionError, OverflowError, ValueError):
        return np.nan


def perform_comparative_analysis(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    df = df.copy()

    df["Utilization_Pct"] = np.where(
        df["Total_Receipt"] > 0,
        (df["Payment_Expenditure"] / df["Total_Receipt"]) * 100.0,
        0.0,
    ).round(2)

    def classify_risk(row: pd.Series) -> str:
        if row["Closing_Balance"] < 0:
            return "CRITICAL (Negative Balance)"
        if row["Total_Receipt"] >= 2_000_000 and row["Utilization_Pct"] <= 20.0:
            return "HIGH RISK (Severe Fund Parking)"
        if row["Utilization_Pct"] < 40.0:
            return "MODERATE RISK (Slow Progress)"
        return "NORMAL"

    df["Audit_Risk_Flag"] = df.apply(classify_risk, axis=1)

    # 1. Macro Summary
    macro_records = []
    for fy in sorted(df["Financial_Year"].unique()):
        sub = df[df["Financial_Year"] == fy]
        rec = sub["Total_Receipt"].sum()
        pay = sub["Payment_Expenditure"].sum()
        cb = sub["Closing_Balance"].sum()
        util = (pay / rec * 100.0) if rec > 0 else 0.0

        severe_parking = sub[
            (sub["Total_Receipt"] >= 2_000_000) & (sub["Utilization_Pct"] <= 20.0)
        ]

        macro_records.append({
            "Financial_Year": fy,
            "Total_Receipts": rec,
            "Total_Expenditure": pay,
            "Closing_Unspent_Balance": cb,
            "Fund_Utilization_Pct": round(util, 2),
            "Severe_Fund_Parking_Count": len(severe_parking),
            "Total_Audit_Units": sub["Entity_Name"].nunique(),
        })

    df_macro = pd.DataFrame(macro_records)

    # 2. YoY Matrix
    pivot_exp = pd.pivot_table(
        df,
        values="Payment_Expenditure",
        index=["Block_Administrative_Centre", "Entity_Name", "Normalized_Bucket"],
        columns="Financial_Year",
        aggfunc="sum",
    ).fillna(0.0)

    years = sorted(list(pivot_exp.columns))
    df_yoy = pivot_exp.copy()

    for i in range(len(years) - 1):
        y_prev, y_curr = years[i], years[i + 1]
        delta_col = f"Delta_{y_prev}_to_{y_curr}"
        pct_col = f"PctChange_{y_prev}_to_{y_curr}"
        df_yoy[delta_col] = df_yoy[y_curr] - df_yoy[y_prev]
        df_yoy[pct_col] = np.where(
            df_yoy[y_prev] != 0,
            (df_yoy[delta_col] / df_yoy[y_prev].abs()) * 100.0,
            np.nan,
        )

    if len(years) >= 2:
        span = len(years) - 1
        df_yoy["CAGR_Pct"] = [
            compute_cagr(start, end, span) * 100.0
            for start, end in zip(df_yoy[years[0]], df_yoy[years[-1]])
        ]

    # 3. Severe Parking Filter
    df_parking = df[
        (df["Total_Receipt"] >= 2_000_000) & (df["Utilization_Pct"] <= 20.0)
    ].sort_values("Closing_Balance", ascending=False)

    return {
        "Macro_Summary": df_macro,
        "YoY_Matrix": df_yoy.reset_index(),
        "Fund_Parking": df_parking,
        "Enriched_Ledger": df,
    }


def format_excel(filepath: Path) -> None:
    wb = openpyxl.load_workbook(filepath)
    hdr_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    hdr_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    thin_border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB"),
    )

    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            hdr_val = str(ws.cell(row=1, column=col[0].column).value or "")

            for cell in col:
                cell.border = thin_border
                if cell.row > 1:
                    cell.font = data_font
                    if isinstance(cell.value, (int, float)):
                        if "Pct" in hdr_val or "%" in hdr_val:
                            cell.number_format = "0.00%"
                        else:
                            cell.number_format = "#,##,##0.00"
                            cell.alignment = Alignment(horizontal="right")
                val_len = len(str(cell.value or ""))
                if val_len > max_len:
                    max_len = val_len
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

    wb.save(filepath)


def main() -> None:
    log_msg("INFO", "Initializing Stage 2: Comparative Financial Analytics...")
    if not INPUT_MASTER.exists():
        log_msg("ERROR", f"File '{INPUT_MASTER}' not found. Run Stage 1 first.")
        return

    excel_file = pd.ExcelFile(INPUT_MASTER)
    if "GPU_Master_Ledger" not in excel_file.sheet_names:
        log_msg("ERROR", "Worksheet 'GPU_Master_Ledger' was not generated by Stage 1. Rerun Stage 1 extraction.")
        return

    df_gpu = pd.read_excel(INPUT_MASTER, sheet_name="GPU_Master_Ledger")
    results = perform_comparative_analysis(df_gpu)

    OUTPUT_ANALYSED.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_ANALYSED, engine="openpyxl") as writer:
        results["Macro_Summary"].to_excel(writer, sheet_name="Macro_MultiYear_Summary", index=False)
        results["YoY_Matrix"].to_excel(writer, sheet_name="YoY_Expenditure_Matrix", index=False)
        results["Fund_Parking"].to_excel(writer, sheet_name="Severe_Fund_Parking", index=False)
        results["Enriched_Ledger"].to_excel(writer, sheet_name="Enriched_Ledger", index=False)

    format_excel(OUTPUT_ANALYSED)
    log_msg("SUCCESS", f"Comparative analytics complete. Saved to: {OUTPUT_ANALYSED.resolve()}")


if __name__ == "__main__":
    main()