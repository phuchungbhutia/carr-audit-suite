"""
01_extract_universal.py
=======================
Fully Generic Multi-Year Adaptive Ingestion & Normalization Engine.
- Automatically discovers all PDF reports inside input_pdfs/
- Extracts Financial Year (FY) from metadata, text, or filenames
- Self-discovers Appendix II and III page boundaries dynamically
- Parses 6-column, 7-column, and 8-column layouts using tail-first alignment
"""

from __future__ import annotations

import datetime
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import pdfplumber

INPUT_DIR = Path("input_pdfs")
OUTPUT_DIR = Path("output")
OUTPUT_MASTER = OUTPUT_DIR / "carr_multiyear_master.xlsx"

DISTRICT_NAMES = [
    "GANGTOK", "PAKYONG", "GYALSHING", "MANGAN",
    "NAMCHI", "SORENG", "EAST", "WEST", "NORTH", "SOUTH"
]


def log_msg(level: str, msg: str) -> None:
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    icons = {"INFO": "[INFO]", "WARN": "[WARN]", "ERROR": "[FAIL]", "SUCCESS": "[ OK ]"}
    print(f"[{ts}] {icons.get(level, '[INFO]')} {msg}")


def parse_indian_amount(val: Any) -> float:
    """Robustly parse strings with Indian commas, parentheses, nil signs, and decimals."""
    if val is None or pd.isna(val):
        return 0.0
    s = str(val).strip()
    if not s or s.lower() in ["-", "--", "---", "nil", "null", "none", "n/a", ".."]:
        return 0.0

    is_negative = False
    if s.startswith("(") and s.endswith(")"):
        is_negative = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        is_negative = True
        s = s[1:].strip()

    s_clean = re.sub(r"[Rs₹$\s,]", "", s)
    try:
        f_val = float(s_clean)
        return -f_val if is_negative else f_val
    except ValueError:
        return 0.0


def normalize_scheme(raw_scheme: str) -> str:
    """Classify scheme text into unified audit buckets."""
    u = str(raw_scheme).upper()
    if "15" in u and ("FC" in u or "FINANCE" in u):
        return "15th Finance Commission"
    if "14" in u and ("FC" in u or "FINANCE" in u):
        return "14th Finance Commission"
    if "SFC" in u or "STATE FINANCE" in u:
        return "State Finance Commission"
    if "OSR" in u or "REVENUE" in u or "TRADE" in u or "LICENSE" in u:
        return "Own Source Revenue"
    if "GIA" in u or "SALARY" in u or "HONORARIUM" in u:
        return "Grant-In-Aid / Salaries"
    if "SBM" in u or "SWACHH" in u:
        return "Swachh Bharat Mission"
    if "JJM" in u:
        return "Jal Jeevan Mission"
    if "EFF" in u or "EMPOWERED" in u:
        return "Empowered Functionaries Fund"
    return "Other / Miscellaneous"


def detect_financial_year(pdf: pdfplumber.PDF, file_path: Path) -> str:
    """
    Autonomously extract the reporting Financial Year:
    1. Checks the first 5 pages for explicit 'FINANCIAL YEAR 20XX-XX' phrases.
    2. Fallbacks to filename regex patterns.
    """
    for page in pdf.pages[:5]:
        text = (page.extract_text() or "").upper()
        # Look for: FINANCIAL YEAR 2024-25, FOR THE YEAR 2023-2024, etc.
        match = re.search(r"(?:FINANCIAL\s+YEAR|YEAR|F\.Y\.)\s*[:\s]*((?:20\d{2})[-_–/](?:\d{2,4}))", text)
        if match:
            raw_fy = match.group(1).replace("_", "-").replace("–", "-")
            parts = raw_fy.split("-")
            if len(parts[1]) == 4:
                return f"{parts[0]}-{parts[1][2:]}"
            return raw_fy

    # Filename fallback regex
    stem = file_path.stem.upper()
    match_file = re.search(r"(20\d{2})[-_](20)?(\d{2})", stem)
    if match_file:
        return f"{match_file.group(1)}-{match_file.group(3)}"

    match_year = re.search(r"20(\d{2})", stem)
    if match_year:
        end_year = int(f"20{match_year.group(1)}")
        return f"{end_year - 1}-{str(end_year)[2:]}"

    return "Unknown-FY"


def discover_appendix_boundaries_dynamically(pdf: pdfplumber.PDF) -> Dict[str, Tuple[int, int]]:
    """
    Self-discovers start and end pages for Appendix II and III without hardcoding:
    - Skips initial Table of Contents pages.
    - Scans for distinct appendix title occurrences.
    - Calculates page spans from the distance between consecutive section starts.
    """
    total_pages = len(pdf.pages)
    # The Table of Contents is always in the first 15% of pages
    scan_start = max(5, int(total_pages * 0.15))

    app2_start: Optional[int] = None
    app3_start: Optional[int] = None
    app4_start: Optional[int] = None

    for idx in range(scan_start, total_pages):
        p_num = idx + 1
        page_text = (pdf.pages[idx].extract_text() or "").upper()

        if "APPENDIX" in page_text:
            if ("APPENDIX- II" in page_text or "APPENDIX - II" in page_text or "APPENDIX II" in page_text):
                if any(k in page_text for k in ["FINANCIAL POSITION", "GRAM PANCHAYAT"]):
                    if app2_start is None:
                        app2_start = p_num

            elif ("APPENDIX- III" in page_text or "APPENDIX - III" in page_text or "APPENDIX III" in page_text):
                if any(k in page_text for k in ["FINANCIAL POSITION", "ZILLA PANCHAYAT"]):
                    if app3_start is None:
                        app3_start = p_num

            elif ("APPENDIX- IV" in page_text or "APPENDIX - IV" in page_text or "APPENDIX IV" in page_text):
                if any(k in page_text for k in ["FINANCIAL POSITION", "NAGAR PANCHAYAT", "MUNICIPAL"]):
                    if app4_start is None:
                        app4_start = p_num
                        break

    # Robust fallback boundaries if delimiters aren't explicit
    if app2_start is None:
        app2_start = int(total_pages * 0.65)
    if app3_start is None:
        app3_start = app2_start + 20
    if app4_start is None:
        app4_start = total_pages

    return {
        "App_II": (app2_start, app3_start - 1),
        "App_III": (app3_start, app4_start - 1),
    }


def extract_gpu_records(
    pdf: pdfplumber.PDF, start_page: int, end_page: int, fy: str, filename: str
) -> List[Dict[str, Any]]:
    """Generic table extractor that handles any column layout via tail-first indexing."""
    records: List[Dict[str, Any]] = []
    current_bac = "General"
    current_gpu = "General"

    for p in range(start_page, end_page + 1):
        if p > len(pdf.pages):
            break
        page = pdf.pages[p - 1]
        tables = page.extract_tables()
        for tbl in tables:
            for row in tbl:
                row_cells = [str(c).strip() if c is not None else "" for c in row]
                if len(row_cells) < 4:
                    continue

                full_txt = " ".join(row_cells).upper()
                if "OPENING BALANCE" in full_txt or "GRAND TOTAL" in full_txt:
                    continue

                # Filter aggregate subtotal rows
                if "TOTAL" in full_txt and not any(
                    x in full_txt for x in ["RECEIPT", "PAYMENT", "INTEREST"]
                ):
                    continue

                # Collect numbers starting from the row tail
                trailing_nums: List[Tuple[int, float]] = []
                for i in range(len(row_cells) - 1, -1, -1):
                    val = row_cells[i]
                    if val != "" and (re.search(r"\d", val) or val in ["-", "NIL", "nil"]):
                        parsed = parse_indian_amount(val)
                        trailing_nums.insert(0, (i, parsed))
                    else:
                        break

                if len(trailing_nums) < 3:
                    continue

                # Tail-first column mapping
                if len(trailing_nums) >= 5:
                    ob = trailing_nums[-5][1]
                    tot_rec = trailing_nums[-3][1]
                    pay = trailing_nums[-2][1]
                    cb = trailing_nums[-1][1]
                    text_cutoff = trailing_nums[-5][0]
                elif len(trailing_nums) == 4:
                    ob = trailing_nums[-4][1]
                    tot_rec = trailing_nums[-3][1]
                    pay = trailing_nums[-2][1]
                    cb = trailing_nums[-1][1]
                    text_cutoff = trailing_nums[-4][0]
                else:
                    tot_rec = trailing_nums[-3][1]
                    pay = trailing_nums[-2][1]
                    cb = trailing_nums[-1][1]
                    ob = 0.0
                    text_cutoff = trailing_nums[-3][0]

                text_tokens = [c for c in row_cells[:text_cutoff] if c]
                if not text_tokens:
                    continue

                if len(text_tokens) >= 3:
                    current_bac = text_tokens[0]
                    current_gpu = text_tokens[1]
                    raw_scheme = " ".join(text_tokens[2:])
                elif len(text_tokens) == 2:
                    if any(
                        s in text_tokens[1].upper()
                        for s in ["FC", "SFC", "OSR", "GIA", "JJM", "SCHEME", "TOTAL"]
                    ):
                        current_gpu = text_tokens[0]
                        raw_scheme = text_tokens[1]
                    else:
                        current_bac = text_tokens[0]
                        current_gpu = text_tokens[1]
                        raw_scheme = "General"
                else:
                    raw_scheme = text_tokens[0]

                if tot_rec > 0 or pay > 0 or cb > 0 or ob > 0:
                    records.append({
                        "Financial_Year": fy,
                        "Tier": "Gram Panchayat Unit",
                        "Block_Administrative_Centre": current_bac.strip(),
                        "Entity_Name": current_gpu.strip(),
                        "Raw_Scheme": raw_scheme.strip(),
                        "Normalized_Bucket": normalize_scheme(raw_scheme),
                        "Opening_Balance": ob,
                        "Total_Receipt": tot_rec,
                        "Payment_Expenditure": pay,
                        "Closing_Balance": cb,
                        "Source_File": filename,
                        "Page_Number": p,
                    })

    return records


def extract_zp_records(
    pdf: pdfplumber.PDF, start_page: int, end_page: int, fy: str, filename: str
) -> List[Dict[str, Any]]:
    """Extract ZP rows invariant of district count or column headers."""
    records: List[Dict[str, Any]] = []
    current_zp = "District Zilla Panchayat"

    for p in range(start_page, end_page + 1):
        if p > len(pdf.pages):
            break
        page = pdf.pages[p - 1]
        tables = page.extract_tables()
        for tbl in tables:
            for row in tbl:
                row_cells = [str(c).strip() if c is not None else "" for c in row]
                if len(row_cells) < 4:
                    continue

                full_txt = " ".join(row_cells).upper()
                if "OPENING BALANCE" in full_txt or "GRAND TOTAL" in full_txt:
                    continue
                if "TOTAL" in row_cells[0].upper() or (len(row_cells) > 1 and "TOTAL" in row_cells[1].upper()):
                    continue

                nums = [parse_indian_amount(c) for c in row_cells if re.search(r"\d", c)]
                if len(nums) < 3:
                    continue

                for dist in DISTRICT_NAMES:
                    if dist in full_txt:
                        current_zp = f"{dist.title()} Zilla Panchayat"
                        break

                cb = nums[-1]
                pay = nums[-2]
                tot_rec = nums[-3]
                ob = nums[-4] if len(nums) >= 4 else 0.0
                raw_scheme = row_cells[1] if len(row_cells) > 1 and not re.search(r"\d", row_cells[1]) else "General"

                records.append({
                    "Financial_Year": fy,
                    "Tier": "Zilla Panchayat",
                    "Block_Administrative_Centre": "District Level",
                    "Entity_Name": current_zp,
                    "Raw_Scheme": raw_scheme,
                    "Normalized_Bucket": normalize_scheme(raw_scheme),
                    "Opening_Balance": ob,
                    "Total_Receipt": tot_rec,
                    "Payment_Expenditure": pay,
                    "Closing_Balance": cb,
                    "Source_File": filename,
                    "Page_Number": p,
                })

    return records


def format_excel(filepath: Path) -> None:
    """Format workbook sheets with headers, frozen rows, and Indian numbering masks."""
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
            for cell in col:
                cell.border = thin_border
                if cell.row > 1:
                    cell.font = data_font
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##,##0.00"
                        cell.alignment = Alignment(horizontal="right")
                val_len = len(str(cell.value or ""))
                if val_len > max_len:
                    max_len = val_len
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

    wb.save(filepath)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    all_gpu: List[Dict[str, Any]] = []
    all_zp: List[Dict[str, Any]] = []

    log_msg("INFO", "Initializing Generic Multi-Year PDF Discovery...")

    # Automatically discover all PDF files inside input_pdfs/
    pdf_files = sorted(list(INPUT_DIR.glob("*.pdf")))
    if not pdf_files:
        log_msg("ERROR", f"No PDF files detected in '{INPUT_DIR.resolve()}'.")
        return

    for pdf_path in pdf_files:
        log_msg("INFO", f"Inspecting '{pdf_path.name}'...")
        with pdfplumber.open(pdf_path) as pdf:
            fy = detect_financial_year(pdf, pdf_path)
            log_msg("INFO", f"  • Discovered Financial Year: {fy}")

            bounds = discover_appendix_boundaries_dynamically(pdf)
            g_start, g_end = bounds["App_II"]
            z_start, z_end = bounds["App_III"]

            log_msg("INFO", f"  • Dynamic Appendix II Range: Pages {g_start} to {g_end}")
            extracted_gpus = extract_gpu_records(pdf, g_start, g_end, fy, pdf_path.name)
            all_gpu.extend(extracted_gpus)
            log_msg("INFO", f"  • Extracted {len(extracted_gpus)} GPU records.")

            log_msg("INFO", f"  • Dynamic Appendix III Range: Pages {z_start} to {z_end}")
            extracted_zps = extract_zp_records(pdf, z_start, z_end, fy, pdf_path.name)
            all_zp.extend(extracted_zps)
            log_msg("INFO", f"  • Extracted {len(extracted_zps)} ZP records.")

    if not all_gpu and not all_zp:
        log_msg("ERROR", "No records could be extracted from any PDF.")
        return

    df_gpu = pd.DataFrame(all_gpu)
    df_zp = pd.DataFrame(all_zp)

    with pd.ExcelWriter(OUTPUT_MASTER, engine="openpyxl") as writer:
        if not df_gpu.empty:
            df_gpu.to_excel(writer, sheet_name="GPU_Master_Ledger", index=False)
        if not df_zp.empty:
            df_zp.to_excel(writer, sheet_name="ZP_Master_Ledger", index=False)

    format_excel(OUTPUT_MASTER)
    log_msg("SUCCESS", f"Generic extraction complete. Saved {len(df_gpu)} total records to: {OUTPUT_MASTER.resolve()}")


if __name__ == "__main__":
    main()