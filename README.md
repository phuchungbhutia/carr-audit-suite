# 🏛️ Multi-Year Audit Reporting & Analytics Suite

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.9%2B-blue?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.9+"/>
  <img src="https://img.shields.io/badge/Status-Production%20Ready-success?style=for-the-badge" alt="Status"/>
  <img src="https://img.shields.io/badge/Coverage-100%25%20Audited%20Units-green?style=for-the-badge" alt="Coverage"/>
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey?style=for-the-badge" alt="Platform"/>
  <img src="https://img.shields.io/badge/Visualization-Plotly%20%7C%20Offline-orange?style=for-the-badge&logo=plotly&logoColor=white" alt="Plotly"/>
</p>

<p align="center">
  <b>Autonomous, layout-invariant financial audit pipeline for state and municipal statutory accounts.</b><br>
  Transforms unstructured PDF annual audit reports into clean analytical workbooks and a self-contained, interactive HTML dashboard.
</p>

---

## ⚡ Key Highlights & Core Metrics


```

┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  212 Local      │     │  ₹200+ Crore    │     │  100% Offline   │
│  Bodies Covered │     │  Audit Ledger   │     │  HTML Dashboard │
└─────────────────┘     └─────────────────┘     └─────────────────┘
│                       │                       │
▼                       ▼                       ▼
• 199 Gram Panchayats   • 15th & 14th FC Grants • Zero web servers
• 6 Zilla Panchayats    • State Finance Comm.   • Plotly interactive
• 7 Urban Local Bodies  • Own Source Revenue    • Filterable schedules

```

| Metric | Details |
| :--- | :--- |
| **Audit Cycles Supported** | FY 2021–22, 2022–23, 2023–24, and 2024–25 (CARR 2023 through 2026) |
| **Parsing Engine** | Tail-first column resolver handling 6, 7, and 8-column layout drift |
| **Risk Detection** | Automated flagging for **Severe Fund Parking** (utilization $\le 20\%$ on $\ge \text{₹}20\text{L}$) |
| **Zero-Config Ingestion** | Autonomously detects filenames, page offsets, and financial years |

---

## 📚 Essential Documentation Links

* 🛠️ **[INSTALLATION.md](INSTALLATION.md)** — Complete step-by-step setup for **Windows Terminal**, **VSCodium**, **Linux (Ubuntu/Debian)**, and **GitHub Pages deployment**.
* 🔍 **[TROUBLESHOOT.md](TROUBLESHOOT.md)** — Diagnostics, system binary resolutions (Poppler, Tesseract, Ghostscript), and common parsing/math fixes.

---

## 🚀 Quick Start (Windows & Linux)

### 1. Clone & Set Up Virtual Environment

**Windows (PowerShell):**
```powershell
# Navigate and configure local execution
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

# Create and activate virtual environment
python -m venv venv
.\venv\Scripts\Activate.ps1

# Install core dependencies
python -m pip install --upgrade pip
pip install -r requirements.txt

```

**Linux / macOS (Bash):**

```bash
python3 -m venv venv
source venv/bin/activate
python3 -m pip install --upgrade pip
pip install -r requirements.txt

```

---

### 2. Add Source Audit Reports

Drop your PDF files into `input_pdfs/`:

```text
input_pdfs/
├── CARR 2023.pdf
├── CARR 2024.pdf
├── CARR 2025.pdf
└── CARR 2026.pdf

```

---

### 3. Run the Unified Pipeline

Execute the full suite in a single command:

**Python Orchestrator (Cross-Platform):**

```bash
python run_pipeline.py

```

**Native Scripts:**

* **Windows:** `.\run_pipeline.bat`
* **Linux / macOS:** `chmod +x run_pipeline.sh && ./run_pipeline.sh`

---

## 📂 Project Structure

```text
carr_audit_suite/
├── input_pdfs/                        <-- Drop raw PDF annual reports here
│   ├── CARR 2023.pdf
│   ├── CARR 2024.pdf
│   ├── CARR 2025.pdf
│   └── CARR 2026.pdf
├── output/                            <-- Auto-generated Excel & HTML artifacts
│   ├── carr_multiyear_master.xlsx
│   ├── carr_comparative_analysed.xlsx
│   └── carr_comparative_dashboard.html
├── templates/
│   └── dashboard_template.html        <-- Offline Jinja2 dashboard UI
├── 01_extract_universal.py            <-- Dynamic PDF boundary & table parser
├── 02_comparative_analytics.py        <-- Multi-year YoY, CAGR & risk analytics
├── 03_dashboard_builder.py            <-- Interactive Plotly HTML generator
├── run_pipeline.py                    <-- Sequential pipeline runner
├── run_pipeline.bat                   <-- Windows batch launcher
├── run_pipeline.sh                    <-- Linux/macOS shell runner
├── requirements.txt                   <-- Python package manifest
├── INSTALLATION.md                    <-- Setup & GitHub deployment guide
├── TROUBLESHOOT.md                    <-- Issue runbook & bug solutions
└── README.md

```

---

## 📊 Pipeline Processing Stages

```
   input_pdfs/*.pdf
          │
          ▼
┌───────────────────────────────┐
│ 01_extract_universal.py       │ ──► Auto-detects FY and Appendix II/III page ranges
│                               │ ──► Tail-first normalization (6/7/8 column invariant)
└───────────────────────────────┘
          │
          ▼
   output/carr_multiyear_master.xlsx
          │
          ▼
┌───────────────────────────────┐
│ 02_comparative_analytics.py   │ ──► Cross-year YoY absolute & % changes
│                               │ ──► Multi-period CAGR & utilization ratios
│                               │ ──► Exception filter: Idle capital >= ₹20L & <= 20%
└───────────────────────────────┘
          │
          ▼
   output/carr_comparative_analysed.xlsx
          │
          ▼
┌───────────────────────────────┐
│ 03_dashboard_builder.py       │ ──► Plotly multi-axis macro trajectories
│                               │ ──► Scheme-wise trends & anomaly bubble plots
│                               │ ──► Self-contained, responsive offline HTML
└───────────────────────────────┘
          │
          ▼
   output/carr_comparative_dashboard.html

```

---

## 🎯 Deliverables & Outputs

| Deliverable | Format | Operational Role |
| --- | --- | --- |
| **Master Ledger** | `output/carr_multiyear_master.xlsx` | Normalized, granular extraction records with parent-child grouping (`District -> BAC -> Unit -> Scheme`) and source pagination tracking. |
| **Analytics Workbook** | `output/carr_comparative_analysed.xlsx` | Multi-year macro trajectory, wide-format YoY variance tables, and severe fund parking registers. |
| **Interactive Dashboard** | `output/carr_comparative_dashboard.html` | Double-click runnable HTML report featuring responsive Plotly visuals, interactive tabs, KPI metric cards, and sortable schedules. |

---

## 🔧 Customization Quick Reference

### Adjusting Fund Parking Severity

Edit the threshold logic in `02_comparative_analytics.py`:

```python
# Default: Balance >= ₹20,00,000 and utilization <= 20%
if row["Total_Receipt"] >= 2_000_000 and row["Utilization_Pct"] <= 20.0:
    return "HIGH RISK (Severe Fund Parking)"

```

### Adding Normalized Scheme Mappings

Add classification rules to `normalize_scheme()` in `01_extract_universal.py`:

```python
def normalize_scheme(raw_scheme: str) -> str:
    u = str(raw_scheme).upper()
    if "YOUR_TAG" in u:
        return "Custom Scheme Bucket"
    # ... standard mappings ...

```

---

## 📄 License & Attribution

Built for statutory and public finance audits. Designed in compliance with the reporting guidelines of the Sikkim Local Fund Audit Act, 2012, and Central Finance Commission monitoring frameworks.