# Troubleshooting Guide & Operational Runbook

This guide covers common environmental, dependency, parsing, and data validation issues encountered when deploying and running the Local Fund Audit Multi-Year Pipeline across Windows, Linux, and macOS environments.

---

## Quick Reference: Top Failure Modes

| Error Signature | Stage | Root Cause | Immediate Resolution |
| :--- | :--- | :--- | :--- |
| `ModuleNotFoundError: No module named 'plotly'` | Stage 3 | Missing package in active Python environment | Run `python -m pip install plotly jinja2` inside the active virtual environment (`venv`). |
| `ValueError: Worksheet named 'GPU_Master_Ledger' not found` | Stage 2 | Stage 1 extracted 0 records; sheet was not generated | Correct inverted page ranges or update `01_extract_universal.py` with dynamic page detection. |
| `ERROR: Could not open requirements file: [Errno 2]` | Setup | Terminal working directory is not the project root | Run `pwd` / `cd` to locate the folder containing `requirements.txt`. |
| `NameError: name 'discover_appendix_pages' is not defined` | Stage 1 | Function identifier mismatch in script | Synchronize function definition with its invocation in `01_extract_universal.py`. |
| `pdf2image.exceptions.PDFInfoNotInstalledError` | Stage 1 | Native Poppler binary not present in System `PATH` | Install Poppler and add its `bin/` directory to the system environment path. |
| `pytesseract.TesseractNotFoundError` | Stage 1 | Tesseract OCR engine executable not mapped | Install Tesseract OCR and append its root folder to System `PATH`. |
| `ZeroDivisionError: float division by zero` | Stage 2 | Empty baseline receipts in YoY or CAGR math | Wrap computations with safe-division handlers (`np.where(base != 0, ...)`). |

---

## 1. Environment & Setup Issues

### Issue 1.1: `ERROR: Could not open requirements file: [Errno 2] No such file or directory: 'requirements.txt'`
* **Symptom:** `pip install -r requirements.txt` terminates immediately with `Errno 2`.
* **Root Cause:** The terminal shell prompt is running in an unexpected working directory (e.g., user home directory rather than the project workspace).
* **Fix:**
  1. Inspect the active working directory:
     * **PowerShell / Bash:** `pwd`
     * **CMD:** `cd`
  2. Navigate to the folder containing your project files:
     ```powershell
     cd C:\Users\$env:USERNAME\Documents\audit_suite
     ```
  3. Verify that `requirements.txt` exists:
     * **PowerShell:** `Get-ChildItem requirements.txt`
     * **CMD:** `dir requirements.txt`
     * **Linux / macOS:** `ls -l requirements.txt`
  4. Run installation using the explicit path if preferred:
     ```powershell
     python -m pip install -r .\requirements.txt
     ```

---

### Issue 1.2: `ModuleNotFoundError: No module named 'xyz'` inside VSCodium
* **Symptom:** Terminal installations succeed, but running the code inside VSCodium raises `ModuleNotFoundError` for packages like `plotly`, `pdfplumber`, or `jinja2`.
* **Root Cause:** VSCodium is using the global system Python interpreter instead of the project's local virtual environment (`.\venv\Scripts\python.exe`).
* **Fix:**
  1. In VSCodium, press `Ctrl + Shift + P` to display the Command Palette.
  2. Type and select **`Python: Select Interpreter`**.
  3. Choose the interpreter with the path pointing to `.\venv\Scripts\python.exe`.
  4. Open a fresh terminal (`Ctrl + Shift + \``) and verify:
     ```powershell
     # Windows PowerShell
     Get-Command python | Select-Object Source
     # Output should point to: ...\audit_suite\venv\Scripts\python.exe
     ```

---

### Issue 1.3: PowerShell Script Execution Policy Restriction
* **Symptom:** Activating `venv` produces:
  ```text
  .\venv\Scripts\Activate.ps1 : File cannot be loaded because running scripts is disabled on this system.

```

* **Root Cause:** Default Windows PowerShell security settings restrict arbitrary execution of unsigned scripts.
* **Fix:**
  Grant local script execution for the active terminal session:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
.\venv\Scripts\Activate.ps1

```

---

## 2. External System Binary Dependencies

### Issue 2.1: Poppler Not Detected (`PDFInfoNotInstalledError`)

* **Symptom:** `pdf2image` crashes while attempting raster conversions on image-only documents.
* **Root Cause:** Native Poppler binaries (`pdftoppm`, `pdfinfo`) are missing or omitted from the Windows System `PATH`.
* **Fix:**

1. Download the latest Windows binary package from [Poppler for Windows releases](https://github.com/oschwartz10612/poppler-windows/releases/?utm_source=gemini).
2. Extract the archive (e.g., to `C:\Program Files\poppler`).
3. Add the `Library\bin` (or `bin`) subfolder to the system environment path:

```powershell
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files\poppler\Library\bin", [EnvironmentVariableTarget]::Machine)

```

4. Test availability in a new terminal window:

```powershell
pdfinfo -v

```

---

### Issue 2.2: Tesseract OCR Engine Missing (`TesseractNotFoundError`)

* **Symptom:** Scanned or image-based PDF fallbacks crash with `tesseract is not installed or it's not in your PATH`.
* **Root Cause:** The `tesseract.exe` executable is missing or not registered in the system path.
* **Fix:**

1. Download and run the 64-bit installer from [UB-Mannheim Tesseract](https://www.google.com/search?q=https://github.com/UB-Mannheim/tesseract/wiki&utm_source=gemini).
2. The default installation path is `C:\Program Files\Tesseract-OCR`.
3. Verify the engine in PowerShell:

```powershell
& "C:\Program Files\Tesseract-OCR\tesseract.exe" --version

```

4. Add this directory to your system path, or define it explicitly in code if system permissions are limited:

```python
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

```

---

## 3. Extraction & PDF Parsing Bugs (Stage 1)

### Issue 3.1: Extracted 0 Rows / Inverted Page Ranges (`pages 79-72`)

* **Symptom:**

```text
[INFO] Extracted 0 GPU rows from pages 79-72.
[INFO] Extracted 0 GPU rows from pages 123-122.
ValueError: Worksheet named 'GPU_Master_Ledger' not found

```

* **Root Cause:**

1. Page bounds discovery scanned the entire PDF and updated boundary pointers to the *final* occurrence of strings like `"APPENDIX - II"` instead of the initial anchor.
2. Table of Contents lines (located in the first 15 pages) matched heuristic targets, setting table ranges before the actual data starts.

* **Fix:**
* Skip the initial Table of Contents section (scan only pages $> \text{Page 40}$).
* Sort discovered page indices and anchor `App_II_Start` to the first hit.
* Set hard delimiters based on actual validated physical coordinates:
* `CARR 2023.pdf`: Appendix II pages **56–72** | Appendix III page **73**
* `CARR 2024.pdf`: Appendix II pages **94–122** | Appendix III page **123**
* `CARR 2025.pdf`: Appendix II pages **73–91** | Appendix III pages **92–93**
* `CARR 2026.pdf`: Appendix II pages **83–107** | Appendix III page **108**

---

### Issue 3.2: Column Alignment Shifts (6-Column vs. 7/8-Column Schemas)

* **Symptom:** Closing balances appear in payment columns, or textual descriptions populate numeric cells.
* **Root Cause:** Older volumes (`CARR 2024.pdf`) omit separate `Receipt + Interest` breakdown columns. Counting indices from left to right causes offsets.
* **Fix:**
  Parse from the **tail (right-to-left)**:

```python
# Always true across all years:
closing_balance = trailing_numeric_values[-1]
payment_expenditure = trailing_numeric_values[-2]
total_receipt = trailing_numeric_values[-3]
opening_balance = trailing_numeric_values[-4] if len(trailing_numeric_values) >= 4 else 0.0

```

---

### Issue 3.3: Number Parsing Failures on Indian Formatted Values

* **Symptom:** Values such as `1,23,456.00` turn into `NaN`, or parenthetical negatives `(45,000)` become positive numbers.
* **Root Cause:** Standard `float()` cannot process Indian comma groupings, currency symbols (`₹`, `Rs`), or accounting brackets.
* **Fix:**
  Use a resilient parsing utility:

```python
def parse_indian_amount(val: Any) -> float:
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
        parsed = float(s_clean)
        return -parsed if is_negative else parsed
    except ValueError:
        return 0.0

```

---

## 4. Analytical & Ratio Math Failures (Stage 2)

### Issue 4.1: Division by Zero & Floating Point Overflow

* **Symptom:**

```text
ZeroDivisionError: float division by zero
OverflowError: (34) Result too large

```

* **Root Cause:** Units with zero base expenditures or newly introduced schemes produce infinite percentages during YoY or CAGR calculations.
* **Fix:**
  Wrap computations in vectorized safety checks:

```python
# Safe YoY percentage calculation
df_yoy[pct_col] = np.where(
    df_yoy[base_year] != 0,
    ((df_yoy[target_year] - df_yoy[base_year]) / df_yoy[base_year].abs()) * 100.0,
    np.nan
)

# Safe CAGR calculation
def compute_cagr(start_val: float, end_val: float, periods: int) -> float:
    if periods <= 0 or start_val <= 0 or end_val <= 0:
        return np.nan
    try:
        return (end_val / start_val) ** (1.0 / periods) - 1.0
    except (ZeroDivisionError, OverflowError, ValueError):
        return np.nan

```

---

### Issue 4.2: Double-Counting Embedded Subtotals

* **Symptom:** Total receipts and expenditures appear inflated by roughly 2× compared to published executive summaries.
* **Root Cause:** Subtotal rows (e.g., `Chungthang Total`, `Daramdin Total`) embedded within tables were parsed alongside individual ledger records.
* **Fix:**
  Filter subtotal rows using text markers:

```python
full_row_str = " ".join(row_cells).upper()
if "TOTAL" in full_row_str and not any(k in full_row_str for k in ["RECEIPT", "PAYMENT", "INTEREST"]):
    continue  # Exclude subtotal lines from granular transaction rows

```

---

## 5. Dashboard Generation & Rendering Issues (Stage 3)

### Issue 5.1: Blank White Browser Window when Opening `dashboard.html`

* **Symptom:** Double-clicking `output/carr_comparative_dashboard.html` displays an empty white page.
* **Root Cause:** An unhandled error occurred in the embedded JavaScript, often caused by unescaped JSON data or missing Plotly script tags.
* **Fix:**

1. Open Developer Tools in your browser (`F12` or `Ctrl + Shift + I`) and view the **Console** tab.
2. If Plotly is missing, verify the CDN link in `templates/dashboard_template.html`:

```html
<script src="[https://cdn.plot.ly/plotly-2.27.0.min.js](https://cdn.plot.ly/plotly-2.27.0.min.js)"></script>

```

3. Ensure your Jinja2 template variables use the `| safe` filter so HTML chart fragments render properly:

```html
<div class="card">{{ chart_macro | safe }}</div>

```

---

### Issue 5.2: Plotly Charts Fail to Resize on Tab Switch

* **Symptom:** Charts render with zero width or appear cropped after switching tabs.
* **Root Cause:** Plotly cannot calculate container widths while tabs are hidden (`display: none`).
* **Fix:**
  Add a `window.dispatchEvent` call inside the tab-switching JavaScript function:

```javascript
function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

    const target = document.getElementById(tabId);
    if (target) target.classList.add('active');
    event.currentTarget.classList.add('active');

    // Forces Plotly to recompute responsive widths
    window.dispatchEvent(new Event('resize'));
}

```

---

## 6. End-to-End Diagnostic Checklist

Before running the full pipeline, verify your environment with this checklist:

```powershell
# 1. Verify that Python is running from the virtual environment
python -c "import sys; print(sys.executable)"
# Output must end in: \audit_suite\venv\Scripts\python.exe

# 2. Verify all core libraries import cleanly
python -c "import pdfplumber, openpyxl, pandas, plotly, jinja2, scipy; print('Core imports successful.')"

# 3. Confirm all four source PDFs are located in input_pdfs/
Get-ChildItem .\input_pdfs\*.pdf | Select-Object Name, Length

# 4. Clean out old output artifacts if running a fresh batch
Remove-Item -Path .\output\* -Recurse -Force -ErrorAction SilentlyContinue

# 5. Execute the pipeline orchestrator
python run_pipeline.py

```
---

## 7. GitHub Pages Deployment & Routing Fixes

### Issue 7.1: GitHub Pages 404 Error (`https://<username>.github.io/<repo>/`)
* **Symptom:** Navigating to the deployed URL returns a `404 File Not Found` error page.
* **Root Causes:**
  1. The entry dashboard file is not named `index.html`.
  2. GitHub Pages is configured to publish from `/ (root)` instead of `/docs`.
  3. GitHub's internal Jekyll compiler ignores or misinterprets folders with specific naming conventions.
  4. The deployment Action workflow has not completed building.
* **Fix & Automated Terminal Deploy:**
  Run the following commands in PowerShell to prepare and push the deployment bundle:
  ```powershell
  # 1. Ensure docs directory exists
  New-Item -ItemType Directory -Force -Path docs

  # 2. Copy and rename compiled dashboard to index.html inside docs/
  Copy-Item output\carr_comparative_dashboard.html docs\index.html -Force

  # 3. Add .nojekyll file to bypass Jekyll processing
  New-Item -ItemType File -Force -Path docs\.nojekyll

  # 4. Commit and push to GitHub
  git add docs/
  git commit -m "docs: deploy dashboard as index.html with .nojekyll"
  git push origin main
  ```

Repository Configuration Steps:

Go to your repository on GitHub: https://github.com/phuchungbhutia/carr-audit-suite/settings/pages.

Under Build and deployment:

Source: Deploy from a branch

Branch: main

Folder: /docs

Click Save.

Open the Actions tab (https://github.com/phuchungbhutia/carr-audit-suite/actions) and wait for the pages build and deployment job to show a green checkmark (✔).

Perform a hard refresh in your browser (Ctrl + F5).
