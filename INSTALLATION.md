# Comprehensive Installation & Deployment Guide

This guide provides end-to-end setup instructions for running the **Sikkim Local Fund Audit Multi-Year Pipeline** using **VSCodium** and native terminals on both **Windows** and **Linux**, as well as step-by-step instructions for publishing and deploying the project on **GitHub**.

---

## 1. System Requirements & External Binaries

The extraction engine uses `pdfplumber`, `pdf2image`, `pytesseract`, and `camelot-py`. These rely on underlying system binaries:
* **Python 3.9+**
* **Poppler** (PDF rendering and raster conversion)
* **Tesseract OCR** (Fallback optical character recognition for scanned/image PDFs)
* **Ghostscript** (PostScript/PDF interpreter required by Camelot)

---

## 2. Windows Setup (Windows Terminal + VSCodium)

### Step 2.1: Install Native System Packages

1. **Python 3.9+**:
   - Download the official installer from [python.org](https://www.python.org/downloads/).
   - **Crucial:** Ensure the checkbox **"Add python.exe to PATH"** is selected during installation.

2. **Poppler for Windows**:
   - Download the latest binary archive from [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases/).
   - Extract the archive to `C:\Program Files\poppler`.
   - Add `C:\Program Files\poppler\Library\bin` (or `C:\Program Files\poppler\bin`) to your Windows System `PATH`.

3. **Tesseract OCR**:
   - Download the Windows 64-bit installer from [UB-Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki).
   - Install to the default location (`C:\Program Files\Tesseract-OCR`).
   - Add `C:\Program Files\Tesseract-OCR` to your Windows System `PATH`.

4. **Ghostscript (for Camelot)**:
   - Download and run the AGPL 64-bit executable from [Ghostscript.com](https://ghostscript.com/releases/gsdnld.html).
   - Ensure the directory containing `gswin64c.exe` is registered in your `PATH`.

> **How to Update PATH in Windows:**  
> Press `Win + R` $\to$ type `sysdm.cpl` $\to$ click **Advanced** $\to$ **Environment Variables...** $\to$ select **Path** under *System variables* $\to$ click **Edit...** $\to$ add the paths $\to$ click **OK**.

---

### Step 2.2: Workspace Setup via Windows Terminal

Open **Windows Terminal** (PowerShell) and run:

```powershell
# 1. Navigate to your projects directory and create workspace
cd C:\Users\$env:USERNAME\Documents
mkdir carr_audit_suite
cd carr_audit_suite

# 2. Open project folder in VSCodium
codium .

```

---

### Step 2.3: Virtual Environment & Dependency Installation

In **Windows Terminal** or the VSCodium integrated terminal (`Ctrl + ` `):

```powershell
# 1. Enable execution of local scripts for PowerShell (run once if restricted)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

# 2. Create the Python virtual environment
python -m venv venv

# 3. Activate the virtual environment
.\venv\Scripts\Activate.ps1

# 4. Upgrade pip and install all Python libraries
python -m pip install --upgrade pip
pip install -r requirements.txt

```

---

### Step 2.4: Configure VSCodium Python Interpreter

1. In VSCodium, press `Ctrl + Shift + P` to display the Command Palette.
2. Type and select **`Python: Select Interpreter`**.
3. Choose the virtual environment path: `.\venv\Scripts\python.exe`.
4. Open the Extensions sidebar (`Ctrl + Shift + X`) and install:
* **Python** (Open VSX)
* **Pylance** or **Pyright**



---

### Step 2.5: Verify Windows Environment

```powershell
python -c "import sys; print(sys.executable)"
# Output must end in: \carr_audit_suite\venv\Scripts\python.exe

python -c "import pdfplumber, openpyxl, pandas, plotly, jinja2, scipy; print('All core libraries verified!')"

```

---

## 3. Linux Setup (Ubuntu / Debian Terminal + VSCodium)

### Step 3.1: Install Native System Packages

Open your Linux terminal (`Ctrl + Alt + T`) and install Python, Poppler, Tesseract, and Ghostscript:

```bash
# Update package repositories
sudo apt-get update

# Install Python, build tools, and PDF/OCR system binaries
sudo apt-get install -y \
    python3 \
    python3-pip \
    python3-venv \
    poppler-utils \
    tesseract-ocr \
    libtesseract-dev \
    ghostscript \
    default-jre

```

---

### Step 3.2: Workspace Setup

```bash
# 1. Create directory and navigate into it
mkdir -p ~/projects/carr_audit_suite
cd ~/projects/carr_audit_suite

# 2. Launch VSCodium
codium .

```

---

### Step 3.3: Virtual Environment & Dependency Installation

In your terminal or the VSCodium integrated terminal (`Ctrl + ` `):

```bash
# 1. Create a dedicated virtual environment
python3 -m venv venv

# 2. Activate the virtual environment
source venv/bin/activate

# 3. Upgrade pip and install Python packages
python3 -m pip install --upgrade pip
pip install -r requirements.txt

```

---

### Step 3.4: Configure VSCodium Python Interpreter (Linux)

1. Press `Ctrl + Shift + P`.
2. Type and select **`Python: Select Interpreter`**.
3. Select the environment located at: `~/projects/carr_audit_suite/venv/bin/python`.

---

### Step 3.5: Verify Linux Environment

```bash
python3 -c "import sys; print(sys.executable)"
# Output must end in: /carr_audit_suite/venv/bin/python

python3 -c "import pdfplumber, openpyxl, pandas, plotly, jinja2, scipy; print('All core libraries verified!')"

```

---

## 4. Running the Pipeline

Before executing, place all audit report PDFs (e.g., `CARR 2023.pdf`, `CARR 2024.pdf`, `CARR 2025.pdf`, `CARR 2026.pdf`) inside the `input_pdfs/` directory.

### Running Everything in One Command

* **Windows (PowerShell):**
```powershell
python run_pipeline.py

```


*Or via Batch Script:*
```cmd
.\run_pipeline.bat

```


* **Linux (Bash):**
```bash
chmod +x run_pipeline.sh
./run_pipeline.sh

```



---

### Running Individual Stages Manually

1. **Stage 1 — Multi-Year Adaptive Ingestion**:
```bash
python 01_extract_universal.py

```


*Scans `input_pdfs/`, discovers reporting appendices, resolves columns using tail-first alignment, and writes `output/carr_multiyear_master.xlsx`.*
2. **Stage 2 — Comparative Calculations & Anomaly Detection**:
```bash
python 02_comparative_analytics.py

```


*Generates `output/carr_comparative_analysed.xlsx` containing multi-year variances, CAGR, and severe fund parking flags.*
3. **Stage 3 — Compile Interactive Dashboard**:
```bash
python 03_dashboard_builder.py

```


*Produces the standalone HTML dashboard at `output/carr_comparative_dashboard.html`.*

---

## 5. GitHub Deployment & Version Control

Follow these steps to initialize Git, ignore heavy data/PDF files, push to a remote repository, and publish the dashboard using **GitHub Pages**.

### Step 5.1: Create `.gitignore`

Prevent large PDF files, cached bytecode, and local virtual environments from cluttering Git history:

```gitignore
# Byte-compiled / optimized files
__pycache__/
*.py[cod]
*$py.class

# Virtual environments
venv/
env/
ENV/

# Large binaries & PDF datasets
input_pdfs/*.pdf

# Local IDE configuration
.vscode/
.vscodium/
.idea/

# OS specific files
.DS_Store
Thumbs.db

```

---

### Step 5.2: Initialize and Push to GitHub

Run these commands in your project root:

```bash
# 1. Initialize local git repository
git init -b main

# 2. Stage configuration, code, and template files
git add .gitignore requirements.txt README.md TROUBLESHOOT.md INSTALLATION.md
git add 01_extract_universal.py 02_comparative_analytics.py 03_dashboard_builder.py
git add run_pipeline.py run_pipeline.bat run_pipeline.sh
git add templates/

# 3. Commit staged files
git commit -m "feat: initial commit of multi-year audit pipeline suite"

# 4. Link your remote GitHub repository (replace with your URL)
git remote add origin [https://github.com/](https://github.com/)<your-username>/carr-audit-suite.git

# 5. Push to GitHub
git push -u origin main

```

---

### Step 5.3: Automated Hosting via GitHub Pages (Optional)

To serve the generated `carr_comparative_dashboard.html` via GitHub Pages:

1. Create a docs directory or configure Pages to read from a dedicated branch:
```bash
# Copy the dashboard output to an index file for static web hosting
mkdir -p docs
cp output/carr_comparative_dashboard.html docs/index.html
git add docs/index.html
git commit -m "docs: publish interactive dashboard to GitHub Pages"
git push origin main

```


2. Navigate to your repository on **GitHub.com**.
3. Go to **Settings** $\to$ **Pages** (under the "Code and automation" menu).
4. Under **Build and deployment**:
* **Source:** Deploy from a branch
* **Branch:** `main` / folder: `/docs`


5. Click **Save**. GitHub Pages will deploy your standalone interactive dashboard at:
`https://<your-username>.github.io/carr-audit-suite/`
