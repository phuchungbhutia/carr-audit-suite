# Comprehensive Installation & Deployment Guide

This guide provides end-to-end setup instructions for running the **Sikkim Local Fund Audit Multi-Year Pipeline** using **VSCodium** and native terminals on both **Windows** and **Linux**, as well as complete instructions for publishing and deploying the project on **GitHub**.

---

## 1. System Requirements & External Binaries

The extraction engine relies on `pdfplumber`, `pdf2image`, `pytesseract`, and `camelot-py`, which depend on underlying system binaries:
* **Python 3.9+**
* **Poppler** (PDF rendering and raster conversion)
* **Tesseract OCR** (Fallback optical character recognition for scanned/image PDFs)
* **Ghostscript** (PostScript/PDF interpreter required by Camelot)

---

## 2. Windows Setup (Windows Terminal + VSCodium)

### Step 2.1: Install Native System Packages

1. **Python 3.9+**:
   - Download the official installer from [python.org](https://www.python.org/downloads/).
   - Ensure the checkbox **"Add python.exe to PATH"** is selected during installation.

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

> **Updating PATH in Windows:**  
> Press `Win + R` $\to$ type `sysdm.cpl` $\to$ click **Advanced** $\to$ **Environment Variables...** $\to$ select **Path** under *System variables* $\to$ click **Edit...** $\to$ add paths $\to$ click **OK**.

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

Open your Linux terminal (`Ctrl + Alt + T`):

```bash
sudo apt-get update
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
mkdir -p ~/projects/carr_audit_suite
cd ~/projects/carr_audit_suite
codium .

```

---

### Step 3.3: Virtual Environment & Dependency Installation

```bash
python3 -m venv venv
source venv/bin/activate
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

Place all audit report PDFs (e.g., `CARR 2023.pdf`, `CARR 2024.pdf`, `CARR 2025.pdf`, `CARR 2026.pdf`) inside the `input_pdfs/` directory.

### Running Everything in One Command

* **Windows (PowerShell):**

```powershell
python run_pipeline.py

```

* **Windows (Batch):**

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

2. **Stage 2 — Comparative Calculations & Anomaly Detection**:

```bash
python 02_comparative_analytics.py

```

3. **Stage 3 — Compile Interactive Dashboard**:

```bash
python 03_dashboard_builder.py

```

---

## 5. GitHub Repository Setup & Deployment (phuchungbhutia)

To create and publish your GitHub repository under your account **`phuchungbhutia`** using **Windows Terminal (PowerShell)** and **VSCodium**, follow these steps:

### Step 1: Install Git & GitHub CLI (gh)

If you haven't installed Git or the GitHub CLI tool yet:

1. Open **Windows Terminal (PowerShell)**.
2. Install Git and the GitHub CLI using Windows Package Manager (`winget`):

```powershell
winget install --id Git.Git -e --source winget
winget install --id GitHub.cli -e --source winget

```

3. Restart Windows Terminal so the new system commands load into your session.

### Step 2: Configure Git & Authenticate with GitHub

Set your identity and log into your account:

1. Configure your Git username and email:

```powershell
git config --global user.name "phuchungbhutia"
git config --global user.email "your-email-associated-with-github@example.com"

```

2. Log into your GitHub account using the GitHub CLI:

```powershell
gh auth login

```

* Choose: **GitHub.com**
* Preferred protocol: **HTTPS**
* Authenticate Git with your GitHub credentials: **Yes**
* How to authenticate: **Login with a web browser**
* Press Enter, copy the 8-character code, and authorize it in your browser.

### Step 3: Open Your Project in VSCodium

Navigate to your project folder:

```powershell
cd C:\Users\$env:USERNAME\Documents\audit_suite

```

Open the workspace in VSCodium:

```powershell
codium .

```

### Step 4: Ensure .gitignore Exists

Make sure your repository does not track large PDF datasets, virtual environments, or compiled caches.

In VSCodium, verify you have a `.gitignore` file with the following lines:

```gitignore
# Virtual environment
venv/
env/

# Python caches
__pycache__/
*.pyc

# Local IDE configuration
.vscode/
.vscodium/

# Source annual report PDFs (keep repos light)
input_pdfs/*.pdf

# System files
Thumbs.db
.DS_Store

```

### Step 5: Initialize Git and Create the Remote Repository

Run these commands directly in **Windows Terminal** (or VSCodium’s integrated terminal `Ctrl + ` `):

1. Initialize the local Git repository:

```powershell
git init -b main

```

2. Stage your files:

```powershell
git add .

```

3. Make your initial commit:

```powershell
git commit -m "feat: initial commit of audit analytics pipeline"

```

4. Create the GitHub repository under **phuchungbhutia** and push:

```powershell
gh repo create carr-audit-suite --public --source=. --remote=origin --push

```

*(Replace `carr-audit-suite` with your preferred repository name, or change `--public` to `--private` if you want it private).*

### Step 6: Verify the Repository

Open your repository in your browser:

```powershell
gh repo view --web

```

Your repository will now be live at:
[https://github.com/phuchungbhutia/carr-audit-suite](https://github.com/phuchungbhutia/carr-audit-suite?utm_source=gemini)

---

### Everyday Workflow Commands

When you make changes to files inside VSCodium, push updates using standard Git commands:

```powershell
git add .
git commit -m "Update parsing and dashboard logic"
git push

```

---

## 6. Automated Hosting via GitHub Pages (Optional)

To serve the generated `carr_comparative_dashboard.html` directly from your repository:

1. Create a `docs` directory and copy the dashboard output:

```powershell
mkdir -p docs
cp output/carr_comparative_dashboard.html docs/index.html
git add docs/index.html
git commit -m "docs: publish interactive dashboard to GitHub Pages"
git push

```

2. In your repository on **GitHub.com**, navigate to **Settings** $\to$ **Pages**.
3. Under **Build and deployment**:

* **Source:** Deploy from a branch
* **Branch:** `main` / folder: `/docs`

4. Save the configuration. Your dashboard will be hosted at:
   `https://phuchungbhutia.github.io/carr-audit-suite/`
