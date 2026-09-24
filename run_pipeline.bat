@echo off
echo ==========================================================
echo Starting Sikkim Local Fund Audit Comparative Pipeline
echo ==========================================================

python 01_extract_universal.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Stage 1 failed. Aborting pipeline.
    exit /b %ERRORLEVEL%
)

python 02_comparative_analytics.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Stage 2 failed. Aborting pipeline.
    exit /b %ERRORLEVEL%
)

python 03_dashboard_builder.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Stage 3 failed. Aborting pipeline.
    exit /b %ERRORLEVEL%
)

echo ==========================================================
echo [SUCCESS] Pipeline completed successfully.
echo Launching output\carr_comparative_dashboard.html...
echo ==========================================================
start output\carr_comparative_dashboard.html