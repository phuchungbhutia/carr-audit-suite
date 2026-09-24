#!/usr/bin/env bash
set -e

echo "=========================================================="
echo " Starting Sikkim Local Fund Audit Comparative Pipeline"
echo "=========================================================="

python3 01_extract_universal.py
python3 02_comparative_analytics.py
python3 03_dashboard_builder.py

echo ""
echo "=========================================================="
echo " [SUCCESS] Pipeline completed successfully."
echo " Interactive Dashboard: output/carr_comparative_dashboard.html"
echo "=========================================================="