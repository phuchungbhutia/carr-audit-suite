"""
run_pipeline.py
===============
Pipeline runner for cross-platform execution.
Executes extraction, comparative calculations, and dashboard generation sequentially.
"""

import subprocess
import sys


def run_stage(script_name: str) -> None:
    print(f"\n========================================================")
    print(f"  RUNNING STAGE: {script_name}")
    print(f"========================================================")
    proc = subprocess.run([sys.executable, script_name], capture_output=True, text=True)
    if proc.returncode != 0:
        print(f"[ERROR] Stage '{script_name}' failed:\n")
        print(proc.stderr)
        sys.exit(proc.returncode)
    print(proc.stdout)


if __name__ == "__main__":
    print("\n--------------------------------------------------------")
    print("  SIKKIM LOCAL FUND AUDIT: MULTI-YEAR REPORTING PIPELINE")
    print("--------------------------------------------------------")
    run_stage("01_extract_universal.py")
    run_stage("02_comparative_analytics.py")
    run_stage("03_dashboard_builder.py")
    print("\n[SUCCESS] Pipeline executed successfully.")
    print("👉 Interactive Dashboard ready at: output/carr_comparative_dashboard.html\n")