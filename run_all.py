"""
Run the full Ridgeline Foods freight accrual pipeline:
  1. Run the accrual engine against data/shipments_apr2026.csv
  2. Build the Excel workbook (output/ridgeline_freight_accrual_apr2026.xlsx)
  3. Build dashboard_data.json for the HTML dashboard

Usage:
    python3 run_all.py
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "src"))

from build_excel_report import build_workbook
from build_dashboard_data import build as build_dashboard_data

DATA_DIR = os.path.join(HERE, "data")
OUT_DIR = os.path.join(HERE, "output")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    print("=" * 60)
    print("STEP 1-2: Accrual engine + Excel workbook")
    print("=" * 60)
    build_workbook(
        os.path.join(DATA_DIR, "shipments_apr2026.csv"),
        os.path.join(OUT_DIR, "ridgeline_freight_accrual_apr2026.xlsx"),
        os.path.join(DATA_DIR, "freight_invoices_oct2025_mar2026_v2.csv"),
    )

    print("\n" + "=" * 60)
    print("STEP 3: Dashboard data")
    print("=" * 60)
    build_dashboard_data()

    print("\nDone. Next steps:")
    print("  - Open output/ridgeline_freight_accrual_apr2026.xlsx for the JE-ready workbook")
    print("  - cd dashboard && python3 -m http.server 8000, then open http://localhost:8000")


if __name__ == "__main__":
    main()
