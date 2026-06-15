"""
Build dashboard_data.json from the April 2026 shipment log + rate cards
(via accrual_engine) and Denise's historical accrual log.

Run:
    python3 build_dashboard_data.py

Output:
    ../dashboard/dashboard_data.json
"""

import json
import os
import sys
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from accrual_engine import run_accrual_engine

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(HERE)
DATA_DIR = os.path.join(REPO_ROOT, "data")
SHIPMENTS_CSV = os.path.join(DATA_DIR, "shipments_apr2026.csv")
INVOICES_CSV = os.path.join(DATA_DIR, "freight_invoices_oct2025_mar2026_v2.csv")
DENISE_CSV = os.path.join(DATA_DIR, "denise_accruals_v2.csv")
OUT_DIR = os.path.join(REPO_ROOT, "dashboard")
OUT_PATH = os.path.join(OUT_DIR, "dashboard_data.json")


def build():
    result = run_accrual_engine(SHIPMENTS_CSV, INVOICES_CSV)

    # ── Carrier summary for KPI cards / charge breakdown ───────────────
    carrier_summary = {}
    for carrier, s in result["summary"].items():
        carrier_summary[carrier] = {
            "shipments": s["shipment_count"],
            "base": s["base_total"],
            "fuel": s["fuel_total"],
            "accessorial": s["accessorial_total"],
            "adjustment": s["adjustment"],
            "net": s["net_accrual"],
        }

    # ── Denise's historical actuals (Oct 2025 - Mar 2026), for the trend chart ─
    denise = pd.read_csv(DENISE_CSV)
    historical = {}
    for month in denise["month"].unique():
        sub = denise[denise["month"] == month]
        historical[month] = {
            row["carrier"]: round(row["actual_invoiced"])
            for _, row in sub.iterrows()
        }

    # ── Denise's April estimate = trailing 3-month average of actuals (Jan-Mar 2026) ─
    trailing_months = ["January 2026", "February 2026", "March 2026"]
    denise_avgs = {}
    for carrier in ["Peak Logistics", "Heartland Freight", "Coastal Express"]:
        vals = [historical[m][carrier] for m in trailing_months]
        denise_avgs[carrier] = round(sum(vals) / len(vals))

    # ── Denise's track record (for "why this is better" framing) ───────
    denise["abs_variance"] = denise["variance_dollars"].abs()
    denise_track_record = {
        "avg_abs_variance": round(denise["abs_variance"].mean()),
        "min_abs_variance": int(denise["abs_variance"].min()),
        "max_abs_variance": int(denise["abs_variance"].max()),
        "months_covered": denise["month"].nunique(),
    }

    # ── Shipment-level detail for the interactive table ─────────────────
    shipment_detail = []
    for est in result["estimates"]:
        shipment_detail.append({
            "id": est.shipment_id,
            "carrier": est.carrier,
            "city": est.dest_city,
            "state": est.dest_state,
            "zip": est.dest_zip,
            "weight": round(est.weight, 1),
            "zone": est.zone_or_region,
            "base": est.base_charge,
            "fuel": est.fuel_surcharge,
            "acc": est.accessorial_total,
            "acc_detail": est.accessorial_detail,
            "total": est.total_charge,
            "notes": est.notes,
            "flags": est.flags,
        })

    data = {
        "service_month": "April 2026",
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "adjustment_rate": result["adjustment_rate"],
        "total": result["total_accrual"],
        "carrier_summary": carrier_summary,
        "denise_avgs": denise_avgs,
        "denise_track_record": denise_track_record,
        "historical": historical,
        "flags": result["flags"],
        "shipment_detail": shipment_detail,
    }

    os.makedirs(OUT_DIR, exist_ok=True)
    with open(OUT_PATH, "w") as f:
        json.dump(data, f, indent=2)

    print(f"Wrote {OUT_PATH}")
    print(f"Total accrual: ${data['total']:,.2f}  |  "
          f"{len(shipment_detail)} shipments  |  {len(result['flags'])} flags")
    return data


if __name__ == "__main__":
    build()
