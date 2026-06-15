# Ridgeline Foods — April 2026 Freight Accrual

A rate-card-based freight accrual engine, journal-entry-ready Excel workbook, and
interactive close dashboard — built to replace Denise's trailing 3-month-average
accrual for Ridgeline Foods' three freight carriers (Peak Logistics, Heartland
Freight, Coastal Express).

**TL;DR:** `python3 run_all.py` runs the full pipeline. It produces:

- `output/ridgeline_freight_accrual_apr2026.xlsx` — journal-entry-ready workbook
  (Accrual Summary, Shipment Detail, Data Quality Flags, Assumptions & Controls)
- `dashboard/dashboard_data.json` — the data file the HTML dashboard reads

---

## Why this exists

Ridgeline's April books need to close by business day 8, but carrier invoices
don't arrive until the 8th–12th of the following month. Denise's old process —
a trailing 3-month average of actual invoices — missed the actual invoice by
$5K–$13K most months because it's based on *last quarter's* shipment mix, not
*this month's*.

This tool instead estimates April's freight expense **directly from April's
3PL shipment log**, applying each carrier's actual contracted rate card
(per-mile tiers, zone flat rates, per-pound regional rates, fuel surcharges,
accessorials, minimums, and volume-discount tiers). It's transparent
(every shipment shows its full charge math), auditable (every assumption is
listed and most are computed from the data itself, not hardcoded), and
repeatable (drop in next month's shipment log and re-run).

---

## Repo structure

```
ridgeline-freight-accrual/
├── README.md
├── requirements.txt
├── run_all.py                  # one-command pipeline
├── data/                        # input data (provided by Numeric)
│   ├── shipments_apr2026.csv
│   ├── freight_invoices_oct2025_mar2026_v2.csv
│   ├── denise_accruals_v2.csv
│   ├── rate_card_peak_logistics.csv
│   ├── rate_card_heartland_freight.csv
│   └── rate_card_coastal_express.csv
├── src/
│   ├── accrual_engine.py        # core rate-card estimation engine
│   ├── build_excel_report.py    # builds the 4-tab Excel workbook
│   └── build_dashboard_data.py  # builds dashboard_data.json
├── dashboard/
│   ├── index.html               # interactive close dashboard
│   └── dashboard_data.json   
└── output/
    └── ridgeline_freight_accrual_apr2026.xlsx
```

---

## How to run it

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the full pipeline

```bash
python3 run_all.py
```

This regenerates both `output/ridgeline_freight_accrual_apr2026.xlsx` and
`dashboard/dashboard_data.json` from the CSVs in `data/`.

### 3. View the dashboard

The dashboard fetches `dashboard_data.json` via `fetch()`, which browsers
block on `file://` pages — so serve it locally:

```bash
cd dashboard
python3 -m http.server 8000
```

Then open **http://localhost:8000/index.html**.

### 4. Open the Excel workbook

`output/ridgeline_freight_accrual_apr2026.xlsx` has 4 tabs:

1. **Accrual Summary** — carrier-by-carrier totals, suggested journal entry
   (Dr. 6300 Freight Expense / Cr. 2100 Accrued Liabilities), and a comparison
   to Denise's trailing 3-month average.
2. **Shipment Detail** — all 160 shipments with full base/fuel/accessorial
   breakdown, filterable/sortable, with a Notes/Flags column.
3. **Data Quality Flags** — every exception the engine found (duplicates,
   imputed weights, out-of-territory shipments, unusual routings).
4. **Assumptions & Controls** — every rate, mileage source, and judgment
   call documented for an auditor or teammate.

---

## Methodology by carrier

### Peak Logistics (Mountain West) — per-mile × weight tier
- Rate tiers: `<500 lbs` = $3.25/mi, `500–2,000 lbs` = $4.80/mi, `>2,000 lbs` = $6.40/mi
- 14% fuel surcharge on base charge (fixed through Q2 2026)
- $185 minimum charge per shipment
- Mileage: 11 destinations are carrier-verified from the rate card. The
  remaining destinations (Fort Collins, Pueblo, Grand Junction, Laramie,
  Casper) are **back-calculated from 6 months of invoice history**
  (`median(base_charge / per-mile rate)`), excluding invoices that hit the
  $185 minimum.
- Reno, NV is outside Peak's territory — flagged, 900-mile proxy used.
- 5 shipments originate from Salt Lake City (unusual backhaul) — flagged for
  CFO review; Denver-based mileage used as a proxy.
- Accessorials priced from `special_handling`: Liftgate $75, Residential $45,
  Appointment $50, Inside Delivery $125.

### Heartland Freight (Midwest) — flat rate by zone
- Zone assigned from the **first 3 digits of the destination ZIP** using the
  carrier's ZIP-prefix lookup table (not state — e.g. Kansas City KS is Zone 1
  despite being in Kansas).
- Flat rates: Zone 1 $320, Zone 2 $485, Zone 3 $610, Zone 4 $780 — **fuel
  surcharge is included**, no separate FSC line.
- Volume discount: April = first month of Q2, so every shipment is **Tier 1
  (0% discount)** per the quarterly reset. (January's $7,133 miss vs. Denise's
  estimate was driven by exactly this Q1-reset dynamic — the engine explicitly
  accounts for it this quarter.)
- Accessorials: Liftgate $85, Inside Delivery $135, Appointment $40.

### Coastal Express (West Coast) — per-lb by region
- Region from destination ZIP: SoCal (90000–92899) $0.48/lb, NorCal
  (93000–96199) $0.55/lb, PNW (97000–99499) $0.72/lb.
- $28 minimum base charge.
- 9.5% fuel surcharge on base only (not on accessorials), fixed through Q2 2026.
- Residential surcharge (only if `residential = TRUE`): `<50 lbs` = $12.50,
  `50–500 lbs` = $35, `>500 lbs` = $65.
- Accessorials: Liftgate $90, Inside Delivery $110, Appointment $55.

### General controls (apply to all carriers)
- **Carrier name normalization** — the 3PL log has 9 spelling/casing variants
  across 3 carriers (e.g. "PEAK LOG", "Coastal Express LLC"); all mapped to
  canonical names.
- **Duplicate detection** — SHP-10033 appears twice (identical); second
  occurrence excluded.
- **Missing weight imputation** — 5 shipments had blank `weight_lbs`, imputed
  as `units × carrier-specific avg lbs/unit` from the rest of April's shipments.
- **Adjustment buffer** — computed dynamically each run as
  `sum(historical adjustments) / sum(pre-adjustment gross)` across 6 months of
  invoice history (currently ~0.3%, all 22 historical adjustments are negative
  credits), applied as a conservative haircut to the gross estimate.

Every one of these is also surfaced in the **Data Quality Flags** tab and the
dashboard's flags panel — nothing is silently corrected.

---

## What I'd build next with more time

- **Actuals-to-estimate reconciliation**: once May invoices land, automatically
  match `shipment_ref` → `shipment_id` and compute per-shipment variance, broken
  out by base/fuel/accessorial so we can see *which assumption* drove any miss.
- **Rate card versioning**: parse the rate card CSVs directly (they're already
  structured close to machine-readable) so a rate change next quarter doesn't
  require a code edit — just a new CSV.
- **Heartland volume-tier tracking**: maintain a running QTD shipment counter so
  Tier 2/3/4 discounts kick in automatically mid-quarter instead of assuming
  Tier 1 all quarter.
- **Confidence intervals**: use the historical variance distribution (from
  `denise_accruals_v2.csv`) to show an estimate range, not just a point estimate.
- **Auto-posting**: push the journal entry to NetSuite via API once it's been
  reviewed/approved in the dashboard.
