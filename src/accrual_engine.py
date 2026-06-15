"""
Ridgeline Foods — Freight Accrual Engine
Applies contracted rate cards to April 2026 3PL shipment log.

Returns a dict with:
  estimates       — list of ShipmentEstimate dataclass objects
  summary         — dict carrier → {shipment_count, base_total, fuel_total,
                                     accessorial_total, adjustment, net_accrual}
  flags           — list of {shipment, carrier, flag} dicts
  total_accrual   — float
  adjustment_rate — float (historical adj / gross, derived from invoice history)
"""

from __future__ import annotations
import os
import re
from dataclasses import dataclass, field
from typing import Optional
import pandas as pd

# ─── Carrier name normalisation ───────────────────────────────────────────────
_CARRIER_ALIASES: dict[str, str] = {
    # Peak variants
    "peak logistics":     "Peak Logistics",
    "peak log":           "Peak Logistics",
    "peak logistics llc": "Peak Logistics",
    "peak":               "Peak Logistics",
    # Heartland variants
    "heartland freight":     "Heartland Freight",
    "heartland freight inc": "Heartland Freight",
    "heartland freight co.": "Heartland Freight",
    "heartland":             "Heartland Freight",
    # Coastal variants
    "coastal express":     "Coastal Express",
    "coastal express llc": "Coastal Express",
    "coastal exp":         "Coastal Express",
    "coastal":             "Coastal Express",
}

def _normalise_carrier(raw: str) -> str:
    return _CARRIER_ALIASES.get(raw.strip().lower(), raw.strip())


# ─── Mileage table ────────────────────────────────────────────────────────────
# Carrier-verified distances from rate card (Denver origin).
# Back-calculated cities are derived from invoice median below.
_PEAK_MILES_CARRIER_VERIFIED: dict[str, int] = {
    "Colorado Springs": 70,
    "Cheyenne":         100,
    "Pocatello":        440,
    "Idaho Falls":      480,
    "Provo":            490,
    "Salt Lake City":   525,
    "Billings":         550,
    "Ogden":            590,
    "Great Falls":      710,
    "Boise":            830,
    "Missoula":         830,
}

# Back-calculated from 6-months invoice history (median implied miles).
# Formula: median(base_charge / per_mile_rate) excluding minimum-charge hits.
_PEAK_MILES_BACK_CALC: dict[str, int] = {
    "Fort Collins":    52,
    "Pueblo":          89,
    "Laramie":        100,
    "Grand Junction": 222,
    "Casper":         242,
}

PEAK_MILES: dict[str, int] = {**_PEAK_MILES_CARRIER_VERIFIED, **_PEAK_MILES_BACK_CALC}

# Proxy for Reno (out-of-territory) — flagged separately
_RENO_PROXY_MILES = 900


# ─── Rate card constants ──────────────────────────────────────────────────────

# Peak Logistics
PEAK_FSC_RATE      = 0.14
PEAK_MIN_CHARGE    = 185.00
PEAK_ACCESSORIALS  = {"Liftgate": 75, "Residential Delivery": 45,
                       "Appointment Delivery": 50, "Inside Delivery": 125}

def peak_per_mile_rate(weight_lbs: float) -> float:
    if weight_lbs < 500:    return 3.25
    if weight_lbs <= 2000:  return 4.80
    return 6.40

# Heartland Freight
HEARTLAND_ZONE_RATES = {1: 320.00, 2: 485.00, 3: 610.00, 4: 780.00}
HEARTLAND_ACCESSORIALS = {"Liftgate": 85, "Inside Delivery": 135,
                            "Appointment Delivery": 40}

# ZIP prefix → zone lookup (exact rules from rate card)
def _heartland_zone_from_zip(zip_code: int) -> int:
    p = zip_code // 100   # first 3 digits
    if 640 <= p <= 641:   return 1   # KC MO metro core
    if 660 <= p <= 662:   return 1   # KC KS side
    if 663 <= p <= 668:   return 2   # Eastern KS
    if 669 <= p <= 679:   return 2   # Rest of KS (Wichita etc.)
    if 630 <= p <= 639:   return 2   # St. Louis / eastern MO
    if 650 <= p <= 659:   return 2   # Springfield / central MO
    if 500 <= p <= 529:   return 2   # All Iowa
    if 680 <= p <= 685:   return 2   # Omaha / Lincoln NE
    if 600 <= p <= 609:   return 3   # Chicago metro
    if 610 <= p <= 629:   return 3   # Rest of IL (Rockford, Peoria)
    if 530 <= p <= 546:   return 3   # Milwaukee / Madison WI
    if 550 <= p <= 554:   return 3   # Twin Cities MN
    if 547 <= p <= 549:   return 4   # Northern WI
    if 555 <= p <= 567:   return 4   # Greater MN (Duluth, St. Cloud)
    if 686 <= p <= 693:   return 4   # Northern / western NE (Norfolk etc.)
    return 4  # fallback — flag it

# Coastal Express
COASTAL_REGION_RATES  = {"SoCal": 0.48, "NorCal": 0.55, "PNW": 0.72}
COASTAL_MIN_BASE      = 28.00
COASTAL_FSC_RATE      = 0.095
COASTAL_ACCESSORIALS  = {"Liftgate": 90, "Inside Delivery": 110,
                           "Appointment Delivery": 55}
COASTAL_RES_TIERS     = [(50, 12.50), (500, 35.00), (float("inf"), 65.00)]

def _coastal_region(zip_code: int) -> Optional[str]:
    if 90000 <= zip_code <= 92899: return "SoCal"
    if 93000 <= zip_code <= 96199: return "NorCal"
    if 97000 <= zip_code <= 99499: return "PNW"
    return None

def _coastal_res_surcharge(weight: float) -> float:
    for threshold, fee in COASTAL_RES_TIERS:
        if weight < threshold:
            return fee
    return 65.00


# ─── Accessorial parser ───────────────────────────────────────────────────────

def _parse_accessorials(special_handling: str, acc_table: dict[str, float]) -> tuple[float, str]:
    """Return (total_fee, detail_string) for the given special_handling field."""
    if not isinstance(special_handling, str) or not special_handling.strip():
        return 0.0, ""
    total = 0.0
    matched: list[str] = []
    # Split on comma or semicolon
    parts = [p.strip() for p in re.split(r"[,;]", special_handling) if p.strip()]
    for part in parts:
        for label, fee in acc_table.items():
            # substring match, case-insensitive
            if label.lower() in part.lower() or part.lower() in label.lower():
                total += fee
                matched.append(f"{label} ${fee:.0f}")
                break
    return round(total, 2), "; ".join(matched)


# ─── ShipmentEstimate ─────────────────────────────────────────────────────────

@dataclass
class ShipmentEstimate:
    shipment_id:      str
    carrier:          str
    dest_city:        str
    dest_state:       str
    dest_zip:         int
    weight:           float
    zone_or_region:   str
    base_charge:      float
    fuel_surcharge:   float
    accessorial_total: float
    accessorial_detail: str
    total_charge:     float
    notes:            str = ""
    flags:            list[str] = field(default_factory=list)


# ─── Per-carrier estimators ───────────────────────────────────────────────────

def _estimate_peak(row: pd.Series, imputed: bool) -> ShipmentEstimate:
    flags: list[str] = []
    notes_parts: list[str] = []

    city  = row["destination_city"]
    state = row["destination_state"]
    zip_  = int(row["destination_zip"])
    wt    = float(row["weight_lbs"])
    is_slc = str(row["origin_city"]).strip().lower() == "salt lake city"

    # Mileage resolution
    if city == "Reno" and state == "NV":
        miles = _RENO_PROXY_MILES
        flags.append("OUT-OF-TERRITORY: Reno NV is outside Peak's Mountain West coverage. 900-mi proxy used.")
    elif city in PEAK_MILES:
        miles = PEAK_MILES[city]
        if city in _PEAK_MILES_BACK_CALC:
            notes_parts.append(f"Miles back-calc from invoice history ({miles} mi)")
    else:
        # Unknown city — use 300 mi proxy and flag
        miles = 300
        flags.append(f"UNKNOWN DESTINATION: {city} not in mileage table. 300-mi proxy used.")

    if is_slc:
        flags.append("SLC BACKHAUL: Shipment originates Salt Lake City (unusual). Denver-routed mileage used as proxy. Flag for CFO review.")

    if imputed:
        flags.append("WEIGHT IMPUTED: weight_lbs was blank; imputed from units × avg lbs/unit.")

    rate    = peak_per_mile_rate(wt)
    base    = max(PEAK_MIN_CHARGE, round(miles * rate, 2))
    fuel    = round(base * PEAK_FSC_RATE, 2)
    acc, acc_detail = _parse_accessorials(row.get("special_handling", ""), PEAK_ACCESSORIALS)

    if base == PEAK_MIN_CHARGE:
        notes_parts.append("Minimum charge applied ($185)")

    zone_str = f"{miles} mi (wt tier: {'<500' if wt<500 else '500-2000' if wt<=2000 else '>2000'} lbs @ ${rate}/mi)"

    return ShipmentEstimate(
        shipment_id=row["shipment_id"], carrier="Peak Logistics",
        dest_city=city, dest_state=state, dest_zip=zip_, weight=wt,
        zone_or_region=zone_str,
        base_charge=base, fuel_surcharge=fuel,
        accessorial_total=acc, accessorial_detail=acc_detail,
        total_charge=round(base + fuel + acc, 2),
        notes="; ".join(notes_parts), flags=flags,
    )


def _estimate_heartland(row: pd.Series, imputed: bool) -> ShipmentEstimate:
    flags: list[str] = []
    notes_parts: list[str] = []

    city  = row["destination_city"]
    state = row["destination_state"]
    zip_  = int(row["destination_zip"])
    wt    = float(row["weight_lbs"])

    zone = _heartland_zone_from_zip(zip_)
    base = HEARTLAND_ZONE_RATES[zone]

    # Volume discount: April = Q2 week 1 → Tier 1 (0% discount, quarterly reset)
    discount = 0.0
    notes_parts.append("Q2 Tier 1 — 0% discount (quarter reset Apr 1)")

    if imputed:
        flags.append("WEIGHT IMPUTED: weight_lbs was blank; imputed from units × avg lbs/unit.")

    base = round(base * (1 - discount), 2)
    # FSC is INCLUDED in flat rate — do not add separate line
    fuel = 0.0
    acc, acc_detail = _parse_accessorials(row.get("special_handling", ""), HEARTLAND_ACCESSORIALS)

    zone_str = f"Zone {zone} — ${HEARTLAND_ZONE_RATES[zone]:.0f} flat (FSC incl.)"

    return ShipmentEstimate(
        shipment_id=row["shipment_id"], carrier="Heartland Freight",
        dest_city=city, dest_state=state, dest_zip=zip_, weight=wt,
        zone_or_region=zone_str,
        base_charge=base, fuel_surcharge=fuel,
        accessorial_total=acc, accessorial_detail=acc_detail,
        total_charge=round(base + acc, 2),
        notes="; ".join(notes_parts), flags=flags,
    )


def _estimate_coastal(row: pd.Series, imputed: bool) -> ShipmentEstimate:
    flags: list[str] = []
    notes_parts: list[str] = []

    city  = row["destination_city"]
    state = row["destination_state"]
    zip_  = int(row["destination_zip"])
    wt    = float(row["weight_lbs"])
    is_res = str(row.get("residential", "")).strip().upper() in ("TRUE", "1", "YES")

    region = _coastal_region(zip_)
    if region is None:
        region = "Unknown"
        flags.append(f"OUT-OF-TERRITORY: ZIP {zip_} ({city}, {state}) is not in Coastal Express service territory.")

    rate = COASTAL_REGION_RATES.get(region, 0.55)
    base = max(COASTAL_MIN_BASE, round(wt * rate, 2))
    fuel = round(base * COASTAL_FSC_RATE, 2)  # FSC on base only, NOT on accessorials

    if imputed:
        flags.append("WEIGHT IMPUTED: weight_lbs was blank; imputed from units × avg lbs/unit.")

    if base == COASTAL_MIN_BASE:
        notes_parts.append("Minimum base charge applied ($28)")

    # Residential surcharge (separate from accessorials table)
    res_fee = _coastal_res_surcharge(wt) if is_res else 0.0
    if is_res:
        notes_parts.append(f"Residential surcharge ${res_fee:.2f}")

    # Other accessorials (liftgate, appointment, inside delivery)
    acc_other, acc_detail = _parse_accessorials(row.get("special_handling", ""), COASTAL_ACCESSORIALS)
    acc_total = round(res_fee + acc_other, 2)
    if res_fee > 0:
        res_label = f"Residential ${res_fee:.2f}"
        acc_detail = (res_label + "; " + acc_detail).rstrip("; ") if acc_detail else res_label

    zone_str = f"{region} — ${rate}/lb (ZIP {zip_})"

    return ShipmentEstimate(
        shipment_id=row["shipment_id"], carrier="Coastal Express",
        dest_city=city, dest_state=state, dest_zip=zip_, weight=wt,
        zone_or_region=zone_str,
        base_charge=base, fuel_surcharge=fuel,
        accessorial_total=acc_total, accessorial_detail=acc_detail,
        total_charge=round(base + fuel + acc_total, 2),
        notes="; ".join(notes_parts), flags=flags,
    )


# ─── Adjustment buffer ────────────────────────────────────────────────────────

def _compute_adjustment_rate(invoices_df: pd.DataFrame) -> float:
    """
    Compute historical adjustment rate = sum(adjustments) / sum(pre-adj gross)
    across all 6 months. Adjustments are credits (negative), so rate is negative.
    Applied as a conservative haircut on the gross estimate.
    """
    gross = invoices_df["total_charge"].sum() - invoices_df["adjustments"].sum()
    adj   = invoices_df["adjustments"].sum()
    if gross == 0:
        return 0.0
    return adj / gross   # negative number, e.g. -0.003


# ─── Weight imputation ────────────────────────────────────────────────────────

def _impute_weights(df: pd.DataFrame) -> tuple[pd.DataFrame, set[str]]:
    """
    For rows with blank weight_lbs, impute as units × carrier-specific avg lbs/unit.
    Returns (modified_df, set_of_imputed_shipment_ids).
    """
    df = df.copy()
    # Normalise carrier first so we can group
    df["_carrier_norm"] = df["carrier"].apply(_normalise_carrier)

    # Compute avg lbs/unit per carrier from rows with valid weights
    mask_valid = df["weight_lbs"].notna() & (df["weight_lbs"] > 0)
    avg_lbs_per_unit = (
        df[mask_valid]
        .groupby("_carrier_norm")
        .apply(lambda g: (g["weight_lbs"] / g["units"]).median())
    )

    imputed_ids: set[str] = set()
    mask_missing = df["weight_lbs"].isna() | (df["weight_lbs"] == 0)
    for idx in df[mask_missing].index:
        carrier = df.at[idx, "_carrier_norm"]
        units   = df.at[idx, "units"]
        avg     = avg_lbs_per_unit.get(carrier, 20.0)
        df.at[idx, "weight_lbs"] = round(float(units) * float(avg), 1)
        imputed_ids.add(df.at[idx, "shipment_id"])

    df = df.drop(columns=["_carrier_norm"])
    return df, imputed_ids


# ─── Main engine ─────────────────────────────────────────────────────────────

def run_accrual_engine(
    shipments_csv: str,
    invoices_csv:  Optional[str] = None,
) -> dict:
    """
    Run the full accrual engine.

    Parameters
    ----------
    shipments_csv : path to shipments_apr2026.csv
    invoices_csv  : path to freight_invoices_oct2025_mar2026_v2.csv (optional;
                    used for back-calc mileage and adjustment buffer)

    Returns
    -------
    dict with keys: estimates, summary, flags, total_accrual, adjustment_rate
    """

    # ── Load shipments ────────────────────────────────────────────────────────
    df = pd.read_csv(shipments_csv, dtype={"destination_zip": str})
    df["destination_zip"] = pd.to_numeric(df["destination_zip"], errors="coerce").fillna(0).astype(int)
    df["weight_lbs"] = pd.to_numeric(df["weight_lbs"], errors="coerce")
    df["units"]      = pd.to_numeric(df["units"], errors="coerce").fillna(1).astype(int)

    # ── Carrier normalisation ─────────────────────────────────────────────────
    df["carrier"] = df["carrier"].apply(_normalise_carrier)

    # ── Duplicate detection ───────────────────────────────────────────────────
    # Exact duplicate = same shipment_id AND same content
    dup_flags: list[dict] = []
    key_cols = ["shipment_id", "date", "origin_city", "destination_city",
                "destination_zip", "carrier", "weight_lbs"]
    before = len(df)
    df_dedup = df.drop_duplicates(subset=key_cols, keep="first")
    removed = before - len(df_dedup)
    if removed:
        # Identify which shipment_ids were duped
        counts = df.groupby("shipment_id").size()
        duped = counts[counts > 1].index.tolist()
        for sid in duped:
            dup_flags.append({
                "shipment": sid,
                "carrier": df.loc[df["shipment_id"] == sid, "carrier"].iloc[0],
                "flag": f"DUPLICATE: shipment {sid} appears {counts[sid]}× in 3PL log. Second occurrence excluded.",
            })
    df = df_dedup.reset_index(drop=True)

    # ── Weight imputation ──────────────────────────────────────────────────────
    df, imputed_ids = _impute_weights(df)

    # ── Adjustment buffer from invoice history ─────────────────────────────────
    adjustment_rate = 0.0
    if invoices_csv and os.path.exists(invoices_csv):
        inv_df = pd.read_csv(invoices_csv)
        adjustment_rate = _compute_adjustment_rate(inv_df)

    # ── Estimate each shipment ────────────────────────────────────────────────
    estimates:    list[ShipmentEstimate] = []
    global_flags: list[dict]             = list(dup_flags)

    for _, row in df.iterrows():
        carrier  = row["carrier"]
        imputed  = row["shipment_id"] in imputed_ids

        try:
            if carrier == "Peak Logistics":
                est = _estimate_peak(row, imputed)
            elif carrier == "Heartland Freight":
                est = _estimate_heartland(row, imputed)
            elif carrier == "Coastal Express":
                est = _estimate_coastal(row, imputed)
            else:
                # Unknown carrier — flag and skip
                global_flags.append({
                    "shipment": row["shipment_id"],
                    "carrier": carrier,
                    "flag": f"UNKNOWN CARRIER '{carrier}' — no rate card available. Excluded from accrual.",
                })
                continue
        except Exception as exc:
            global_flags.append({
                "shipment": row["shipment_id"],
                "carrier": carrier,
                "flag": f"ESTIMATION ERROR: {exc}",
            })
            continue

        estimates.append(est)

        # Bubble up shipment-level flags to global list
        for f in est.flags:
            global_flags.append({
                "shipment": est.shipment_id,
                "carrier":  est.carrier,
                "flag":     f,
            })

    # ── Per-carrier rollup ────────────────────────────────────────────────────
    carriers_ordered = ["Peak Logistics", "Heartland Freight", "Coastal Express"]
    summary: dict[str, dict] = {}

    for carrier in carriers_ordered:
        subset = [e for e in estimates if e.carrier == carrier]
        gross_base  = sum(e.base_charge for e in subset)
        gross_fuel  = sum(e.fuel_surcharge for e in subset)
        gross_acc   = sum(e.accessorial_total for e in subset)
        gross_total = sum(e.total_charge for e in subset)

        # Apply adjustment buffer (negative haircut)
        adj = round(gross_total * adjustment_rate, 2)
        net = round(gross_total + adj, 2)

        summary[carrier] = {
            "shipment_count":     len(subset),
            "base_total":         round(gross_base, 2),
            "fuel_total":         round(gross_fuel, 2),
            "accessorial_total":  round(gross_acc, 2),
            "gross_total":        round(gross_total, 2),
            "adjustment":         adj,
            "net_accrual":        net,
        }

    total_accrual = round(sum(s["net_accrual"] for s in summary.values()), 2)

    return {
        "estimates":       estimates,
        "summary":         summary,
        "flags":           global_flags,
        "total_accrual":   total_accrual,
        "adjustment_rate": adjustment_rate,
    }
