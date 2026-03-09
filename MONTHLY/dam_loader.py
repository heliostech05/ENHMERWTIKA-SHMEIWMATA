#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
dam_loader.py — Φόρτωση τιμών DAM (Day-Ahead Market) από Energy Charts CSV.

Χειρίζεται:
- Εύρεση header γραμμής (disclaimer detection)
- Αυτόματη αναγνώριση στηλών timestamp/price
- DST-safe duplicate handling
- Φιλτράρισμα μόνο 15-λεπτων δεδομένων (≥2025-10-01)
"""

import logging
import pandas as pd
from .config import DAM_QUARTER_CUTOFF

log = logging.getLogger("MONTHLY.dam_loader")

# ===== Header detection keywords =====
HEADER_TS_KEYS = [
    "date", "time", "timestamp", "cet", "ce(s)t",
    "gmt", "utc", "eet", "athens", "gmt+2",
]
HEADER_PRICE_KEYS = [
    "price", "eur/mwh", "€/mwh", "auction", "day-ahead", "day ahead",
]


def _find_header_line(path, max_scan=200):
    """Βρίσκει τη γραμμή header στο CSV (παρακάμπτοντας disclaimer γραμμές)."""
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        for i in range(max_scan):
            line = f.readline()
            if not line:
                break
            low = line.strip().lower()
            if (any(k in low for k in HEADER_TS_KEYS) and
                    any(k in low for k in HEADER_PRICE_KEYS)):
                return i
    return 1


def _infer_dam_columns(df: pd.DataFrame):
    """
    Αυτόματη αναγνώριση στηλών Timestamp και Price στο DAM DataFrame.

    Πρώτα ψάχνει βάσει keywords στα ονόματα στηλών.
    Αν δεν βρει, κάνει fallback με heuristic (parse rate ≥80%).
    """
    cols = list(df.columns)
    lower = {c: c.lower() for c in cols}

    ts_cands = [c for c in cols if any(k in lower[c] for k in HEADER_TS_KEYS)]
    price_cands = [c for c in cols if any(k in lower[c] for k in HEADER_PRICE_KEYS)]

    ts_col = ts_cands[0] if ts_cands else None
    price_cands = [c for c in price_cands if c != ts_col]
    price_col = price_cands[0] if price_cands else None

    # Fallback: heuristic detection
    if not ts_col:
        best, best_rate = None, -1
        for c in cols:
            try:
                parsed = pd.to_datetime(df[c], errors="coerce")
                rate = parsed.notna().sum() / max(1, df[c].notna().sum())
                if rate >= 0.8 and rate > best_rate:
                    best, best_rate = c, rate
            except Exception:
                pass
        ts_col = best

    if not price_col:
        best, best_rate = None, -1
        for c in cols:
            if c == ts_col:
                continue
            s = pd.to_numeric(
                df[c].astype(str).str.replace(",", ".", regex=False),
                errors="coerce"
            )
            rate = s.notna().sum() / max(1, df[c].notna().sum())
            if rate >= 0.8 and rate > best_rate:
                best, best_rate = c, rate
        price_col = best

    if not ts_col or not price_col:
        raise ValueError(
            f"Δεν βρέθηκαν στήλες Timestamp/Price στο DAM CSV. Columns: {list(df.columns)}"
        )
    return ts_col, price_col


def load_dam_quarterly_endtime(dam_csv_path: str, month: str):
    """
    Διαβάζει το Energy Charts CSV, βρίσκει header, θεωρεί ότι το timestamp είναι
    ΗΔΗ local START time ανά 15λεπτο (00:00, 00:15, ..., 23:45) και
    ΔΕΝ το μετακινεί -15'.

    Επιστρέφει DataFrame με: TIMESTAMP (local START), DAM Price (€/MWh), dup_idx
    """
    try:
        header_line = _find_header_line(dam_csv_path)
        dam = pd.read_csv(
            dam_csv_path, sep=None, engine="python",
            encoding="utf-8-sig", header=header_line,
        )
        dam = dam.loc[:, ~dam.columns.astype(str).str.fullmatch(r"Unnamed: \d+")]
        dam.columns = [str(c).strip() for c in dam.columns]

        ts_col, price_col = _infer_dam_columns(dam)

        # Parse as UTC-if-possible, then convert to Europe/Athens
        ts_aware = pd.to_datetime(dam[ts_col], errors="coerce", utc=True)
        if ts_aware.isna().all():
            start_local = pd.to_datetime(dam[ts_col], errors="coerce")  # naive local
        else:
            start_local = ts_aware.dt.tz_convert("Europe/Athens").dt.tz_localize(None)

        price = pd.to_numeric(
            dam[price_col].astype(str).str.replace(",", ".", regex=False),
            errors="coerce",
        )

        out = pd.DataFrame({
            "TIMESTAMP": start_local,
            "DAM Price (€/MWh)": price,
        }).dropna(subset=["TIMESTAMP"])

        # Φιλτράρισμα: μόνο ≥ cutoff (15-λεπτα δεδομένα) και μόνο τον ζητούμενο μήνα
        lb = pd.Timestamp(DAM_QUARTER_CUTOFF)
        out = out[out["TIMESTAMP"] >= lb]
        out = out[out["TIMESTAMP"].dt.strftime("%Y-%m") == month].copy()

        # ΔΕΝ κάνουμε sort: κρατάμε τη σειρά αρχείου, dup_idx για DST fallback
        out["dup_idx"] = out.groupby("TIMESTAMP").cumcount()

        log.info("DAM 15' rows after filters: %d for %s", len(out), month)
        return out.reset_index(drop=True)

    except Exception as e:
        log.error("Failed to load DAM prices: %s", e)
        return None
