#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
production.py — Φόρτωση, φιλτράρισμα και υπολογισμός ημερήσιων συνόψεων παραγωγής.

Περιλαμβάνει:
- Φόρτωση producers.xlsx και ομαδοποίηση ανά email
- Ανάγνωση CSV παραγωγής
- Φιλτράρισμα κατά μήνα
- Υπολογισμός ημερήσιων συνόψεων (pairing παραγωγής–DAM τιμών)
"""

import logging
import pandas as pd
from collections import defaultdict

from .config import DST_SKIP_DATE, DST_FALLBACK_DATE
from .helpers import normalize_name

log = logging.getLogger("MONTHLY.production")


# ===== Producers / Email groups =====

def load_producers(filepath="producers.xlsx"):
    """Φόρτωση producers.xlsx με validation στηλών."""
    try:
        df = pd.read_excel(filepath, dtype={'Code': str})

        if 'Εταιρεία' not in df.columns:
            raise ValueError("Λείπει στήλη 'Εταιρεία'")
        if 'Μοναδιαία Χρέωση ΦοΣΕ' not in df.columns:
            raise ValueError("Λείπει στήλη 'Μοναδιαία Χρέωση ΦοΣΕ'")
        if 'Email' not in df.columns:
            df['Email'] = ""
        if 'Όνομα Φακέλου' not in df.columns:
            df['Όνομα Φακέλου'] = ""

        df['normalized_name'] = df['Εταιρεία'].astype(str).apply(normalize_name)
        log.info("Loaded %d producers", len(df))
        return df

    except Exception as e:
        log.error("Failed to load producers: %s", e)
        return None


def build_email_groups(producers_df):
    """Ομαδοποίηση εταιρειών ανά email (για PDF subfolder naming)."""
    email_to_companies = defaultdict(set)
    email_to_customs = defaultdict(set)

    for _, row in producers_df.iterrows():
        email = (str(row.get('Email', '') or '').strip()) or "NO_EMAIL"
        comp = str(row.get('Εταιρεία', '') or '').strip()
        cust = str(row.get('Όνομα Φακέλου', '') or '').strip()
        if comp:
            email_to_companies[email].add(comp)
        if cust:
            email_to_customs[email].add(cust)

    email_to_companies = {em: sorted(v) for em, v in email_to_companies.items()}
    email_to_customs = {em: sorted(v) for em, v in email_to_customs.items()}
    return email_to_companies, email_to_customs


# ===== Production data reading =====

def read_production_data(file_path):
    """Ανάγνωση CSV παραγωγής (auto-detect separator)."""
    try:
        df = pd.read_csv(file_path, sep=None, engine="python", encoding="utf-8-sig")
        log.debug("Read %s: %d rows", file_path, len(df))
        return df
    except Exception as e:
        log.error("Failed to read %s: %s", file_path, e)
        return None


def filter_monthly_data(df, month):
    """
    Φιλτράρισμα γραμμών παραγωγής για τον μήνα 'YYYY-MM'.

    Εφαρμόζει ειδικούς κανόνες:
    - Πετάει 00:15/00:30/00:45 για DST_SKIP_DATE
    - Κρατάει μόνο πρώτη εμφάνιση σε duplicates
    """
    if 'TIMESTAMP' not in df.columns or 'Μήνας' not in df.columns:
        log.warning("Missing TIMESTAMP/Μήνας columns")
        return pd.DataFrame()

    df = df[df['Μήνας'] == month].copy()
    if df.empty:
        log.debug("No rows for month %s", month)
        return df

    ts = pd.to_datetime(
        df['TIMESTAMP'],
        format="%d/%m/%Y %H:%M",
        errors='coerce',
        dayfirst=True,
    )
    df = df[ts.notna()].copy()
    df['TIMESTAMP'] = ts[ts.notna()]

    df = df[df['TIMESTAMP'].dt.minute.isin([0, 15, 30, 45])].copy()

    # DST_SKIP_DATE: πετάμε START 00:15/00:30/00:45
    dst_skip = pd.Timestamp(DST_SKIP_DATE).date()
    mask_skip = (
        (df['TIMESTAMP'].dt.date == dst_skip) &
        (df['TIMESTAMP'].dt.hour == 0) &
        (df['TIMESTAMP'].dt.minute.isin([15, 30, 45]))
    )
    if mask_skip.any():
        df = df[~mask_skip].copy()

    df = df.sort_index(kind='stable')
    dup_mask = df.duplicated(subset=['TIMESTAMP'], keep='first')
    if dup_mask.any():
        df = df[~dup_mask].copy()

    log.info("Filtered %d rows for month=%s", len(df), month)
    return df


# ===== Daily summary calculation =====

def _prepare_production_for_summary(df_prod):
    prod = df_prod.copy()
    prod['END_TS'] = pd.to_datetime(
        prod['TIMESTAMP'],
        format="%d/%m/%Y %H:%M",
        errors='coerce',
        dayfirst=True,
    )
    prod = prod.dropna(subset=['END_TS'])
    if prod.empty:
        return None

    prod['ΕΝΕΡΓΕΙΑ (kWh)'] = pd.to_numeric(
        prod['ΕΝΕΡΓΕΙΑ (kWh)'], errors='coerce'
    ).fillna(0.0)
    return prod.sort_values('END_TS').reset_index(drop=True)


def _prepare_dam_for_summary(df_dam_15m, month: str | None = None):
    dam = df_dam_15m.copy()
    dam['START_TS'] = pd.to_datetime(dam['TIMESTAMP'], errors='coerce')
    dam = dam.dropna(subset=['START_TS'])
    if month is not None:
        dam = dam[dam['START_TS'].dt.strftime("%Y-%m") == month].copy()
    if dam.empty:
        return None
    return dam.sort_values('START_TS').reset_index(drop=True)


def _calculate_daily_summary_for_days(prod, dam, producer_row, days):
    dst_skip = pd.Timestamp(DST_SKIP_DATE).date()
    dst_fallback = pd.Timestamp(DST_FALLBACK_DATE).date()

    all_quarters = []

    for D in days:
        D_ts = pd.Timestamp(str(D))
        D_next = D_ts + pd.Timedelta(days=1)

        # Prod(END): D 00:15..23:45 + (D+1) 00:00
        day_prod = prod[
            (prod['END_TS'] > D_ts) & (prod['END_TS'] <= D_next)
        ].copy()

        if D == dst_skip:
            mask_skip = (
                (day_prod['END_TS'].dt.date == D) &
                (day_prod['END_TS'].dt.hour == 0) &
                (day_prod['END_TS'].dt.minute.isin([15, 30, 45]))
            )
            day_prod = day_prod[~mask_skip].copy()

        if D == dst_fallback:
            mask_win = (
                ((day_prod['END_TS'].dt.hour == 3) &
                 day_prod['END_TS'].dt.minute.isin([15, 30, 45])) |
                ((day_prod['END_TS'].dt.hour == 4) &
                 (day_prod['END_TS'].dt.minute == 0))
            )
            dup = day_prod[mask_win].duplicated(subset=['END_TS'], keep='first')
            day_prod = day_prod.drop(index=day_prod[mask_win].loc[dup].index)

        day_prod = day_prod.sort_values('END_TS').reset_index(drop=True)
        if day_prod.empty:
            continue

        # DAM(START): D 00:00..23:45
        day_dam = dam[
            (dam['START_TS'] >= D_ts) &
            (dam['START_TS'] <= D_ts + pd.Timedelta(hours=23, minutes=45))
        ].copy()

        if D == dst_fallback:
            mask_dam_win = (
                (day_dam['START_TS'].dt.hour == 3) &
                (day_dam['START_TS'].dt.minute.isin([0, 15, 30, 45]))
            )
            dup_dam = day_dam[mask_dam_win].duplicated(
                subset=['START_TS'], keep='first'
            )
            day_dam = day_dam.drop(
                index=day_dam[mask_dam_win].loc[dup_dam].index
            )

        day_dam = day_dam.sort_values('START_TS').reset_index(drop=True)
        if day_dam.empty:
            continue

        n_p = len(day_prod)
        n_d = len(day_dam)
        n = min(n_p, n_d)
        if n == 0:
            continue
        if n_p != n_d:
            log.warning("Length mismatch on %s: prod=%d, dam=%d, using first %d", D, n_p, n_d, n)

        day_prod = day_prod.iloc[:n].copy()
        day_dam = day_dam.iloc[:n].copy()

        kwh = day_prod['ΕΝΕΡΓΕΙΑ (kWh)'].to_numpy()
        price = day_dam['DAM Price (€/MWh)'].to_numpy()
        value_eur = (kwh * price) / 1000.0

        per_quarter = pd.DataFrame({
            'Περίοδος εκκαθάρισης': [D] * n,
            'ΕΝΕΡΓΕΙΑ (kWh)': kwh,
            'Αξία ενέργειας βάσει μετρήσεων (€)': value_eur,
        })
        all_quarters.append(per_quarter)

    if not all_quarters:
        return None, None

    df_all = pd.concat(all_quarters, ignore_index=True)

    df_daily = df_all.groupby('Περίοδος εκκαθάρισης', as_index=False).agg({
        'ΕΝΕΡΓΕΙΑ (kWh)': 'sum',
        'Αξία ενέργειας βάσει μετρήσεων (€)': 'sum',
    })

    df_daily['Περίοδος εκκαθάρισης'] = pd.to_datetime(
        df_daily['Περίοδος εκκαθάρισης']
    ).dt.strftime('%d/%m/%y')

    rate = float(producer_row['Μοναδιαία Χρέωση ΦοΣΕ'].values[0])
    df_daily['Προμήθεια GREEN VALUE (€)'] = (
        df_daily['ΕΝΕΡΓΕΙΑ (kWh)'] / 1000.0 * rate
    ).round(2)

    df_daily['Μεσοσταθμική Τιμή Αγοράς κατά τις ώρες παραγωγής του σταθμού'] = df_daily.apply(
        lambda row: 0 if row['ΕΝΕΡΓΕΙΑ (kWh)'] == 0 else round(
            (row['Αξία ενέργειας βάσει μετρήσεων (€)'] / row['ΕΝΕΡΓΕΙΑ (kWh)']) * 1000.0, 2
        ),
        axis=1,
    )

    sum_energy = round(float(df_daily['ΕΝΕΡΓΕΙΑ (kWh)'].sum()), 2)
    sum_value = round(float(df_daily['Αξία ενέργειας βάσει μετρήσεων (€)'].sum()), 2)
    sum_prov = round(float(df_daily['Προμήθεια GREEN VALUE (€)'].sum()), 2)

    summary_row = pd.DataFrame([{
        'Περίοδος εκκαθάρισης': 'Σύνολο',
        'ΕΝΕΡΓΕΙΑ (kWh)': sum_energy,
        'Αξία ενέργειας βάσει μετρήσεων (€)': sum_value,
        'Προμήθεια GREEN VALUE (€)': sum_prov,
        'Μεσοσταθμική Τιμή Αγοράς κατά τις ώρες παραγωγής του σταθμού': (
            round((sum_value / sum_energy) * 1000.0, 2) if sum_energy > 0 else 0
        ),
    }])

    df_final = pd.concat([df_daily, summary_row], ignore_index=True)
    return df_final, (sum_energy, sum_value, sum_prov)


def calculate_daily_summary_quarterly_for_days(df_prod, df_dam_15m, producer_row, days):
    """
    Ίδια ακριβώς λογική με calculate_daily_summary_quarterly, αλλά για ρητή λίστα ημερών.
    Χρήσιμο για weekly flow χωρίς υπολογισμό άσχετων ημερών του μήνα.
    """
    try:
        prod = _prepare_production_for_summary(df_prod)
        if prod is None:
            return None, None

        dam = _prepare_dam_for_summary(df_dam_15m)
        if dam is None:
            return None, None

        day_list = [pd.Timestamp(d).date() for d in days]
        if not day_list:
            return None, None

        return _calculate_daily_summary_for_days(prod, dam, producer_row, day_list)

    except Exception as e:
        log.error("Daily summary calculation failed for explicit days: %s", e)
        return None, None

def calculate_daily_summary_quarterly(df_prod, df_dam_15m, producer_row, month):
    """
    Υπολογισμός ημερήσιων συνόψεων — pairing ανά index:
      - Prod(END): D 00:15..23:45 + (D+1) 00:00
      - DAM(START): D 00:00..23:45
      - P[i] ↔ DAM[i] χωρίς shift.

    DST_SKIP_DATE: πετάμε τις 00:15/00:30/00:45
    DST_FALLBACK_DATE: κρατάμε μόνο την πρώτη εμφάνιση στα διπλά 03:00–04:00
    """
    try:
        month_str = month
        prod = _prepare_production_for_summary(df_prod)
        # Σημαντικό: δεν φιλτράρουμε ξανά με βάση END_TS μήνα.
        # Το df_prod έχει ήδη φιλτραριστεί από filter_monthly_data μέσω της στήλης "Μήνας",
        # ώστε να κρατά και το 00:00 της 1ης επόμενης ημέρας ως τελευταίο 15λεπτο
        # της τελευταίας ημέρας του μήνα (π.χ. 01/02 00:00 -> 2026-01).
        if prod is None:
            log.info("No production rows for %s", month_str)
            return None, None

        dam = _prepare_dam_for_summary(df_dam_15m, month=month_str)
        if dam is None:
            log.info("No DAM rows for %s", month_str)
            return None, None

        days = sorted({d for d in prod['END_TS'].dt.date if str(d).startswith(month_str)})
        if not days:
            log.info("No days in production for %s", month_str)
            return None, None

        df_final, summary = _calculate_daily_summary_for_days(prod, dam, producer_row, days)
        if df_final is None:
            log.info("No quarter rows after pairing for %s", month_str)
            return None, None

        log.info("Daily summary: %d rows for month=%s", len(df_final), month_str)
        return df_final, summary

    except Exception as e:
        log.error("Daily summary calculation failed: %s", e)
        return None, None
