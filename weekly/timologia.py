#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
timologia.py

Weekly orchestrator για ΣΗΘΥΑ, χτισμένο πάνω στο MONTHLY pipeline.

Η μόνη ουσιαστική διαφοροποίηση είναι:
- χρησιμοποιούμε την ήδη υπάρχουσα μηνιαία λογική για producers, DAM,
  παραγωγή και ημερήσια σύνοψη
- φιλτράρουμε μόνο τις ημέρες της ζητούμενης εβδομάδας
- εξάγουμε με weekly template / ονομασία
"""

import logging
import os
import re
from pathlib import Path

import pandas as pd

from MONTHLY.config import (
    PRODUCERS_PATH,
    PRODUCTION_DIR,
    dam_file_for_year,
)
from MONTHLY.dam_loader import load_dam_quarterly_endtime
from MONTHLY.helpers import normalize_name, sanitize_name
from MONTHLY.production import (
    build_email_groups,
    calculate_daily_summary_quarterly_for_days,
    filter_monthly_data,
    load_producers,
    read_production_data,
)
from admie_merged_production import split_files_by_code

from .excel_export import (
    determine_pdf_subfolder_name,
    export_to_pdf,
    generate_invoice_excel_weekly,
    make_week_dirs,
    pdf_filename_weekly,
)

log = logging.getLogger("weekly.timologia")

SITHYA_ALIASES = {
    "ΣΗΘΥΑ", "ΣΗΘ", "ΣΗΘΥΑ/ΣΗΘ", "ΣΗΘΥΑ - CHP", "CHP", "ΣΗΘ-YA", "ΒΙΟΑΕΡΙΟ",
}
ISO_WEEK_RE = re.compile(r"^(\d{4})-W(0[1-9]|[1-4]\d|5[0-3])$", re.IGNORECASE)


def load_producers_sithya(filepath=PRODUCERS_PATH):
    """Φορτώνει το producers.xlsx και κρατά μόνο παραγωγούς ΣΗΘΥΑ/CHP."""
    df = load_producers(str(filepath))
    if df is None:
        return None

    for col in ["IBAN", "Code", "Τεχνολογία"]:
        if col not in df.columns:
            df[col] = ""

    tech = df["Τεχνολογία"].astype(str).str.strip().str.upper()
    mask = tech.isin(SITHYA_ALIASES) | tech.str.contains(r"\bΣΗΘΥΑ\b", regex=True)
    df = df[mask].copy()

    if df.empty:
        log.warning("Δεν βρέθηκαν παραγωγοί ΣΗΘΥΑ.")
        return df

    if "normalized_name" not in df.columns:
        df["normalized_name"] = df["Εταιρεία"].astype(str).apply(normalize_name)

    log.info("Loaded %d ΣΗΘΥΑ producers", len(df))
    return df


def _month_iter(start_date: pd.Timestamp, end_date: pd.Timestamp):
    cur = pd.Timestamp(start_date.year, start_date.month, 1)
    final = pd.Timestamp(end_date.year, end_date.month, 1)
    while cur <= final:
        yield f"{cur.year}-{cur.month:02d}"
        cur = cur + pd.offsets.MonthBegin(1)


def _resolve_week_range(start_date_or_week: str, end_date_str: str | None = None):
    """
    Δέχεται:
    - ISO week: YYYY-Www
    - ή start/end dates: YYYY-MM-DD, YYYY-MM-DD

    Επιστρέφει πάντα ακριβές διάστημα 7 ημερών, Δευτέρα έως Κυριακή.
    """
    start_value = (start_date_or_week or "").strip()
    end_value = (end_date_str or "").strip()

    iso_match = ISO_WEEK_RE.match(start_value)
    if iso_match and not end_value:
        iso_year = int(iso_match.group(1))
        iso_week = int(iso_match.group(2))
        try:
            start_date = pd.Timestamp.fromisocalendar(iso_year, iso_week, 1).floor("D")
            end_date = pd.Timestamp.fromisocalendar(iso_year, iso_week, 7).floor("D")
        except ValueError as exc:
            raise ValueError(f"Μη έγκυρο ISO week: {start_value}") from exc
        return start_date, end_date

    try:
        start_date = pd.to_datetime(start_value).floor("D")
        end_date = pd.to_datetime(end_value).floor("D")
    except Exception as exc:
        raise ValueError("Μη έγκυρες ημερομηνίες. Χρήση: YYYY-MM-DD ή ISO week YYYY-Www") from exc

    if end_date < start_date:
        raise ValueError("Το τέλος είναι πριν την αρχή.")

    if start_date.weekday() != 0 or end_date.weekday() != 6 or (end_date - start_date).days != 6:
        raise ValueError("Το weekly flow απαιτεί ακριβώς 7 ημέρες, από Δευτέρα έως Κυριακή.")

    return start_date, end_date


def ensure_production_files(start_date: pd.Timestamp, end_date: pd.Timestamp):
    """
    Επαναχρησιμοποιεί το υπάρχον production split pipeline.

    Για κάθε μήνα που καλύπτεται από το διάστημα, τρέχει το ίδιο splitting
    που χρησιμοποιεί και η monthly ροή.
    """
    processed_months = 0
    for month in _month_iter(start_date, end_date):
        downloads_dir = Path("downloads") / month
        if not downloads_dir.is_dir():
            log.info("skip split_files_by_code: no downloads for %s", month)
            continue
        print(f"🔧 Ενημέρωση ΠΑΡΑΓΩΓΗ από downloads/{month} ...")
        split_files_by_code(month)
        processed_months += 1

    if processed_months:
        print("✅ Έτοιμο το ΠΑΡΑΓΩΓΗ/ για το επιλεγμένο διάστημα.")
    else:
        print("ℹ️ Δεν βρέθηκαν διαθέσιμα downloads για ενημέρωση του ΠΑΡΑΓΩΓΗ/.")


def _find_production_file_for_producer(company_name: str) -> Path | None:
    """
    Εντοπίζει το αντίστοιχο ΠΑΡΑΓΩΓΗ_{εταιρεία}.csv.

    Ξεκινά με το canonical filename και αν χρειαστεί κάνει fallback scan
    βάσει normalized name.
    """
    safe_name = sanitize_name(company_name).replace(" ", "_")
    direct = PRODUCTION_DIR / f"ΠΑΡΑΓΩΓΗ_{safe_name}.csv"
    if direct.exists():
        return direct

    company_norm = normalize_name(company_name)
    if not PRODUCTION_DIR.is_dir():
        return None

    for filename in os.listdir(PRODUCTION_DIR):
        if not (filename.startswith("ΠΑΡΑΓΩΓΗ_") and filename.endswith(".csv")):
            continue
        match = re.match(r"ΠΑΡΑΓΩΓΗ_(.+)\.csv", filename)
        if not match:
            continue
        if normalize_name(match.group(1)) == company_norm:
            return PRODUCTION_DIR / filename

    return None


def _filter_production_for_period(
    df_prod_raw: pd.DataFrame,
    start_date: pd.Timestamp,
    end_date: pd.Timestamp,
):
    parts = []
    for month_tag in _month_iter(start_date, end_date):
        filtered = filter_monthly_data(df_prod_raw, month_tag)
        if filtered is not None and not filtered.empty:
            parts.append(filtered)

    if not parts:
        return pd.DataFrame()

    return pd.concat(parts, ignore_index=True)


def _load_dam_for_period(start_date: pd.Timestamp, end_date: pd.Timestamp):
    parts = []
    for month_tag in _month_iter(start_date, end_date):
        year = int(month_tag[:4])
        dam_file = dam_file_for_year(year)
        if not dam_file.exists():
            raise FileNotFoundError(f"Λείπει DAM CSV: {dam_file.name}")
        df_month = load_dam_quarterly_endtime(str(dam_file), month_tag)
        if df_month is not None and not df_month.empty:
            parts.append(df_month)

    if not parts:
        return pd.DataFrame()

    return pd.concat(parts, ignore_index=True)

def timologia_weekly(start_date_or_week: str, end_date_str: str | None = None):
    try:
        start_date, end_date = _resolve_week_range(start_date_or_week, end_date_str)
    except ValueError as exc:
        print(str(exc))
        return
    tag, root, xlsx_dir, pdf_dir = make_week_dirs(start_date, end_date)

    producers_df = load_producers_sithya(PRODUCERS_PATH)
    if producers_df is None or producers_df.empty:
        print("Δεν βρέθηκαν παραγωγοί ΣΗΘΥΑ.")
        return

    email_to_companies, _ = build_email_groups(producers_df)

    try:
        df_dam_period = _load_dam_for_period(start_date, end_date)
    except FileNotFoundError as exc:
        print(str(exc))
        return
    if df_dam_period is None or df_dam_period.empty:
        print("Αποτυχία: DAM 15' prices (empty).")
        return
    log.info("Weekly period %s -> %s | ISO %s", start_date.date(), end_date.date(), tag)

    if not PRODUCTION_DIR.is_dir():
        print(f"Λείπει φάκελος: {PRODUCTION_DIR}")
        return

    processed = 0
    skipped_no_file = 0
    skipped_errors = 0

    for _, prod_row_series in producers_df.iterrows():
        prod_row = producers_df.loc[[prod_row_series.name]]
        company_name = str(prod_row_series.get("Εταιρεία", "")).strip()
        if not company_name:
            continue

        file_path = _find_production_file_for_producer(company_name)
        if file_path is None or not file_path.exists():
            skipped_no_file += 1
            print(f"\n=== {company_name} ===")
            print("  -> SKIP: Δεν βρέθηκε αρχείο παραγωγής (ΠΑΡΑΓΩΓΗ_*.csv)")
            continue

        print(f"\n=== Επεξεργασία παραγωγού: {company_name} ===")
        print(f"  Production file: {file_path.name}")

        df = read_production_data(str(file_path))
        if df is None:
            skipped_errors += 1
            print("  -> SKIP: read_production_data επέστρεψε None")
            continue

        df = _filter_production_for_period(df, start_date, end_date)
        if df.empty:
            skipped_errors += 1
            print("  -> SKIP: Δεν υπάρχουν γραμμές παραγωγής για το επιλεγμένο εβδομαδιαίο διάστημα")
            continue

        days = pd.date_range(start_date, end_date, freq="D")
        df_week, summary = calculate_daily_summary_quarterly_for_days(
            df, df_dam_period, prod_row, days
        )
        if df_week is None:
            skipped_errors += 1
            print("  -> SKIP: calculate_daily_summary_quarterly_for_days επέστρεψε None")
            continue

        xlsx_path, company_name_out, email_value = generate_invoice_excel_weekly(
            df_week, summary, prod_row, start_date, end_date, xlsx_dir, tag
        )
        if not xlsx_path:
            skipped_errors += 1
            print("  -> SKIP: generate_invoice_excel_weekly απέτυχε")
            continue

        email_key = (email_value or "NO_EMAIL").strip() or "NO_EMAIL"
        subfolder = determine_pdf_subfolder_name(email_key, email_to_companies)
        target_dir = pdf_dir / subfolder
        target_dir.mkdir(parents=True, exist_ok=True)

        pdf_name = pdf_filename_weekly(company_name_out, tag)
        pdf_path = target_dir / pdf_name

        ok, how = export_to_pdf(xlsx_path, str(pdf_path))
        status = "✅ PDF" if ok else "❌ PDF"
        print(f"{status} [{how}] → {pdf_path}")

        processed += 1

    print("\n" + "=" * 70)
    print("Ολοκληρώθηκε.")
    print(f"Processed: {processed}")
    print(f"Skipped (no production file): {skipped_no_file}")
    print(f"Skipped (errors): {skipped_errors}")
    print(f"Δες: {root}/XLSX και {root}/PDF")


def main():
    start_or_week = input("Δώσε ISO week (YYYY-Www) ή αρχή εβδομάδας (YYYY-MM-DD): ").strip()
    end = input("Δώσε τέλος εβδομάδας (YYYY-MM-DD) ή άφησέ το κενό αν έδωσες ISO week: ").strip()
    timologia_weekly(start_or_week, end or None)


if __name__ == "__main__":
    main()
