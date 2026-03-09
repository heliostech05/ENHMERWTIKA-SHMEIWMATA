#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
timologia.py — Κεντρικός orchestrator μηνιαίων ενημερωτικών σημειωμάτων.

Χρήση:
    python -m MONTHLY
    ή
    from MONTHLY.timologia import timologia
    timologia("2026-01")
"""

import logging
import os
import re

from .config import (
    PRODUCERS_PATH, PRODUCTION_DIR, MAX_FOLDER_CHARS,
    dam_file_for_year,
)
from .helpers import normalize_name
from .dam_loader import load_dam_quarterly_endtime
from .production import (
    load_producers, build_email_groups,
    read_production_data, filter_monthly_data,
    calculate_daily_summary_quarterly,
)
from .excel_export import (
    make_base_dirs, determine_pdf_subfolder_name,
    generate_invoice_excel, export_to_pdf, pdf_filename,
)

log = logging.getLogger("MONTHLY.timologia")


def timologia(month):
    """
    Κύρια ροή: για κάθε αρχείο παραγωγής στο ΠΑΡΑΓΩΓΗ/,
    δημιουργεί ενημερωτικό σημείωμα (XLSX + PDF).
    """
    # 1. Φόρτωση producers
    producers_df = load_producers(str(PRODUCERS_PATH))
    if producers_df is None:
        print("Αποτυχία: producers.xlsx")
        return

    email_to_companies, email_to_customs = build_email_groups(producers_df)

    # 2. Φόρτωση DAM τιμών (αυτόματη επιλογή έτους από τον μήνα)
    year = int(month.split('-')[0])
    dam_file_path = dam_file_for_year(year)
    dam_file = str(dam_file_path)
    if not dam_file_path.exists():
        print(f"Λείπει αρχείο DAM: {dam_file}")
        return

    df_dam_15m = load_dam_quarterly_endtime(dam_file, month)
    if df_dam_15m is None or df_dam_15m.empty:
        print("Αποτυχία: DAM 15' prices")
        return

    # 3. Δημιουργία φακέλων εξόδου
    root, xlsx_dir, pdf_dir = make_base_dirs(month)

    # 4. Επεξεργασία κάθε αρχείου παραγωγής
    base_folder = str(PRODUCTION_DIR)
    if not os.path.isdir(base_folder):
        print(f"Λείπει φάκελος: {base_folder}")
        return

    for filename in os.listdir(base_folder):
        if not (filename.startswith('ΠΑΡΑΓΩΓΗ_') and filename.endswith('.csv')):
            continue

        print(f"\n=== Επεξεργασία αρχείου: {filename} ===")

        file_path = os.path.join(base_folder, filename)
        m = re.match(r'ΠΑΡΑΓΩΓΗ_(.+)\.csv', filename)
        if not m:
            log.warning("Bad filename pattern: %s", filename)
            print("  -> SKIP: Bad filename pattern")
            continue

        company_key = m.group(1)
        prod_row = producers_df[
            producers_df['normalized_name'] == normalize_name(company_key)
        ]
        if prod_row.empty:
            log.warning("No producer match for: %s", filename)
            print("  -> SKIP: Δεν βρέθηκε παραγωγός στο producers.xlsx για αυτό το filename")
            continue

        company_name = str(prod_row['Εταιρεία'].values[0])
        log.info("")
        log.info("=" * 60)
        log.info("PRODUCER: %s  (%s)", company_name, filename)
        log.info("=" * 60)
        print(f"  Εταιρεία: {company_name}")

        df = read_production_data(file_path)
        if df is None:
            print("  -> SKIP: read_production_data επέστρεψε None")
            continue

        df = filter_monthly_data(df, month)
        if df.empty:
            print(f"  -> SKIP: Δεν υπάρχουν γραμμές παραγωγής για μήνα {month}")
            continue

        df_daily, summary = calculate_daily_summary_quarterly(df, df_dam_15m, prod_row, month)
        if df_daily is None:
            print("  -> SKIP: calculate_daily_summary_quarterly επέστρεψε None")
            continue

        xlsx_path, company_name, email_value = generate_invoice_excel(
            df_daily, summary, prod_row, month, xlsx_dir
        )
        if not xlsx_path:
            print("  -> SKIP: generate_invoice_excel απέτυχε")
            continue

        email_key = (email_value or "NO_EMAIL").strip() or "NO_EMAIL"
        subfolder = determine_pdf_subfolder_name(
            email_key, email_to_companies, email_to_customs
        )
        target_dir = os.path.join(pdf_dir, subfolder[:MAX_FOLDER_CHARS])
        os.makedirs(target_dir, exist_ok=True)

        pdf_name = pdf_filename(company_name, month)
        pdf_path = os.path.join(target_dir, pdf_name)

        ok, how = export_to_pdf(xlsx_path, pdf_path)
        status = "✅ PDF" if ok else "❌ PDF"
        print(f"  {status} [{how}] → {pdf_path}")
        if not ok:
            log.error("PDF export failed: %s | method=%s", xlsx_path, how)
        log.info("")

    print(f"\nΈτοιμο. Δες: ΕΝΗΜΕΡΩΤΙΚΑ_ΣΗΜΕΙΩΜΑΤΑ/{month}/XLSX και /PDF")
