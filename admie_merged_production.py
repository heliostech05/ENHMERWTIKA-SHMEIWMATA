# -*- coding: utf-8 -*-
"""
admie_merged_production.py — Διαχωρισμός GREEN_VE6 CSVs ανά παραγωγό.

Διαβάζει ημερήσια GREEN_VE6 αρχεία από downloads/{YYYY-MM}/,
τα σπάει ανά ΚΩΔΙΚΟ ΕΔΡΕΘ, και τα συγχωνεύει σε per-company
αρχεία CSV στο ΠΑΡΑΓΩΓΗ/.

Χρήση:
    python admie_merged_production.py
"""

import logging
import os
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

from MONTHLY.config import BASE_DIR, PRODUCERS_PATH, PRODUCTION_DIR
from MONTHLY.helpers import sanitize_name

# ===== Logging =====
LOG_FILE = BASE_DIR / "logs" / "merged_production.log"
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

log = logging.getLogger("merged_production")
log.setLevel(logging.DEBUG)

if not log.handlers:
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s | %(funcName)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    log.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    log.addHandler(ch)

# ===== Constants =====
DOWNLOADS_DIR = BASE_DIR / "downloads"

# Regex: matches all documented GREEN_VE6 filename variants:
#   GREEN_VE6YYYYMMDD.csv       GREEN_VE6_YYYYMMDD.csv
#   GREEN_VE6YYYYMMDD1.csv      GREEN_VE6_YYYYMMDD_2.csv
_VE6_PATTERN = re.compile(
    r"GREEN_VE6_?(\d{8})_?(\d)?\.csv", re.IGNORECASE
)


# ===== Producer loading =====

def load_producers(filepath):
    """Φόρτωση producers.xlsx — επιστρέφει DataFrame με stripped Code/Εταιρεία."""
    try:
        df = pd.read_excel(filepath, dtype={"Code": str})
        if "Code" not in df.columns or "Εταιρεία" not in df.columns:
            raise ValueError("Λείπουν στήλες 'Code' και/ή 'Εταιρεία'")

        df["Code"] = df["Code"].astype(str).str.strip()
        df["Εταιρεία"] = df["Εταιρεία"].astype(str).str.strip()

        log.info("Loaded %d producers", len(df))
        for _, row in df.iterrows():
            log.debug("  Code %s → %s", row["Code"], row["Εταιρεία"])
        return df

    except Exception as e:
        log.error("Failed to load producers: %s", e)
        return None


# ===== File selection =====

def get_latest_green_ve6_files(folder):
    """
    Επιστρέφει τη νεότερη έκδοση ανά ημερομηνία για αρχεία GREEN_VE6.

    Υποστηριζόμενα patterns:
        GREEN_VE6YYYYMMDD.csv       GREEN_VE6_YYYYMMDD.csv
        GREEN_VE6YYYYMMDD1.csv      GREEN_VE6_YYYYMMDD_2.csv
    """
    date_to_file = defaultdict(list)
    csv_files = [f for f in os.listdir(folder)
                 if f.upper().startswith("GREEN_VE6") and f.endswith(".csv")]

    log.info("Found %d GREEN_VE6 files in %s", len(csv_files), folder)

    for filename in csv_files:
        m = _VE6_PATTERN.match(filename)
        if not m:
            log.warning("Skipping unrecognized file: %s", filename)
            continue
        date = m.group(1)
        edition = int(m.group(2)) if m.group(2) else 0
        date_to_file[date].append((edition, filename))

    latest_files = []
    for date in sorted(date_to_file.keys()):
        best_edition, best_file = max(date_to_file[date])
        log.debug("  %s → edition %d: %s", date, best_edition, best_file)
        latest_files.append(best_file)

    log.info("Selected %d files (one per date)", len(latest_files))
    return latest_files


# ===== Timestamp processing =====

def preprocess_timestamp_column(df):
    """Μετατροπή TIMESTAMP — χειρισμός 24:00 ως 00:00 επόμενης ημέρας."""
    if "TIMESTAMP" not in df.columns:
        raise ValueError("Λείπει η στήλη TIMESTAMP")

    is_24 = df["TIMESTAMP"].str.contains("24:00", regex=False)
    new_ts = df["TIMESTAMP"].copy()
    new_ts[is_24] = (
        pd.to_datetime(
            df.loc[is_24, "TIMESTAMP"].str.replace("24:00", "00:00"),
            format="%d/%m/%Y %H:%M", errors="coerce",
        ) + pd.Timedelta(days=1)
    ).dt.strftime("%d/%m/%Y %H:%M")

    df["TIMESTAMP"] = new_ts
    df["datetime"] = pd.to_datetime(df["TIMESTAMP"], format="%d/%m/%Y %H:%M", errors="coerce")
    log.debug("Processed %d timestamps", len(df))
    return df


def assign_month_column(df):
    """
    Προσθήκη στήλης Μήνας (YYYY-MM).

    Ειδική περίπτωση: 00:00 της 1ης ημέρας ανήκει στον ΠΡΟΗΓΟΥΜΕΝΟ μήνα
    (είναι το τελευταίο 15λεπτο εκείνου του μήνα).
    """
    at_month_start = (
        (df["datetime"].dt.day == 1)
        & (df["datetime"].dt.hour == 0)
        & (df["datetime"].dt.minute == 0)
    )
    df["Μήνας"] = df["datetime"].dt.to_period("M").astype(str)
    df.loc[at_month_start, "Μήνας"] = (
        df.loc[at_month_start, "datetime"] - pd.DateOffset(days=1)
    ).dt.to_period("M").astype(str)

    log.debug("Assigned month column to %d rows", len(df))
    return df


# ===== Merge with existing CSV =====

def merge_with_existing_csv(group_df, out_file):
    """Συγχώνευση νέων δεδομένων με υπάρχον CSV — νέα δεδομένα κερδίζουν σε σύγκρουση."""
    group_df = group_df.copy()
    group_df.set_index("TIMESTAMP", inplace=True)

    if os.path.exists(out_file):
        try:
            existing = pd.read_csv(out_file, delimiter=";", encoding="utf-8-sig")
            existing.set_index("TIMESTAMP", inplace=True)
            # Κρατάμε μόνο τα παλιά rows που ΔΕΝ υπάρχουν στα νέα
            combined = pd.concat([existing[~existing.index.isin(group_df.index)], group_df])
            log.debug("Merged with existing: %s", out_file)
        except Exception as e:
            log.warning("Could not read existing %s: %s — overwriting", out_file, e)
            combined = group_df
    else:
        log.debug("Creating new file: %s", out_file)
        combined = group_df

    # Sort chronologically
    combined = combined.reset_index()
    combined["_sort"] = pd.to_datetime(combined["TIMESTAMP"], format="%d/%m/%Y %H:%M", errors="coerce")
    combined = combined.sort_values("_sort").drop(columns=["_sort"])
    return combined


# ===== Per-file processing =====

def process_file(filepath, producers_df, output_folder):
    """
    Διαβάζει ένα GREEN_VE6 CSV, σπάει ανά ΚΩΔΙΚΟ ΕΔΡΕΘ,
    και γράφει/συγχωνεύει στο αντίστοιχο ΠΑΡΑΓΩΓΗ_{company}.csv.
    """
    log.info("Processing: %s", filepath)
    try:
        df = pd.read_csv(filepath, delimiter=";", encoding="utf-8-sig", skiprows=1)
    except Exception as e:
        log.error("Failed to read %s: %s", filepath, e)
        return

    if "ΚΩΔΙΚΟΣ ΕΔΡΕΘ" not in df.columns:
        log.error("Missing column 'ΚΩΔΙΚΟΣ ΕΔΡΕΘ' in %s", filepath)
        return

    try:
        df = preprocess_timestamp_column(df)
    except ValueError as e:
        log.error("Timestamp error in %s: %s", filepath, e)
        return

    df = assign_month_column(df)

    for code_value, group in df.groupby("ΚΩΔΙΚΟΣ ΕΔΡΕΘ"):
        code_str = str(code_value).strip()
        producer_row = producers_df[producers_df["Code"] == code_str]

        if producer_row.empty:
            log.warning("Unknown ΚΩΔΙΚΟΣ ΕΔΡΕΘ: %s", code_str)
            continue

        company_name = producer_row["Εταιρεία"].values[0]
        safe_name = sanitize_name(company_name).replace(" ", "_")
        out_file = os.path.join(output_folder, f"ΠΑΡΑΓΩΓΗ_{safe_name}.csv")

        final_df = merge_with_existing_csv(group, out_file)
        final_df.to_csv(out_file, index=False, sep=";", encoding="utf-8-sig")
        log.info("  %s (%s) → %s (%d rows)", company_name, code_str, out_file, len(final_df))


# ===== Orchestrator =====

def split_files_by_code(month):
    """
    Κύρια ροή:
    1. Φόρτωση producers.xlsx
    2. Εύρεση τελευταίων GREEN_VE6 CSVs ανά ημέρα
    3. Σπάσιμο ανά κωδικό ΕΔΡΕΘ → per-company output CSVs
    """
    input_folder = str(DOWNLOADS_DIR / month)
    output_folder = str(PRODUCTION_DIR)

    if not os.path.isdir(input_folder):
        log.error("Input folder not found: %s", input_folder)
        print(f"❌ Ο φάκελος {input_folder} δεν υπάρχει.")
        return

    os.makedirs(output_folder, exist_ok=True)

    producers_df = load_producers(str(PRODUCERS_PATH))
    if producers_df is None:
        print("❌ Αποτυχία φόρτωσης producers.xlsx")
        return

    latest_files = get_latest_green_ve6_files(input_folder)
    if not latest_files:
        log.warning("No GREEN_VE6 files found in %s", input_folder)
        print(f"⚠️ Δεν βρέθηκαν GREEN_VE6 αρχεία στο {input_folder}")
        return

    log.info("")
    log.info("=" * 60)
    log.info("SPLIT FILES BY CODE — month=%s", month)
    log.info("=" * 60)

    for filename in latest_files:
        file_path = os.path.join(input_folder, filename)
        process_file(file_path, producers_df, output_folder)

    log.info("Completed: %d files processed", len(latest_files))
    print(f"\n✅ Έτοιμο. {len(latest_files)} αρχεία επεξεργάστηκαν → {output_folder}/")


# ===== Entrypoint =====

if __name__ == "__main__":
    month_input = input("Δώσε μήνα (YYYY-MM): ").strip()
    if not re.match(r"^\d{4}-(0[1-9]|1[0-2])$", month_input):
        print("Μη έγκυρη μορφή. Παράδειγμα: 2025-01")
        sys.exit(1)
    split_files_by_code(month_input)
