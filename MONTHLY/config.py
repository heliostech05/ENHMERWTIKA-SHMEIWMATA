#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
config.py — Κεντρική ρύθμιση για το μηνιαίο pipeline ενημερωτικών σημειωμάτων.

Περιέχει paths, σταθερές, logging, και configuration που χρησιμοποιούνται
από όλα τα υπόλοιπα modules.
"""

import logging
from pathlib import Path

# ===== Base paths =====
BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_PATH = BASE_DIR / "Invoice_GREEN_VALUE_01.xlsx"
PRODUCERS_PATH = BASE_DIR / "producers.xlsx"
PRODUCTION_DIR = BASE_DIR / "ΠΑΡΑΓΩΓΗ"
OUTPUT_DIR = BASE_DIR / "ΕΝΗΜΕΡΩΤΙΚΑ_ΣΗΜΕΙΩΜΑΤΑ"

# ===== DAM price files =====
DAM_FILE_PATTERN = "energy-charts_Electricity_production_and_spot_prices_in_Greece_in_{year}.csv"
DAM_QUARTER_CUTOFF = "2025-10-01"  # Αγνοούμε DAM γραμμές πριν αυτή την ημερομηνία (ωριαίες)
DAM_OUTPUT_DIR = BASE_DIR / "DAM DOWNLOAD" / "output"


def dam_file_for_year(year: int) -> Path:
    """
    Επιστρέφει το path του αρχείου DAM τιμών για ένα συγκεκριμένο έτος.

    Προτεραιότητα:
    1) DAM DOWNLOAD/output (νέα ροή)
    2) BASE_DIR (legacy θέση, για backward compatibility)
    """
    filename = DAM_FILE_PATTERN.format(year=int(year))
    preferred = DAM_OUTPUT_DIR / filename
    if preferred.exists():
        return preferred
    return BASE_DIR / filename


# ===== DST special-case dates =====
# Αυτές οι ημερομηνίες χρειάζονται ειδικό χειρισμό λόγω αλλαγής ώρας.
# TODO: Υπολογισμός δυναμικά βάσει έτους (EU DST κανόνες).
DST_SKIP_DATE = "2025-10-01"       # Πετάμε 00:15/00:30/00:45 (μετάβαση 15λέπτου αρχείου)
DST_FALLBACK_DATE = "2025-10-26"   # Διπλά 03:00–04:00 (fall-back), κρατάμε πρώτη εμφάνιση

# ===== Path / name limits =====
MAX_FOLDER_CHARS = 120
MAX_FILENAME_CHARS = 140

WIN_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}

# ===== Greek months (genitive) =====
GREEK_MONTHS_GENITIVE = {
    '01': 'Ιανουαρίου', '02': 'Φεβρουαρίου', '03': 'Μαρτίου',
    '04': 'Απριλίου',   '05': 'Μαΐου',       '06': 'Ιουνίου',
    '07': 'Ιουλίου',    '08': 'Αυγούστου',    '09': 'Σεπτεμβρίου',
    '10': 'Οκτωβρίου',  '11': 'Νοεμβρίου',    '12': 'Δεκεμβρίου',
}

# ===== Logging =====
LOG_DIR = BASE_DIR / "logs" / "timologia"
LOG_DIR.mkdir(parents=True, exist_ok=True)


def setup_logging() -> logging.Logger:
    """
    Ρύθμιση logging για το MONTHLY package.

    - File handler:    DEBUG+  → logs/timologia/monthly.log
    - Console handler: WARNING+ → stderr
    """
    logger = logging.getLogger("MONTHLY")
    if logger.handlers:
        return logger  # ήδη ρυθμισμένο

    logger.setLevel(logging.DEBUG)

    # File handler — λεπτομερές log σε αρχείο
    fh = logging.FileHandler(
        LOG_DIR / "monthly.log", encoding="utf-8",
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s | %(name)s.%(funcName)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    logger.addHandler(fh)

    # Console handler — μόνο warnings+
    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.addHandler(ch)

    return logger


# Αρχικοποίηση logger στο import
logger = setup_logging()
