"""
Entrypoint for: python -m MONTHLY

Τρέχει το μηνιαίο pipeline ενημερωτικών σημειωμάτων.
"""

import re
from .timologia import timologia

month_input = input("Δώσε μήνα (YYYY-MM): ").strip()
if not re.match(r'^\d{4}-(0[1-9]|1[0-2])$', month_input):
    print("Μη έγκυρη μορφή. Παράδειγμα: 2025-10")
else:
    timologia(month_input)
