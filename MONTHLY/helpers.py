#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
helpers.py — Βοηθητικές συναρτήσεις ονομάτων και paths.

Sanitization ονομάτων αρχείων/φακέλων, κλιπαρισμένα ονόματα
για μήκος path, κ.λπ.
"""

import re
from .config import WIN_RESERVED, MAX_FOLDER_CHARS, MAX_FILENAME_CHARS


def sanitize_name(name: str) -> str:
    """Αφαιρεί μη-επιτρεπόμενους χαρακτήρες αρχείων και ελέγχει reserved names."""
    if name is None:
        return "UNTITLED"
    s = str(name)
    s = re.sub(r'[\\/*?:"<>|]', "", s)
    s = s.strip().rstrip(".")
    if not s:
        s = "UNTITLED"
    if s.upper() in WIN_RESERVED:
        s = "_" + s
    return s


def normalize_name(name: str) -> str:
    """Ομαλοποίηση ονόματος εταιρείας για σύγκριση (lowercase, χωρίς κενά/σημεία)."""
    return re.sub(r'[\s._\\-]', '', str(name).strip().lower())


def join_with_limit(parts, sep=" & ", limit=120):
    """Ενώνει ονόματα με separator, κόβοντας αν ξεπεράσει limit χαρακτήρες."""
    out, total = [], 0
    for p in parts:
        piece = p if not out else (sep + p)
        if total + len(piece) > limit:
            break
        out.append(p)
        total += len(piece)
    if out:
        return sep.join(out)
    return (parts[0] if parts else "UNTITLED")[:limit]


def clipped_folder_name(preferred_names, fallback_names, limit=MAX_FOLDER_CHARS):
    """Δημιουργεί όνομα φακέλου από ονόματα εταιρειών, κλιπαρισμένο σε limit chars."""
    def build_name(items):
        uniq = sorted({sanitize_name(x) for x in items if x})
        return join_with_limit(uniq, sep=" & ", limit=limit)
    if preferred_names:
        return build_name(preferred_names)
    return build_name(fallback_names)


def clipped_filename(company_name: str, month: str, ext: str, limit=MAX_FILENAME_CHARS):
    """Δημιουργεί όνομα αρχείου ενημερωτικού σημειώματος, κλιπαρισμένο σε limit chars."""
    prefix = "ΕΝΗΜΕΡΩΤΙΚΟ_ΣΗΜΕΙΩΜΑ_"
    base = sanitize_name(company_name).replace(" ", "_")
    cand = f"{prefix}{base}_{month}.{ext}"
    if len(cand) <= limit:
        return cand
    over = len(cand) - limit
    base_cut = base[:max(1, len(base) - over)]
    cand = f"{prefix}{base_cut}_{month}.{ext}"
    if len(cand) > limit:
        trunk = f"{prefix}{month}"
        cand = trunk[:limit - (len(ext) + 1)] + "." + ext
    return cand
