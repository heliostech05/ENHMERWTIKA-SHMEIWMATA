#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Ενιαίο Streamlit app για όλα τα στάδια έκδοσης ενημερωτικών σημειωμάτων.

Καρτέλες:
1. Κατέβασμα Modesto
2. Κατηγοριοποίηση παραγωγής ανά πάρκο
3. Μηνιαία ενημερωτικά σημειώματα
4. Εβδομαδιαία ενημερωτικά σημειώματα
"""

import io
import json
import os
import re
import subprocess
import sys
import zipfile
from calendar import monthrange
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

from MONTHLY.config import OUTPUT_DIR, PRODUCERS_PATH


MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
ISO_WEEK_RE = re.compile(r"^(?P<year>\d{4})-W(?P<week>0[1-9]|[1-4]\d|5[0-3])$")

BASE_DIR = Path(__file__).resolve().parent
HISTORY_FILE = BASE_DIR / "logs" / "streamlit_unified_history.json"
WEEKLY_OUTPUT_DIR = BASE_DIR / "ΕΝΗΜΕΡΩΤΙΚΑ_ΣΗΜΕΙΩΜΑΤΑ_ΕΒΔΟΜΑΔΙΑΙΑ"


def _default_month() -> str:
    return date.today().strftime("%Y-%m")


def _default_week_range() -> tuple[str, str]:
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    return monday.isoformat(), sunday.isoformat()


def _parse_date(value: str) -> date | None:
    if not DATE_RE.match(value):
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return None


def _week_tag(start_value: date) -> str:
    iso_year, iso_week, _ = start_value.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


def _parse_iso_week(value: str) -> tuple[date, date] | None:
    match = ISO_WEEK_RE.match(value.strip())
    if not match:
        return None
    iso_year = int(match.group("year"))
    iso_week = int(match.group("week"))
    try:
        monday = date.fromisocalendar(iso_year, iso_week, 1)
        sunday = date.fromisocalendar(iso_year, iso_week, 7)
    except ValueError:
        return None
    return monday, sunday


def _run_command(cmd: list[str], env_extra: dict | None = None) -> tuple[bool, str]:
    env = os.environ.copy()
    if env_extra:
        env.update(env_extra)

    proc = subprocess.run(
        cmd,
        cwd=str(BASE_DIR),
        capture_output=True,
        text=True,
        env=env,
    )
    output = (proc.stdout or "") + ("\n" if proc.stdout and proc.stderr else "") + (proc.stderr or "")
    return proc.returncode == 0, output.strip()


def _run_dam_download():
    cmd = [sys.executable, str(BASE_DIR / "DAM DOWNLOAD" / "dam_api_download.py")]
    return _run_command(cmd)


def _run_modesto(year: int, month_num: int, start_day: int, end_day: int):
    cmd = [
        sys.executable,
        "admie_modesto_files_extraction.py",
        "--year", str(year),
        "--month", str(month_num),
        "--start-day", str(start_day),
        "--end-day", str(end_day),
        "--cert", "./certificates/client_modesto_cert.pem",
        "--key", "./certificates/client_modesto_key.pem",
        "--out", "downloads",
        "--verify-ssl",
    ]
    return _run_command(cmd)


def _run_production_month(month_value: str):
    code = "import os, admie_merged_production as m; m.split_files_by_code(os.environ['PIPE_MONTH'])"
    cmd = [sys.executable, "-c", code]
    return _run_command(cmd, env_extra={"PIPE_MONTH": month_value})


def _run_production_range(start_value: str, end_value: str):
    code = (
        "import os; "
        "from datetime import datetime; "
        "from weekly.timologia import ensure_production_files; "
        "ensure_production_files("
        "datetime.fromisoformat(os.environ['PIPE_START']), "
        "datetime.fromisoformat(os.environ['PIPE_END'])"
        ")"
    )
    cmd = [sys.executable, "-c", code]
    return _run_command(cmd, env_extra={"PIPE_START": start_value, "PIPE_END": end_value})


def _run_monthly_timologia(month_value: str):
    code = "import os; from MONTHLY.timologia import timologia; timologia(os.environ['PIPE_MONTH'])"
    cmd = [sys.executable, "-c", code]
    return _run_command(cmd, env_extra={"PIPE_MONTH": month_value})


def _run_weekly_timologia(start_value: str, end_value: str):
    code = (
        "import os; "
        "from weekly.timologia import timologia_weekly; "
        "timologia_weekly(os.environ['PIPE_START'], os.environ['PIPE_END'])"
    )
    cmd = [sys.executable, "-c", code]
    return _run_command(cmd, env_extra={"PIPE_START": start_value, "PIPE_END": end_value})


def _show_output(title: str, ok: bool, output: str):
    if ok:
        st.success(f"{title}: ΟΚ")
    else:
        st.error(f"{title}: ΑΠΕΤΥΧΕ")
    if output:
        st.code(output)
    else:
        st.info("Δεν υπήρχε output από την εκτέλεση.")


def _load_history() -> list[dict]:
    if not HISTORY_FILE.exists():
        return []
    try:
        data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_history(entries: list[dict]):
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_FILE.write_text(
        json.dumps(entries[-500:], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _snapshot_files(paths: list[Path]) -> dict[str, int]:
    snapshot: dict[str, int] = {}
    for path in paths:
        if not path.exists():
            continue
        if path.is_file():
            snapshot[str(path)] = path.stat().st_mtime_ns
            continue
        for fp in path.rglob("*"):
            if fp.is_file():
                snapshot[str(fp)] = fp.stat().st_mtime_ns
    return snapshot


def _diff_snapshots(before: dict[str, int], after: dict[str, int]) -> list[str]:
    changed = []
    for path_str, mtime in after.items():
        if before.get(path_str) != mtime:
            try:
                changed.append(str(Path(path_str).resolve().relative_to(BASE_DIR)))
            except Exception:
                changed.append(path_str)
    changed.sort()
    return changed


def _record_step_history(step_key: str, period_key: str, ok: bool, output: str, artifacts: list[Path], changed_files: list[str]):
    entries = _load_history()
    entries.append({
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "step": step_key,
        "period": period_key,
        "ok": ok,
        "artifacts": [str(p) for p in artifacts],
        "changed_files": changed_files,
        "output_tail": (output or "")[-3000:],
    })
    _save_history(entries)


def _latest_step_history(step_key: str, period_key: str) -> dict | None:
    entries = _load_history()
    for item in reversed(entries):
        if item.get("step") == step_key and item.get("period") == period_key:
            return item
    return None


def _execute_step(step_key: str, period_key: str, artifacts: list[Path], runner):
    before = _snapshot_files(artifacts)
    ok, output = runner()
    after = _snapshot_files(artifacts)
    changed_files = _diff_snapshots(before, after)
    _record_step_history(step_key, period_key, ok, output, artifacts, changed_files)
    return ok, output, changed_files


def _show_changed_files(changed_files: list[str], allowed_suffixes: set[str] | None = None):
    if allowed_suffixes:
        changed_files = [p for p in changed_files if Path(p).suffix.lower() in allowed_suffixes]
    if changed_files:
        st.caption(f"Νέα/ενημερωμένα αρχεία: {len(changed_files)}")
        st.code("\n".join(changed_files[:300]))
        if len(changed_files) > 300:
            st.caption(f"... και άλλα {len(changed_files) - 300}")
    else:
        st.info("Δεν εντοπίστηκαν νέα ή τροποποιημένα αρχεία.")


def _human_size(num_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB"]
    value = float(num_bytes)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{num_bytes} B"


def _list_files_for_view(
    dir_path: Path,
    max_rows: int = 400,
    allowed_suffixes: set[str] | None = None,
) -> list[dict]:
    files = [p for p in dir_path.rglob("*") if p.is_file()]
    if allowed_suffixes:
        files = [p for p in files if p.suffix.lower() in allowed_suffixes]
    files.sort(key=lambda p: str(p.relative_to(dir_path)).casefold())
    rows = []
    for fp in files[:max_rows]:
        rows.append({
            "Αρχείο": str(fp.relative_to(dir_path)),
            "Μέγεθος": _human_size(fp.stat().st_size),
        })
    return rows


def _zip_dir_bytes(
    dir_path: Path,
    include_root_folder: bool = True,
    allowed_suffixes: set[str] | None = None,
) -> tuple[bytes | None, int]:
    if not dir_path.exists() or not dir_path.is_dir():
        return None, 0

    files = [p for p in dir_path.rglob("*") if p.is_file()]
    if allowed_suffixes:
        files = [p for p in files if p.suffix.lower() in allowed_suffixes]
    if not files:
        return None, 0

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fpath in files:
            rel = fpath.relative_to(dir_path)
            arcname = str(Path(dir_path.name) / rel) if include_root_folder else str(rel)
            zf.write(fpath, arcname)
    return buf.getvalue(), len(files)


def _render_artifacts_panel(
    step_key: str,
    title: str,
    period_key: str,
    dir_path: Path,
    zip_name: str,
    allowed_suffixes: set[str] | None = None,
    panel_key_suffix: str = "",
):
    st.write(f"**{title}**")
    st.write(f"Φάκελος: `{dir_path}`")

    latest = _latest_step_history(step_key, period_key)
    if latest:
        status = "OK" if latest.get("ok") else "FAILED"
        st.caption(f"Τελευταία εκτέλεση: {latest.get('timestamp')} | {status}")
        changed = latest.get("changed_files") or []
        if allowed_suffixes:
            changed = [p for p in changed if Path(p).suffix.lower() in allowed_suffixes]
        if changed:
            st.caption(f"Τελευταία αλλαγμένα αρχεία: {len(changed)}")
            st.code("\n".join(changed[:120]))

    if not dir_path.exists():
        st.info("Δεν υπάρχει ακόμα φάκελος για αυτό το βήμα.")
        return

    zip_data, count = _zip_dir_bytes(
        dir_path,
        include_root_folder=True,
        allowed_suffixes=allowed_suffixes,
    )
    if zip_data:
        st.download_button(
            label=f"Κατέβασμα zip ({count} αρχεία)",
            data=zip_data,
            file_name=zip_name,
            mime="application/zip",
            width="stretch",
            key=f"dl_{step_key}_{period_key}_{panel_key_suffix or 'all'}",
        )
    else:
        st.caption("Δεν υπάρχουν αρχεία για zip.")

    rows = _list_files_for_view(dir_path, allowed_suffixes=allowed_suffixes)
    if rows:
        st.dataframe(rows, width="stretch", hide_index=True)
    else:
        st.caption("Δεν βρέθηκαν αρχεία.")


def _render_dam_block(button_key: str, panel_suffix: str):
    left, right = st.columns([1, 1], gap="large")
    with left:
        st.caption(
            "Κατέβασε ή ανανέωσε τα CSV των τιμών DAM. "
        )
        if st.button("Κατέβασε τιμές DAM", key=button_key, width="stretch"):
            dam_artifacts = [BASE_DIR / "DAM DOWNLOAD" / "output"]
            with st.spinner("Τρέχει το DAM download..."):
                ok, output, changed = _execute_step(
                    step_key="dam_download",
                    period_key="GLOBAL",
                    artifacts=dam_artifacts,
                    runner=_run_dam_download,
                )
            _show_output("DAM", ok, output)
            _show_changed_files(changed, allowed_suffixes={".csv"})

    with right:
        _render_artifacts_panel(
            step_key="dam_download",
            title="Διαθέσιμα αρχεία DAM",
            period_key="GLOBAL",
            dir_path=BASE_DIR / "DAM DOWNLOAD" / "output",
            zip_name="dam_prices_output.zip",
            allowed_suffixes={".csv"},
            panel_key_suffix=panel_suffix,
        )


st.set_page_config(
    page_title="ΕΝΗΜΕΡΩΤΙΚΑ ΣΗΜΕΙΩΜΑΤΑ",
    page_icon="E",
    layout="wide",
)

default_month = _default_month()
default_week_start, default_week_end = _default_week_range()

st.title("ΕΝΙΑΙΟ APP ΕΝΗΜΕΡΩΤΙΚΩΝ ΣΗΜΕΙΩΜΑΤΩΝ")

tab_dam, tab_modesto, tab_producers, tab_production, tab_monthly, tab_weekly = st.tabs([
    "1. DAM Prices",
    "2. Modesto",
    "3. Producers",
    "4. Παραγωγή",
    "5. Μηνιαία",
    "6. Εβδομαδιαία",
])

with tab_dam:
    st.subheader("DAM Prices")
    st.caption("Στο συγκεκριμένο tab μπορείς να κατεβάσεις ή να ανανεώσεις τα CSV αρχεία με τις τιμές DAM του τρεχοντος ετους μέχρι την σημερινή ημερομηνία.")
    _render_dam_block("dam_button_tab", "dam_tab")

with tab_modesto:
    st.subheader("Κατέβασμα αρχείων από Modesto")
    month_col, _ = st.columns([1, 3], gap="large")
    with month_col:
        month_modesto = st.text_input("Μήνας (YYYY-MM)", value=default_month, key="modesto_month").strip()
    month_ok = bool(MONTH_RE.match(month_modesto))

    if month_ok:
        modesto_year = int(month_modesto.split("-")[0])
        modesto_month_num = int(month_modesto.split("-")[1])
        modesto_last_day = monthrange(modesto_year, modesto_month_num)[1]
    else:
        today = date.today()
        modesto_year, modesto_month_num = today.year, today.month
        modesto_last_day = monthrange(today.year, today.month)[1]
        st.warning("Μη έγκυρη μορφή μήνα. Χρησιμοποίησε YYYY-MM.")

    left, right = st.columns([1, 1], gap="large")
    with left:
        st.caption(
            "Κατέβασε τα GREEN_VE6 αρχεία για το επιλεγμένο διάστημα ημερών. "
            "Το Modesto endpoint δέχεται έως 10 ημέρες ανά εκτέλεση."
        )
        col1, col2 = st.columns(2)
        with col1:
            start_day = st.number_input(
                "Από ημέρα",
                min_value=1,
                max_value=modesto_last_day,
                value=1,
                step=1,
                key="modesto_start_day",
            )
        with col2:
            end_day = st.number_input(
                "Έως ημέρα",
                min_value=1,
                max_value=modesto_last_day,
                value=min(modesto_last_day, 10),
                step=1,
                key="modesto_end_day",
            )

        if end_day < start_day:
            st.error("Το 'Έως ημέρα' πρέπει να είναι >= από το 'Από ημέρα'.")
        elif (end_day - start_day) > 9:
            st.error("Το εύρος ημερών στο Modesto πρέπει να είναι έως 10 ημέρες.")

        if st.button("Κατέβασε αρχεία Modesto", key="run_modesto", width="stretch"):
            if not month_ok:
                st.error("Διόρθωσε πρώτα τον μήνα.")
            elif end_day < start_day or (end_day - start_day) > 9:
                st.error("Διόρθωσε πρώτα το εύρος ημερών.")
            else:
                artifacts = [BASE_DIR / "downloads" / month_modesto]
                with st.spinner("Τρέχει το Modesto download..."):
                    ok, output, changed = _execute_step(
                        step_key="modesto_download",
                        period_key=month_modesto,
                        artifacts=artifacts,
                        runner=lambda: _run_modesto(
                            modesto_year, modesto_month_num, int(start_day), int(end_day)
                        ),
                    )
                _show_output("Modesto", ok, output)
                _show_changed_files(changed, allowed_suffixes={".csv"})

    with right:
        if month_ok:
            _render_artifacts_panel(
                step_key="modesto_download",
                title="Διαθέσιμα αρχεία Modesto",
                period_key=month_modesto,
                dir_path=BASE_DIR / "downloads" / month_modesto,
                zip_name=f"{month_modesto}_modesto.zip",
                allowed_suffixes={".csv"},
                panel_key_suffix="modesto",
            )
        else:
            st.info("Δώσε έγκυρο μήνα για να εμφανιστούν αρχεία.")

with tab_producers:
    st.subheader("Preview Producers")
    st.caption("Προβολή του αρχείου `producers.xlsx` που χρησιμοποιείται από το pipeline.")

    if not PRODUCERS_PATH.exists():
        st.error(f"Δεν βρέθηκε το αρχείο: `{PRODUCERS_PATH}`")
    else:
        left, right = st.columns([1, 1], gap="large")
        with left:
            st.write(f"Αρχείο: `{PRODUCERS_PATH}`")
            st.caption(f"Μέγεθος: {_human_size(PRODUCERS_PATH.stat().st_size)}")
            try:
                producers_df = pd.read_excel(PRODUCERS_PATH)
                st.caption(f"Γραμμές: {len(producers_df)} | Στήλες: {len(producers_df.columns)}")
                st.dataframe(producers_df, width="stretch", hide_index=True)
            except Exception as exc:
                st.error(f"Αποτυχία ανάγνωσης του producers.xlsx: {exc}")

        with right:
            try:
                producers_bytes = PRODUCERS_PATH.read_bytes()
                st.download_button(
                    label="Κατέβασμα producers.xlsx",
                    data=producers_bytes,
                    file_name=PRODUCERS_PATH.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key="dl_producers_xlsx",
                )
            except Exception as exc:
                st.error(f"Δεν ήταν δυνατό το download του αρχείου: {exc}")

with tab_production:
    st.subheader("Κατηγοριοποίηση παραγωγής ανά πάρκο")
    st.caption("Ο ίδιος φάκελος `ΠΑΡΑΓΩΓΗ` χρησιμοποιείται και από τη μηνιαία και από την εβδομαδιαία ροή.")

    left, right = st.columns([1, 1], gap="large")
    with left:
        production_month = st.text_input(
            "Μήνας παραγωγής (YYYY-MM)",
            value=default_month,
            key="production_month",
        ).strip()
        production_month_ok = bool(MONTH_RE.match(production_month))
        if not production_month_ok:
            st.warning("Δώσε έγκυρο μήνα σε μορφή YYYY-MM.")

        if st.button("Ενημέρωσε το ΠΑΡΑΓΩΓΗ", key="run_production_month", width="stretch"):
            if not production_month_ok:
                st.error("Διόρθωσε πρώτα τον μήνα.")
            else:
                artifacts = [BASE_DIR / "ΠΑΡΑΓΩΓΗ"]
                with st.spinner("Τρέχει η κατηγοριοποίηση παραγωγής..."):
                    ok, output, changed = _execute_step(
                        step_key="production_month",
                        period_key=production_month,
                        artifacts=artifacts,
                        runner=lambda: _run_production_month(production_month),
                    )
                _show_output("Παραγωγή ανά πάρκο", ok, output)
                _show_changed_files(changed, allowed_suffixes={".csv"})

    with right:
        _render_artifacts_panel(
            step_key="production_view",
            title="Διαθέσιμα αρχεία ΠΑΡΑΓΩΓΗ",
            period_key="ALL",
            dir_path=BASE_DIR / "ΠΑΡΑΓΩΓΗ",
            zip_name="paragogi.zip",
            allowed_suffixes={".csv"},
            panel_key_suffix="production",
        )

with tab_monthly:
    st.subheader("Μηνιαία ενημερωτικά σημειώματα")
    monthly_month = st.text_input("Μήνας τιμολόγησης (YYYY-MM)", value=default_month, key="monthly_month").strip()
    monthly_month_ok = bool(MONTH_RE.match(monthly_month))

    if not monthly_month_ok:
        st.warning("Μη έγκυρη μορφή μήνα. Χρησιμοποίησε YYYY-MM.")

    st.caption(
        "Δημιουργεί τα τελικά μηνιαία ενημερωτικά σημειώματα σε Excel και PDF "
        "για τον επιλεγμένο μήνα."
    )
    if st.button("Βγάλε μηνιαία ενημερωτικά", key="run_monthly_timologia", type="primary", width="stretch"):
        if not monthly_month_ok:
            st.error("Διόρθωσε πρώτα τον μήνα.")
        else:
            artifacts = [Path(OUTPUT_DIR) / monthly_month]
            with st.spinner("Τρέχει η μηνιαία έκδοση ενημερωτικών..."):
                ok, output, changed = _execute_step(
                    step_key="monthly_timologia",
                    period_key=monthly_month,
                    artifacts=artifacts,
                    runner=lambda: _run_monthly_timologia(monthly_month),
                )
            _show_output("Μηνιαία ενημερωτικά", ok, output)
            _show_changed_files(changed)
            if ok:
                st.write(f"Φάκελος εξόδου: `{Path(OUTPUT_DIR) / monthly_month}`")

    st.markdown("---")
    if monthly_month_ok:
        monthly_out = Path(OUTPUT_DIR) / monthly_month
        out_xlsx, out_pdf = st.tabs(["Excel (XLSX)", "PDF"])
        with out_xlsx:
            _render_artifacts_panel(
                step_key="monthly_timologia",
                title="Διαθέσιμα αρχεία Excel",
                period_key=monthly_month,
                dir_path=monthly_out / "XLSX",
                zip_name=f"{monthly_month}_monthly_excel.zip",
                allowed_suffixes={".xlsx"},
                panel_key_suffix="monthly_xlsx",
            )
        with out_pdf:
            _render_artifacts_panel(
                step_key="monthly_timologia",
                title="Διαθέσιμα αρχεία PDF",
                period_key=monthly_month,
                dir_path=monthly_out / "PDF",
                zip_name=f"{monthly_month}_monthly_pdf.zip",
                allowed_suffixes={".pdf"},
                panel_key_suffix="monthly_pdf",
            )

        st.markdown("---")
        all_zip, all_count = _zip_dir_bytes(
            monthly_out,
            include_root_folder=True,
            allowed_suffixes={".xlsx", ".pdf"},
        )
        if all_zip:
            st.download_button(
                label=f"Κατέβασμα όλων (Excel + PDF) σε zip ({all_count} αρχεία)",
                data=all_zip,
                file_name=f"{monthly_month}_monthly_excel_pdf.zip",
                mime="application/zip",
                width="stretch",
                key=f"dl_monthly_all_{monthly_month}",
            )
        else:
            st.caption("Δεν υπάρχουν διαθέσιμα Excel/PDF αρχεία για συνολικό zip.")
    else:
        st.info("Δώσε έγκυρο μήνα για να εμφανιστούν τα παραδοτέα.")

with tab_weekly:
    st.subheader("Εβδομαδιαία ενημερωτικά σημειώματα")
    weekly_input_mode = st.radio(
        "Τρόπος επιλογής εβδομάδας",
        ["ISO week", "Χειροκίνητη επιλογή"],
        horizontal=True,
        key="weekly_input_mode",
    )

    if weekly_input_mode == "ISO week":
        default_week_tag = _week_tag(_parse_date(default_week_start) or date.today())
        iso_week_value = st.text_input(
            "ISO week (YYYY-Www)",
            value=default_week_tag,
            key="weekly_iso_week",
        ).strip().upper()
        iso_week_range = _parse_iso_week(iso_week_value)
        if iso_week_range is None:
            week_start_dt = None
            week_end_dt = None
            week_start_value = ""
            week_end_value = ""
        else:
            week_start_dt, week_end_dt = iso_week_range
            week_start_value = week_start_dt.isoformat()
            week_end_value = week_end_dt.isoformat()
            st.info(
                "Η επιλεγμένη ISO εβδομάδα αφορά το διάστημα "
                f"{week_start_value} έως {week_end_value}."
            )
    else:
        week_start_dt = st.date_input(
            "Αρχή εβδομάδας",
            value=_parse_date(default_week_start) or date.today(),
            format="YYYY-MM-DD",
            key="weekly_start_date",
        )
        week_end_dt = st.date_input(
            "Τέλος εβδομάδας",
            value=_parse_date(default_week_end) or date.today(),
            format="YYYY-MM-DD",
            key="weekly_end_date",
        )
        week_start_value = week_start_dt.isoformat()
        week_end_value = week_end_dt.isoformat()
        st.caption(
            "Χειροκίνητο διάστημα εβδομάδας: "
            f"`{week_start_value}` έως `{week_end_value}`."
        )

    weekly_dates_ok = week_start_dt is not None and week_end_dt is not None
    weekly_range_ok = weekly_dates_ok and week_end_dt >= week_start_dt
    weekly_full_week_ok = (
        weekly_dates_ok and
        week_start_dt.weekday() == 0 and
        week_end_dt.weekday() == 6 and
        (week_end_dt - week_start_dt).days == 6
    )

    if not weekly_dates_ok:
        if weekly_input_mode == "ISO week":
            st.warning("Δώσε έγκυρο ISO week σε μορφή YYYY-Www, π.χ. 2026-W14.")
        else:
            st.warning("Δώσε έγκυρες ημερομηνίες εβδομάδας.")
    elif not weekly_range_ok:
        st.error("Το τέλος εβδομάδας πρέπει να είναι μετά ή ίσο με την αρχή.")
    elif not weekly_full_week_ok:
        st.error("Το weekly pipeline απαιτεί ακριβώς 7 ημέρες, από Δευτέρα έως Κυριακή.")
    else:
        st.caption(f"ISO εβδομάδα: `{_week_tag(week_start_dt)}`")
        st.caption(f"Φάκελος εξόδου: `{WEEKLY_OUTPUT_DIR / _week_tag(week_start_dt)}`")

    st.caption(
        "Δημιουργεί τα εβδομαδιαία ενημερωτικά ΣΗΘΥΑ σε Excel και PDF "
        "για το επιλεγμένο διάστημα Δευτέρα-Κυριακή."
    )
    if st.button("Βγάλε εβδομαδιαία ενημερωτικά", key="run_weekly_timologia", type="primary", width="stretch"):
        if not weekly_dates_ok:
            st.error("Δώσε έγκυρες ημερομηνίες πρώτα.")
        elif not weekly_range_ok:
            st.error("Διόρθωσε πρώτα το διάστημα.")
        elif not weekly_full_week_ok:
            st.error("Το weekly pipeline απαιτεί ακριβώς 7 ημέρες, από Δευτέρα έως Κυριακή.")
        else:
            week_tag = _week_tag(week_start_dt)
            period_key = f"{week_start_value}__{week_end_value}"
            artifacts = [WEEKLY_OUTPUT_DIR / week_tag]
            with st.spinner("Τρέχει η εβδομαδιαία έκδοση ενημερωτικών..."):
                ok, output, changed = _execute_step(
                    step_key="weekly_timologia",
                    period_key=period_key,
                    artifacts=artifacts,
                    runner=lambda: _run_weekly_timologia(week_start_value, week_end_value),
                )
            _show_output("Εβδομαδιαία ενημερωτικά", ok, output)
            _show_changed_files(changed)
            if ok:
                st.write(f"Φάκελος εξόδου: `{WEEKLY_OUTPUT_DIR / week_tag}`")

    st.markdown("---")
    if weekly_dates_ok and weekly_range_ok and weekly_full_week_ok:
        week_tag = _week_tag(week_start_dt)
        week_out = WEEKLY_OUTPUT_DIR / week_tag
        period_key = f"{week_start_value}__{week_end_value}"
        out_xlsx, out_pdf = st.tabs(["Excel (XLSX)", "PDF"])
        with out_xlsx:
            _render_artifacts_panel(
                step_key="weekly_timologia",
                title="Διαθέσιμα αρχεία Excel",
                period_key=period_key,
                dir_path=week_out / "XLSX",
                zip_name=f"{week_tag}_weekly_excel.zip",
                allowed_suffixes={".xlsx"},
                panel_key_suffix="weekly_xlsx",
            )
        with out_pdf:
            _render_artifacts_panel(
                step_key="weekly_timologia",
                title="Διαθέσιμα αρχεία PDF",
                period_key=period_key,
                dir_path=week_out / "PDF",
                zip_name=f"{week_tag}_weekly_pdf.zip",
                allowed_suffixes={".pdf"},
                panel_key_suffix="weekly_pdf",
            )

        st.markdown("---")
        all_zip, all_count = _zip_dir_bytes(
            week_out,
            include_root_folder=True,
            allowed_suffixes={".xlsx", ".pdf"},
        )
        if all_zip:
            st.download_button(
                label=f"Κατέβασμα όλων (Excel + PDF) σε zip ({all_count} αρχεία)",
                data=all_zip,
                file_name=f"{week_tag}_weekly_excel_pdf.zip",
                mime="application/zip",
                width="stretch",
                key=f"dl_weekly_all_{period_key}",
            )
        else:
            st.caption("Δεν υπάρχουν διαθέσιμα Excel/PDF αρχεία για συνολικό zip.")
    else:
        st.info("Δώσε έγκυρο weekly διάστημα για να εμφανιστούν τα παραδοτέα.")
