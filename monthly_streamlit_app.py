#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Streamlit UI για πλήρες pipeline:
1) Download GREEN_VE6 από Modesto
2) Merge παραγωγής ανά πάρκο
3) Εξαγωγή τιμολογίων (XLSX + PDF)

Τρέξιμο:
    streamlit run monthly_streamlit_app.py
"""

import re
import subprocess
import sys
import os
import io
import json
import zipfile
from calendar import monthrange
from datetime import date, datetime
from pathlib import Path

import streamlit as st

from MONTHLY.config import OUTPUT_DIR


MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
BASE_DIR = Path(__file__).resolve().parent
HISTORY_FILE = BASE_DIR / "logs" / "streamlit_pipeline_history.json"


def _default_month() -> str:
    return date.today().strftime("%Y-%m")


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


def _run_step_1_modesto(year: int, month_num: int, start_day: int, end_day: int):
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
    ]
    cmd.append("--verify-ssl")
    return _run_command(cmd)


def _run_step_2_merge_production(month_value: str):
    code = "import os, admie_merged_production as m; m.split_files_by_code(os.environ['PIPE_MONTH'])"
    cmd = [sys.executable, "-c", code]
    return _run_command(cmd, env_extra={"PIPE_MONTH": month_value})


def _run_step_3_timologia(month_value: str):
    code = "import os; from MONTHLY.timologia import timologia; timologia(os.environ['PIPE_MONTH'])"
    cmd = [sys.executable, "-c", code]
    return _run_command(cmd, env_extra={"PIPE_MONTH": month_value})


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
        if isinstance(data, list):
            return data
        return []
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
    for p, mtime in after.items():
        if before.get(p) != mtime:
            try:
                changed.append(str(Path(p).resolve().relative_to(BASE_DIR)))
            except Exception:
                changed.append(p)
    changed.sort()
    return changed


def _record_step_history(step_key: str, month_value: str, ok: bool, output: str, artifacts: list[Path], changed_files: list[str]):
    entries = _load_history()
    entries.append({
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "step": step_key,
        "month": month_value,
        "ok": ok,
        "artifacts": [str(p) for p in artifacts],
        "changed_files": changed_files,
        "output_tail": (output or "")[-3000:],
    })
    _save_history(entries)


def _latest_step_history(step_key: str, month_value: str) -> dict | None:
    entries = _load_history()
    for item in reversed(entries):
        if item.get("step") == step_key and item.get("month") == month_value:
            return item
    return None


def _execute_step(step_key: str, month_value: str, artifacts: list[Path], runner):
    before = _snapshot_files(artifacts)
    ok, output = runner()
    after = _snapshot_files(artifacts)
    changed_files = _diff_snapshots(before, after)
    _record_step_history(step_key, month_value, ok, output, artifacts, changed_files)
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
            if include_root_folder:
                arcname = str(Path(dir_path.name) / rel)
            else:
                arcname = str(rel)
            zf.write(fpath, arcname)
    return buf.getvalue(), len(files)


def _render_artifacts_panel(
    step_key: str,
    title: str,
    month_value: str,
    dir_path: Path,
    zip_name: str,
    allowed_suffixes: set[str] | None = None,
    panel_key_suffix: str = "",
):
    st.write(f"**{title}**")
    st.write(f"Φάκελος: `{dir_path}`")

    latest = _latest_step_history(step_key, month_value)
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
        download_key = f"dl_{step_key}_{month_value}_{panel_key_suffix or 'all'}"
        st.download_button(
            label=f"Κατέβασμα zip ({count} αρχεία)",
            data=zip_data,
            file_name=zip_name,
            mime="application/zip",
            width="stretch",
            key=download_key,
        )
    else:
        st.caption("Δεν υπάρχουν αρχεία για zip.")

    rows = _list_files_for_view(dir_path, allowed_suffixes=allowed_suffixes)
    if rows:
        st.dataframe(rows, width="stretch", hide_index=True)
    else:
        st.caption("Δεν βρέθηκαν αρχεία.")


st.set_page_config(
    page_title="ΜΗΝΙΑΙΑ ΕΝΗΜΕΡΩΤΙΚΑ ΣΗΜΕΙΩΜΑΤΑ",
    page_icon="M",
    layout="wide",
)

st.title("ΜΗΝΙΑΙΑ ΕΝΗΜΕΡΩΤΙΚΑ ΣΗΜΕΙΩΜΑΤΑ")
st.caption("Όλα τα βήματα: Modesto -> Παραγωγή ανά πάρκο -> Ενημερωτικά Σημειώματα")

month = st.text_input("Μήνας (YYYY-MM)", value=_default_month()).strip()
month_ok = bool(MONTH_RE.match(month))

if not month_ok:
    st.warning("Μη έγκυρη μορφή μήνα. Χρησιμοποίησε YYYY-MM (π.χ. 2026-01).")

if month_ok:
    year = int(month.split("-")[0])
    month_num = int(month.split("-")[1])
    last_day = monthrange(year, month_num)[1]
else:
    today = date.today()
    year, month_num = today.year, today.month
    last_day = monthrange(year, month_num)[1]

st.markdown("---")
st.subheader("Βήμα 1: Κατέβασμα αρχείων από Modesto")
step1_left, step1_right = st.columns([1, 1], gap="large")
with step1_left:
    st.caption(
        "Κατέβασε τα αρχεία με τα στοιχεία αγοράς από το Modesto για τις ημέρες που θέλεις. "
        "Αυτό είναι το πρώτο βήμα για να προχωρήσει η μηνιαία διαδικασία."
    )

    col1, col2 = st.columns(2)
    with col1:
        start_day = st.number_input("Από ημέρα", min_value=1, max_value=last_day, value=1, step=1)
    with col2:
        end_day_default = min(last_day, 10)
        end_day = st.number_input("Έως ημέρα", min_value=1, max_value=last_day, value=end_day_default, step=1)

    if end_day < start_day:
        st.error("Το 'Έως ημέρα' πρέπει να είναι >= από το 'Από ημέρα'.")
    elif (end_day - start_day) > 9:
        st.error("Το εύρος ημερών στο Modesto πρέπει να είναι έως 10 ημέρες.")

    run_step1 = st.button("Κατέβασε αρχεία Modesto", width="stretch")
    if run_step1:
        if not month_ok:
            st.error("Διόρθωσε πρώτα τον μήνα pipeline.")
        elif end_day < start_day or (end_day - start_day) > 9:
            st.error("Διόρθωσε πρώτα το εύρος ημερών του Βήματος 1.")
        else:
            step1_artifacts = [BASE_DIR / "downloads" / month]
            with st.spinner("Τρέχει το Βήμα 1..."):
                ok, output, changed = _execute_step(
                    step_key="step1_modesto",
                    month_value=month,
                    artifacts=step1_artifacts,
                    runner=lambda: _run_step_1_modesto(year, month_num, int(start_day), int(end_day)),
                )
            _show_output("Βήμα 1 (Modesto)", ok, output)
            _show_changed_files(changed, allowed_suffixes={".csv"})

with step1_right:
    if month_ok:
        _render_artifacts_panel(
            step_key="step1_modesto",
            title="Διαθέσιμα αρχεία Modesto",
            month_value=month,
            dir_path=BASE_DIR / "downloads" / month,
            zip_name=f"{month}_step1_modesto.zip",
            allowed_suffixes={".csv"},
        )
    else:
        st.info("Δώσε έγκυρο μήνα για να εμφανιστούν αρχεία.")

st.markdown("---")
st.subheader("Βήμα 2: Παραγωγή ανά πάρκο")
step2_left, step2_right = st.columns([1, 1], gap="large")
with step2_left:
    st.caption(
        "Το σύστημα οργανώνει τα δεδομένα παραγωγής και τα χωρίζει αυτόματα ανά πάρκο, "
        "ώστε να είναι έτοιμα για τον τελικό υπολογισμό."
    )
    run_step2 = st.button("Βάλε την παραγωγή ανά πάρκο", width="stretch")
    if run_step2:
        if not month_ok:
            st.error("Διόρθωσε πρώτα τον μήνα pipeline.")
        else:
            step2_artifacts = [BASE_DIR / "ΠΑΡΑΓΩΓΗ"]
            with st.spinner("Τρέχει το Βήμα 2..."):
                ok, output, changed = _execute_step(
                    step_key="step2_production",
                    month_value=month,
                    artifacts=step2_artifacts,
                    runner=lambda: _run_step_2_merge_production(month),
                )
            _show_output("Βήμα 2 (Merge παραγωγής)", ok, output)
            _show_changed_files(changed)
with step2_right:
    if month_ok:
        _render_artifacts_panel(
            step_key="step2_production",
            title="Διαθέσιμα αρχεία παραγωγής ανά πάρκο",
            month_value=month,
            dir_path=BASE_DIR / "ΠΑΡΑΓΩΓΗ",
            zip_name=f"{month}_step2_production.zip",
        )
    else:
        st.info("Δώσε έγκυρο μήνα για να εμφανιστούν αρχεία.")

st.markdown("---")
st.subheader("Βήμα 3: Εξαγωγή τιμολογίων")
step3_left, step3_right = st.columns([1, 1], gap="large")
with step3_left:
    st.caption(
        "Δημιούργησε τα τελικά ενημερωτικά σημειώματα του μήνα για κάθε πάρκο "
        "σε Excel και PDF, έτοιμα για αποστολή ή αρχειοθέτηση."
    )
    run_step3 = st.button("Βγάλε ενημερωτικά σημειώματα", type="primary", width="stretch")
    if run_step3:
        if not month_ok:
            st.error("Διόρθωσε πρώτα τον μήνα pipeline.")
        else:
            step3_artifacts = [Path(OUTPUT_DIR) / month]
            with st.spinner("Τρέχει το Βήμα 3..."):
                ok, output, changed = _execute_step(
                    step_key="step3_timologia",
                    month_value=month,
                    artifacts=step3_artifacts,
                    runner=lambda: _run_step_3_timologia(month),
                )
            _show_output("Βήμα 3 (Τιμολόγια)", ok, output)
            _show_changed_files(changed)
            if ok:
                st.write(f"Φάκελος εξόδου: `{Path(OUTPUT_DIR) / month}`")
with step3_right:
    if month_ok:
        month_out = Path(OUTPUT_DIR) / month
        tab_xlsx, tab_pdf = st.tabs(["Excel (XLSX)", "PDF"])
        with tab_xlsx:
            _render_artifacts_panel(
                step_key="step3_timologia",
                title="Διαθέσιμα αρχεία Excel",
                month_value=month,
                dir_path=month_out / "XLSX",
                zip_name=f"{month}_step3_excel.zip",
                allowed_suffixes={".xlsx"},
                panel_key_suffix="xlsx",
            )
        with tab_pdf:
            _render_artifacts_panel(
                step_key="step3_timologia",
                title="Διαθέσιμα αρχεία PDF",
                month_value=month,
                dir_path=month_out / "PDF",
                zip_name=f"{month}_step3_pdf.zip",
                allowed_suffixes={".pdf"},
                panel_key_suffix="pdf",
            )

        st.markdown("---")
        all_step3_zip, all_step3_count = _zip_dir_bytes(
            month_out,
            include_root_folder=True,
            allowed_suffixes={".xlsx", ".pdf"},
        )
        if all_step3_zip:
            st.download_button(
                label=f"Κατέβασμα όλων (Excel + PDF) σε zip ({all_step3_count} αρχεία)",
                data=all_step3_zip,
                file_name=f"{month}_step3_excel_pdf.zip",
                mime="application/zip",
                width="stretch",
                key=f"dl_step3_all_{month}",
            )
        else:
            st.caption("Δεν υπάρχουν διαθέσιμα Excel/PDF αρχεία για συνολικό zip.")
    else:
        st.info("Δώσε έγκυρο μήνα για να εμφανιστούν αρχεία.")
