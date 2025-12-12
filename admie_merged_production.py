# -*- coding: utf-8 -*-
import os
import re
import pandas as pd
from collections import defaultdict

# ====== ΡΥΘΜΙΣΕΙΣ LOGS ======
LOG_BASE = "logs/merged_production"
os.makedirs(LOG_BASE, exist_ok=True)

for filename in ["load_producers", "get_latest_green_ve6_files", "preprocess_timestamp_column", "assign_month_column", "merge_with_existing_csv", "process_file", "split_files_by_code"]:
    log_path = os.path.join(LOG_BASE, f"{filename}.txt")
    open(log_path, "w", encoding="utf-8").close()

def log(function_name, message):
    """Απλή συνάρτηση καταγραφής σε αρχείο ανά function."""
    log_path = os.path.join(LOG_BASE, f"{function_name}.txt")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(message + "\n")

# ====== HELPERS ======
def load_producers(filepath):
    function_name = "load_producers"
    try:
        df = pd.read_excel(filepath, dtype={'Code': str})
        if 'Code' not in df.columns or 'Εταιρεία' not in df.columns:
            raise ValueError("Το αρχείο δεν περιέχει τις στήλες 'Code' και 'Εταιρεία'")
        df['Code'] = df['Code'].astype(str).str.strip()
        df['Εταιρεία'] = df['Εταιρεία'].astype(str).str.strip()
        log(function_name, "Καταγεγραμμένοι παραγωγοί:")
        for _, row in df.iterrows():
            log(function_name, f"Code: {row['Code']} -> Εταιρεία: {row['Εταιρεία']}")
        log(function_name, f"Συνολικά φορτώθηκαν {len(df)} παραγωγοί.")
        return df
    except Exception as e:
        log(function_name, f"Σφάλμα: {e}")
        return None

def get_latest_green_ve6_files(folder):
    """
    Επιστρέφει τη νεότερη έκδοση ανά ημερομηνία για αρχεία που μοιάζουν με:
      GREEN_VE6YYYYMMDD.csv
      GREEN_VE6_YYYYMMDD.csv
      GREEN_VE6YYYYMMDD1.csv
      GREEN_VE6_YYYYMMDD_2.csv
    κ.λπ. (ανθεκτικό regex).
    Αν δεν γίνει match, ως fallback ταξινομεί αλφαβητικά και τα παίρνει όλα.
    """
    function_name = "get_latest_green_ve6_files"
    date_to_file = defaultdict(list)
    csv_files = [f for f in os.listdir(folder) if f.startswith("GREEN_VE6") and f.endswith(".csv")]
    log(function_name, f"Βρέθηκαν {len(csv_files)} αρχεία στον φάκελο {folder}.")
    for filename in csv_files:
        match = re.match(r"GREEN_VE6(\d{8})(\d)\.csv", filename)
        if match:
            date = match.group(1)
            edition = int(match.group(2))
            date_to_file[date].append((edition, filename))
    latest_files = []
    for date in sorted(date_to_file.keys()):
        files = date_to_file[date]
        latest = sorted(files, reverse=True)[0]
        log(function_name, f"{date} -> έκδοση {latest[0]}: {latest[1]}")
        latest_files.append(latest[1])
    return latest_files

# ====== ΕΠΕΞΕΡΓΑΣΙΑ TIMESTAMP ======
def preprocess_timestamp_column(df):
    function_name = "preprocess_timestamp_column"
    if 'TIMESTAMP' not in df.columns:
        log(function_name, "Λείπει η στήλη TIMESTAMP.")
        raise ValueError("Λείπει η στήλη TIMESTAMP")
    is_24 = df['TIMESTAMP'].str.contains('24:00', regex=False)
    new_timestamps = df['TIMESTAMP'].copy()
    new_timestamps[is_24] = (
        pd.to_datetime(df.loc[is_24, 'TIMESTAMP'].str.replace('24:00', '00:00'),
                       format='%d/%m/%Y %H:%M', errors='coerce') + pd.Timedelta(days=1)
    ).dt.strftime('%d/%m/%Y %H:%M')
    df['TIMESTAMP'] = new_timestamps
    df['datetime'] = pd.to_datetime(df['TIMESTAMP'], format='%d/%m/%Y %H:%M', errors='coerce')
    log(function_name, f"Επεξεργάστηκαν {len(df)} TIMESTAMPs.")
    return df

def assign_month_column(df):
    function_name = "assign_month_column"
    at_month_start = (
        (df['datetime'].dt.day == 1) &
        (df['datetime'].dt.hour == 0) &
        (df['datetime'].dt.minute == 0)
    )
    df['Μήνας'] = df['datetime'].dt.to_period('M').astype(str)
    df.loc[at_month_start, 'Μήνας'] = (
        df.loc[at_month_start, 'datetime'] - pd.DateOffset(days=1)
    ).dt.to_period('M').astype(str)
    log(function_name, f"Προστέθηκε στήλη Μήνας σε {len(df)} εγγραφές.")
    return df

def safe_company_folder_name(name):
    return re.sub(r'[\\/*?:"<>|]', "", name.replace(" ", "_"))

# ====== ΣΥΓΧΩΝΕΥΣΗ ΜΕ ΥΠΑΡΧΟΝ CSV ======
def merge_with_existing_csv(group_df, out_file):
    function_name = "merge_with_existing_csv"
    group_df = group_df.copy()
    group_df.set_index('TIMESTAMP', inplace=True)
    if os.path.exists(out_file):
        try:
            existing_df = pd.read_csv(out_file, delimiter=';', encoding='utf-8-sig')
            existing_df.set_index('TIMESTAMP', inplace=True)
            combined_df = existing_df[~existing_df.index.isin(group_df.index)]
            combined_df = pd.concat([combined_df, group_df])
            log(function_name, f"Συγχώνευση με υπάρχον αρχείο: {out_file}")
        except Exception as e:
            log(function_name, f"Σφάλμα ανάγνωσης υπάρχοντος αρχείου {out_file}: {e}")
            combined_df = group_df
    else:
        log(function_name, f"Δημιουργία νέου αρχείου: {out_file}")
        combined_df = group_df
    combined_df = combined_df.reset_index()
    combined_df['datetime'] = pd.to_datetime(combined_df['TIMESTAMP'], format='%d/%m/%Y %H:%M', errors='coerce')
    combined_df = combined_df.sort_values('datetime').drop(columns=['datetime'])
    return combined_df

def process_file(filepath, producers_df, output_folder):
    """
    - Διαβάζει το CSV (smart)
    - Ελέγχει ότι υπάρχει 'ΚΩΔΙΚΟΣ ΕΔΡΕΘ'
    - Προεπεξεργάζεται TIMESTAMP, προσθέτει 'Μήνας'
    - Για κάθε ΕΔΡΕΘ (Code), βρίσκει την εταιρεία από producers.xlsx
    - Αποθηκεύει σε: ΠΑΡΑΓΩΓΗ/{Εταιρεία}/ΠΑΡΑΓΩΓΗ_{Εταιρεία}.csv (append/merge)
    """
    function_name = "process_file"
    log(function_name, f"Επεξεργασία αρχείου: {filepath}")
    try:
        df = pd.read_csv(filepath, delimiter=';', encoding='utf-8-sig', skiprows=1)
    except Exception as e:
        log(function_name, f"Σφάλμα ανάγνωσης: {e}")
        return
    if 'ΚΩΔΙΚΟΣ ΕΔΡΕΘ' not in df.columns:
        log(function_name, "Λείπει η στήλη ΚΩΔΙΚΟΣ ΕΔΡΕΘ.")
        return
    # TIMESTAMP -> datetime
    try:
        df = preprocess_timestamp_column(df)
    except Exception as e:
        log(function_name, f"Σφάλμα TIMESTAMP: {e}")
        return
    df = assign_month_column(df)
    
    # Για κάθε κωδικό ΕΔΡΕΘ, γράψε στην εταιρεία του
    for code_value, group in df.groupby('ΚΩΔΙΚΟΣ ΕΔΡΕΘ'):
        code_str = str(code_value).strip()
        producers_df['Code'] = producers_df['Code'].astype(str).str.strip()
        producer_row = producers_df[producers_df['Code'] == code_str]
        if producer_row.empty:
            log(function_name, f"Άγνωστος ΚΩΔΙΚΟΣ ΕΔΡΕΘ: {code_str}")
            continue
        company_name = producer_row['Εταιρεία'].values[0]
        safe_name = safe_company_folder_name(company_name)
        out_file = os.path.join(output_folder, f"ΠΑΡΑΓΩΓΗ_{safe_name}.csv")
        final_df = merge_with_existing_csv(group, out_file)
        final_df.to_csv(out_file, index=False, sep=';', encoding='utf-8-sig')
        log(function_name, f"Αποθήκευση για {company_name} ({code_str}) στο {out_file}")

# ====== ORCHESTRATOR ======
def split_files_by_code(month):
    """
    - Παίρνει τον φάκελο downloads/{YYYY-MM}
    - Φορτώνει producers.xlsx
    - Βρίσκει τα τελευταία GREEN_VE6 CSV ανά ημέρα
    - Τα περνάει από process_file
    - Γράφει λογιστικά logs στο logs/merged_production
    """
    input_folder = f'downloads/{month}'
    output_folder = 'ΠΑΡΑΓΩΓΗ'
    os.makedirs(output_folder, exist_ok=True)
    producers_df = load_producers('producers.xlsx')
    if producers_df is None:
        return
    latest_files = get_latest_green_ve6_files(input_folder)
    for filename in latest_files:
        file_path = os.path.join(input_folder, filename)
        process_file(file_path, producers_df, output_folder)
    log("split_files_by_code", "Ο διαχωρισμός και η επεξεργασία ολοκληρώθηκαν.")

# ====== ENTRYPOINT ======
if __name__ == '__main__':
    month_input = input("Δώσε τον χρονο και τον μήνα σε μορφή YYYY-MM (π.χ. 2025-01): ").strip()
    if not re.match(r'^\d{4}-(0[1-9]|1[0-2])$', month_input):
        print("Μη έγκυρη μορφή. Χρησιμοποίησε YYYY-MM (π.χ. 2025-01).")
    else:
        split_files_by_code(month_input)
