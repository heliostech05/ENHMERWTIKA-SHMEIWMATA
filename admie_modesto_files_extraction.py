#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Modesto GREEN_VE6 Downloader (macOS & Windows friendly)
-------------------------------------------------------
- Pulls GREEN_VE6 messages from ADMIE Modesto Web Service via 2-way TLS.
- Saves both XML and decoded CSV for each message in downloads/YYYY-MM/.
- Robust features:
  * CLI flags (non-interactive) or interactive prompts (Greek).
  * Retries with exponential backoff on SOAP calls.
  * Handles Base64 payloads that are either plain UTF-8 text or gzipped.
  * Safer CSV parsing (supports quoted fields).
  * Clear logging.

USAGE (interactive): 
    python admie_modesto_files_extraction.py --verify-ssl

USAGE (non-interactive example):
    python admie_modesto_files_extraction.py \
        --year 2025 --month 3 --start-day 5 --end-day 14 \
        --cert ./certificates/client_modesto_cert.pem \
        --key  ./certificates/client_modesto_key.pem \
        --verify-ssl \
        --wsdl https://market-extranet-api.admie.gr/modestoWS/Service_EME_Port?wsdl \
        --out downloads \
        --verbose
"""

import argparse
import base64
import csv
import gzip
import io
import logging
import os
import sys
import time
from calendar import monthrange
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from zeep.helpers import serialize_object

import requests
from requests import Session
from lxml import etree
from zeep import Client, Settings
from zeep.transports import Transport


# =========================
# Helpers: CLI + Validation
# =========================

def parse_args() -> argparse.Namespace:
    """Ορίζει όλα τα CLI flags."""
    p = argparse.ArgumentParser(
        description="Download GREEN_VE6 files from ADMIE Modesto WS and save XML + CSV."
    )
    
    # Ημερομηνίες
    p.add_argument("--year", type=int, help="Έτος (e.g., 2025).")
    p.add_argument("--month", type=int, help="Μήνας 1-12.")
    p.add_argument("--start-day", type=int, dest="start_day", help="Αρχική ημέρα (1..last_day).")
    p.add_argument("--end-day", type=int, dest="end_day", help="Τελική ημέρα (>= start_day, max 10 μέρες διαφορά).")
    
    # Πιστοποιητικά / ασφάλεια
    p.add_argument("--cert", default="./certificates/client_modesto_cert.pem",
                   help="Client certificate (PEM). Default: ./certificates/client_modesto_cert.pem")
    p.add_argument("--key", default="./certificates/client_modesto_key.pem",
                   help="Client private key (PEM). Default: ./certificates/client_modesto_key.pem")
    p.add_argument("--verify-ssl", action="store_true", help="Verify server SSL certificate (recommended).")
    p.add_argument("--insecure", action="store_true", help="Disable SSL verification (NOT recommended).")

    # Modesto WSDL & δίκτυο
    p.add_argument("--wsdl", default="https://market-extranet-api.admie.gr/modestoWS/Service_EME_Port?wsdl",
                   help="Modesto WSDL URL.")
    p.add_argument("--max-retries", type=int, default=3, help="Max retries per SOAP call. Default: 3")
    p.add_argument("--timeout", type=int, default=60, help="HTTP timeout per call (seconds). Default: 60")
    
    # Έξοδος & logging
    p.add_argument("--out", default="downloads", help="Base downloads folder. Default: downloads")
    p.add_argument("--verbose", action="store_true", help="Verbose logging.")

    return p.parse_args()


def greek_interactive_prompts() -> Tuple[int, int, int, int]:
    """Διαδραστικά prompts στα ελληνικά για έτος/μήνα/ημέρες, με validation."""
    while True:
        try:
            year = int(input("Δώσε το έτος (π.χ. 2025): "))
            if 2000 <= year <= 2100:
                break
            else:
                print("Έτος εκτός ορίων. Δώσε έτος μεταξύ 2000 και 2100.")
        except ValueError:
            print("Μη έγκυρη τιμή. Βάλε έναν αριθμό.")

    while True:
        try:
            selected_month = int(input("Δώσε τον αριθμό του μήνα (1-12): "))
            if 1 <= selected_month <= 12:
                break
            else:
                print("Μήνας εκτός ορίων.")
        except ValueError:
            print("Μη έγκυρη τιμή. Βάλε έναν αριθμό.")

    _, last_day = monthrange(year, selected_month)

    while True:
        try:
            start_day = int(input(f"Διάλεξε την **αρχική ημέρα** του μήνα (1-{last_day}): "))
            if 1 <= start_day <= last_day:
                break
            else:
                print(f"Η ημέρα πρέπει να είναι από 1 έως {last_day}.")
        except ValueError:
            print("Μη έγκυρη τιμή. Βάλε έναν αριθμό.")

    while True:
        try:
            end_day = int(input(f"Διάλεξε την **τελική ημέρα** του μήνα (>= {start_day}, max 10 μέρες διαφορά): "))
            if start_day <= end_day <= last_day and (end_day - start_day) <= 9:
                break
            else:
                print(f"Η τελική ημέρα πρέπει να είναι από {start_day} έως {min(start_day + 9, last_day)}.")
        except ValueError:
            print("Μη έγκυρη τιμή. Βάλε έναν αριθμό.")

    return year, selected_month, start_day, end_day


def validate_or_prompt_dates(ns: argparse.Namespace) -> Tuple[datetime, datetime, str]:
    if ns.year and ns.month and ns.start_day and ns.end_day:
        year = ns.year
        selected_month = ns.month
        start_day = ns.start_day
        end_day = ns.end_day

        # Validate ranges
        if not (2000 <= year <= 2100):
            raise ValueError("Έτος εκτός ορίων (2000..2100).")
        if not (1 <= selected_month <= 12):
            raise ValueError("Μήνας εκτός ορίων (1..12).")

        _, last_day = monthrange(year, selected_month)
        if not (1 <= start_day <= last_day):
            raise ValueError(f"Αρχική ημέρα πρέπει 1..{last_day}.")
        if not (start_day <= end_day <= last_day):
            raise ValueError(f"Τελική ημέρα πρέπει {start_day}..{last_day}.")
        if (end_day - start_day) > 9:
            raise ValueError("Το εύρος πρέπει να είναι max 10 μέρες (διαφορά ≤ 9).")
    else:
        year, selected_month, start_day, end_day = greek_interactive_prompts()

    start_time = datetime(year, selected_month, start_day, 0, 0, 0, tzinfo=timezone.utc)
    end_time = datetime(year, selected_month, end_day, 23, 59, 59, tzinfo=timezone.utc)
    date_folder = f"{start_time.year}-{start_time.month:02}"

    print(f"Θα κατεβούν αρχεία από {start_time.date()} έως {end_time.date()}.")

    return start_time, end_time, date_folder


# =========================
# SOAP (zeep) Client + Retries
# =========================

def build_zeep_client(wsdl_url: str, cert_file: str, key_file: str, verify_ssl: bool, timeout: int) -> Client:
      
    """
    Φτιάχνει Zeep Client με session που περιέχει client cert + key.
    - verify_ssl=True: κανονική επαλήθευση πιστοποιητικών server
    - verify_ssl=False: απενεργοποιημένη (όχι ασφαλές) 
    """
    
    # 🔍 ΠΡΟΣΘΗΚΗ LOG για να δούμε ΑΚΡΙΒΩΣ ποια αρχεία παίρνει
    logging.info("Using client cert: %s", os.path.abspath(cert_file))
    logging.info("Using client key : %s", os.path.abspath(key_file))
    
    session = Session()
    session.cert = (cert_file, key_file)
    session.verify = verify_ssl
    session.timeout = timeout

    transport = Transport(session=session, timeout=timeout)
    settings = Settings(strict=False, xml_huge_tree=True)

    client = Client(wsdl=wsdl_url, transport=transport, settings=settings)
    logging.info("Connected to Modesto Web Service.")
    return client


def request_with_retries(client: Client, header, request, payload, max_retries: int = 3, base_delay: float = 1.0):
    """Call `client.service.request()` with simple exponential backoff retries."""
    attempt = 0
    while True:
        try:
            return client.service.request(Header=header, Request=request, Payload=payload)
        except Exception as e:
            attempt += 1
            if attempt > max_retries:
                logging.error("Request failed after %d attempts: %s", attempt - 1, str(e))
                raise
            delay = base_delay * (2 ** (attempt - 1))
            logging.warning("Request error (%s). Retrying in %.1fs (attempt %d/%d)...",
                            str(e), delay, attempt, max_retries)
            time.sleep(delay)


# =========================
# GREEN_VE6 parsing/saving
# =========================

def ensure_dir(path: Path):
    """Δημιουργεί φάκελο/out dirs αν δεν υπάρχουν."""
    path.mkdir(parents=True, exist_ok=True)


def to_pretty_xml_string(element) -> str:
    """Μετατρέπει Zeep/LXML element σε όμορφο XML string (utf-8)."""
    xml_bytes = etree.tostring(element, pretty_print=True, encoding='utf-8')
    return xml_bytes.decode('utf-8')


def decode_base64_payload_to_text(b64_text: str) -> str:
    """
    Αποκωδικοποιεί Base64 payload που είναι:
      - είτε καθαρό UTF-8 κείμενο,
      - είτε gzip-συμπιεσμένο UTF-8 κείμενο.
    """
    raw = base64.b64decode(b64_text.strip())

    # Try straight UTF-8
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        pass

    # Try gzip
    try:
        return gzip.decompress(raw).decode("utf-8")
    except Exception as e:
        # As a last resort, attempt to detect gzip via file-like
        try:
            with gzip.GzipFile(fileobj=io.BytesIO(raw)) as gz:
                return gz.read().decode("utf-8")
        except Exception:
            raise ValueError(f"Failed to decode Base64 payload as UTF-8 or GZip UTF-8: {e}")


def write_csv_safely(decoded_text: str, csv_path: Path):
    """
    Γράφει CSV με csv.reader ώστε να υποστηρίζει σωστά quoted πεδία/κόμματα.
    """
    reader = csv.reader(io.StringIO(decoded_text))
    with csv_path.open(mode="w", newline="", encoding="utf-8") as f_out:
        writer = csv.writer(f_out)
        writer.writerows(reader)

# =========================
# Parsing MessageList & Download
# =========================

def parse_messagelist_and_collect_green_ve6(payload) -> List[Dict]:
    """
    Από το Payload (dict) βρίσκει MessageList -> Message -> Code/MessageIdentification
    και επιστρέφει μόνο όσα περιέχουν 'GREEN_VE6'.
    """
    messages = []
    if not payload or not payload.get('_value_1'):
        return messages

    for element in payload['_value_1']:
        # Find MessageList elements, then Message children in the expected namespace
        if element.tag.endswith('MessageList'):
            for message in element.findall('{urn:iec62325.504:messages:1:0}Message'):
                code_el = message.find('{urn:iec62325.504:messages:1:0}Code')
                ident_el = message.find('{urn:iec62325.504:messages:1:0}MessageIdentification')
                code_val = int(code_el.text) if (code_el is not None and code_el.text and code_el.text.isdigit()) else None
                ident_val = ident_el.text if ident_el is not None else None

                if ident_val and "GREEN_VE6" in ident_val:
                    messages.append({
                        'Code': code_val,
                        'MessageIdentification': ident_val
                    })
    # Sort by Code if present; keep original order as fallback
    messages.sort(key=lambda x: (x['Code'] is None, x['Code']))
    return messages


def download_and_save_message_by_code(client: Client,
                                      factory,
                                      code_value: int,
                                      out_dir: Path,
                                      message_id_hint: Optional[str] = None,
                                      max_retries: int = 3) -> Optional[Path]:
    """
    Κατεβάζει το πλήρες μήνυμα βάσει Code, σώζει XML + CSV.
    Επιστρέφει το μονοπάτι του CSV ή None εάν δεν βρέθηκε payload.
    """
    # Header για "Any" (λήψη συγκεκριμένου μηνύματος μέσω Option Code)
    header = factory.HeaderType(
        Verb="get",
        Noun="Any",
        Revision="1.0",
        Context="PRODUCTION",
        Timestamp=datetime.now(timezone.utc).isoformat(),
    )

    # Επιλογή βάσει Code
    option = factory.OptionType(
        name="Code",
        value=code_value
    )

    request = factory.RequestType(
        Option=[option],
        ID=[],
        _value_1=[]
    )

    payload = factory.PayloadType(
        _value_1=[],
        Format="xml"
    )

    # Κλήση με retries
    full_response = request_with_retries(client, header, request, payload,
                                         max_retries=max_retries, base_delay=1.0)
    # Ζήτα dict για ασφαλή .get
    full_resp_dict = serialize_object(full_response)
    full_payload = full_resp_dict.get('Payload')

    if not full_payload or not full_payload.get('_value_1'):
        logging.warning("No payload found for code=%s", code_value)
        return None

    # Το πρώτο στοιχείο είναι το XML element με μέσα base64 text
    xml_element = full_payload['_value_1'][0]
    xml_string = to_pretty_xml_string(xml_element)

    # Όνομα αρχείων (XML/CSV)
    filename_stem = message_id_hint or f"GREEN_VE6_{code_value}"

    # Save XML
    xml_path = out_dir / f"{filename_stem}.xml"
    xml_path.write_text(xml_string, encoding="utf-8")
    logging.info("Saved XML: %s", xml_path)

    # Από το xml_string πάρε το text και κάνε base64
    try:
        tree = etree.fromstring(xml_string.encode('utf-8'))
        if tree.text is None:
            logging.warning("XML for code=%s had no text payload.", code_value)
            return None
        decoded_text = decode_base64_payload_to_text(tree.text)
    except Exception as e:
        logging.error("Failed to decode Base64 for code=%s: %s", code_value, e)
        return None

    # Save CSV
    csv_path = out_dir / f"{filename_stem}.csv"
    write_csv_safely(decoded_text, csv_path)
    logging.info("Saved CSV: %s", csv_path)
    return csv_path


# =========================
# Main flow
# =========================

def main():
    ns = parse_args()

    # Logging setup
    log_level = logging.DEBUG if ns.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # SSL verification
    if ns.insecure and ns.verify_ssl:
        logging.warning("--verify-ssl and --insecure both given; proceeding as INSECURE (verify=False).")
    verify_ssl = False if ns.insecure else bool(ns.verify_ssl)

    # Prepare date range
    try:
        start_time, end_time, date_folder = validate_or_prompt_dates(ns)
    except Exception as e:
        logging.error("Σφάλμα στις ημερομηνίες: %s", e)
        sys.exit(1)

    # Prepare output dir
    base_out = Path(ns.out)
    out_dir = base_out / date_folder
    ensure_dir(out_dir)

    # Build zeep client
    try:
        client = build_zeep_client(ns.wsdl, ns.cert, ns.key, verify_ssl, ns.timeout)
    except Exception as e:
        logging.error("Failed to connect to Modesto Web Service: %s", e)
        sys.exit(1)

    # Type factory
    factory = client.type_factory('ns1')

    # Build MessageList request
    header = factory.HeaderType(
        Verb="get",
        Noun="MessageList",
        Revision="1.0",
        Context="PRODUCTION",
        Timestamp=datetime.now(timezone.utc).isoformat(),
    )

    option = factory.OptionType(
        name="IntervalType",
        value="Application"
    )

    request = factory.RequestType(
        StartTime=start_time.isoformat(),
        EndTime=end_time.isoformat(),
        Option=[option],
        ID=[],
        _value_1=[]
    )

    payload = factory.PayloadType(
        _value_1=[],
        Format="xml"
    )

    # Call MessageList with retries
    try:
        response = request_with_retries(client, header, request, payload,
                                        max_retries=ns.max_retries, base_delay=1.0)
    except Exception as e:
        logging.error("MessageList request failed: %s", e)
        sys.exit(1)

    # Parse MessageList and filter GREEN_VE6
    resp_dict = serialize_object(response)
    payload_obj = resp_dict.get('Payload')
    green_msgs = parse_messagelist_and_collect_green_ve6(payload_obj)

    if not green_msgs:
        print("No GREEN_VE6 messages found.")
        return

    # Sort already done; iterate and download
    total = len(green_msgs)
    logging.info("Found %d GREEN_VE6 messages in %s..%s.", total, start_time.date(), end_time.date())

    downloaded = 0
    for i, msg in enumerate(green_msgs, start=1):
        code = msg['Code']
        message_id = msg.get('MessageIdentification')
        if code is None:
            logging.warning("Skipping message without numeric Code: %s", message_id)
            continue

        logging.info("[%d/%d] Downloading Code=%s, MessageIdentification=%s", i, total, code, message_id)
        try:
            csv_path = download_and_save_message_by_code(
                client=client,
                factory=factory,
                code_value=code,
                out_dir=out_dir,
                message_id_hint=message_id,
                max_retries=ns.max_retries
            )
            if csv_path:
                downloaded += 1
        except Exception as e:
            logging.error("Error downloading Code=%s (%s): %s", code, message_id, e)

    logging.info("Done. %d/%d CSV files saved in: %s", downloaded, total, out_dir.resolve())


if __name__ == "__main__":
    main()