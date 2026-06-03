#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import gzip
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable
from urllib.parse import quote
from xml.etree import ElementTree as ET

import requests


ROOT_DIR = Path(__file__).resolve().parents[2]

SOAP_NS = "http://www.w3.org/2003/05/soap-envelope"
MES_NS = "http://iec.ch/TC57/2011/schema/message"
NS = {"soap": SOAP_NS, "mes": MES_NS}

MSG_TYPES = [
    "IMBAL",
    "BMFEE",
    "SUFEE",
    "SUFEE_DISCOUNTS",
    "AGGR",
    "UPLIFT",
    "BCRBC",
    "BERBE",
    "NCC",
]

DEFAULT_ENDPOINT = "https://market-extranet-api.admie.gr/settlement/OutboundMarketParticipantService"


def _qname(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _option(name: str, value: str) -> ET.Element:
    opt = ET.Element(_qname(MES_NS, "Option"))
    ET.SubElement(opt, _qname(MES_NS, "name")).text = name
    ET.SubElement(opt, _qname(MES_NS, "value")).text = value
    return opt


def _build_envelope(verb: str, noun: str, request_fields: Iterable[ET.Element]) -> bytes:
    envelope = ET.Element(_qname(SOAP_NS, "Envelope"))
    ET.SubElement(envelope, _qname(SOAP_NS, "Header"))
    body = ET.SubElement(envelope, _qname(SOAP_NS, "Body"))
    req_msg = ET.SubElement(body, _qname(MES_NS, "RequestMessage"))
    header = ET.SubElement(req_msg, _qname(MES_NS, "Header"))
    ET.SubElement(header, _qname(MES_NS, "Verb")).text = verb
    ET.SubElement(header, _qname(MES_NS, "Noun")).text = noun
    request = ET.SubElement(req_msg, _qname(MES_NS, "Request"))
    for field in request_fields:
        request.append(field)
    return ET.tostring(envelope, encoding="utf-8", xml_declaration=True)


@dataclass
class BaseWSClient:
    cert_path: str
    key_path: str
    endpoint_url: str = DEFAULT_ENDPOINT
    verify_ssl: bool = True
    timeout: int = 30

    def __post_init__(self) -> None:
        self.session = requests.Session()
        self.session.cert = (self.cert_path, self.key_path)
        self.session.verify = self.verify_ssl

    def _post(self, xml_body: bytes) -> requests.Response:
        headers = {
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": '"urn:iec62325.504:wss:1:0/port_TFEDI_type/requestRequest"',
        }
        return self.session.post(self.endpoint_url, data=xml_body, headers=headers, timeout=self.timeout)

    def list_message_results(
        self,
        start_time: str,
        end_time: str,
        msg_types: list[str],
        interval_type: str = "Application",
        sett_stage: str | None = None,
    ) -> requests.Response:
        fields = [ET.Element(_qname(MES_NS, "StartTime")), ET.Element(_qname(MES_NS, "EndTime"))]
        fields[0].text = start_time
        fields[1].text = end_time
        fields.append(_option("IntervalType", interval_type))
        for msg_type in msg_types:
            fields.append(_option("MsgType", msg_type))
        if sett_stage:
            fields.append(_option("SettStage", sett_stage))
        payload = _build_envelope("get", "MessageList", fields)
        return self._post(payload)

    def download_result(self, code: int | str) -> requests.Response:
        payload = _build_envelope("get", "Any", [_option("Code", str(code))])
        return self._post(payload)

    @staticmethod
    def parse_message_list(response_text: str) -> list[dict[str, str]]:
        root = ET.fromstring(response_text)
        msgs = []
        for msg in root.findall(".//{urn:iec62325.504:messages:1:0}Message"):
            item = {}
            for tag in ("Code", "MessageIdentification", "MessageVersion", "Status", "ServerTimestamp", "Type", "Owner"):
                el = msg.find(f"{{urn:iec62325.504:messages:1:0}}{tag}")
                if el is not None and el.text is not None:
                    item[tag] = el.text
            interval = msg.find("{urn:iec62325.504:messages:1:0}ApplicationTimeInterval")
            if interval is not None:
                st = interval.find("{urn:iec62325.504:messages:1:0}start")
                en = interval.find("{urn:iec62325.504:messages:1:0}end")
                if st is not None and st.text is not None:
                    item["start"] = st.text
                if en is not None and en.text is not None:
                    item["end"] = en.text
            msgs.append(item)
        return msgs

    @staticmethod
    def extract_compressed_payload(response_text: str) -> bytes:
        root = ET.fromstring(response_text)
        compressed = root.find(".//{http://iec.ch/TC57/2011/schema/message}Compressed")
        if compressed is None or compressed.text is None:
            raise ValueError("Compressed payload not found in SOAP response.")
        return base64.b64decode(compressed.text.strip())

    @staticmethod
    def decode_payload(blob: bytes) -> bytes:
        try:
            return gzip.decompress(blob)
        except OSError:
            return blob


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Interactive BaSE+ SOAP downloader.")
    p.add_argument("--endpoint", default=DEFAULT_ENDPOINT, help="SOAP endpoint URL.")
    p.add_argument(
        "--cert",
        default=ROOT_DIR / "certificates" / "client-EL801961185@settlement.admie.gr-2045.pem",
        type=Path,
        help="PEM client certificate.",
    )
    p.add_argument(
        "--key",
        default=ROOT_DIR / "certificates" / "client-EL801961185@settlement.admie.gr-2045.key",
        type=Path,
        help="PEM client key.",
    )
    p.add_argument("--insecure", action="store_true", help="Disable TLS verification.")
    p.add_argument("--timeout", type=int, default=30)
    p.add_argument("--outdir", type=Path, default=ROOT_DIR / "BASE" / "artifacts" / "downloads")
    return p.parse_args()


def _prompt_date(prompt: str) -> datetime:
    while True:
        raw = input(prompt).strip()
        try:
            return datetime.strptime(raw, "%Y-%m-%d")
        except ValueError:
            print("Invalid date. Use YYYY-MM-DD")


def _prompt_msg_types() -> list[str]:
    print("Available types:")
    for idx, msg_type in enumerate(MSG_TYPES, start=1):
        print(f"  {idx}. {msg_type}")
    print("Enter comma-separated numbers or names. Example: 1,3 or IMBAL,BMFEE")
    while True:
        raw = input("Types: ").strip()
        if not raw:
            print("Please select at least one type.")
            continue
        chosen: list[str] = []
        for token in [x.strip() for x in raw.split(",") if x.strip()]:
            if token.isdigit():
                idx = int(token)
                if 1 <= idx <= len(MSG_TYPES):
                    chosen.append(MSG_TYPES[idx - 1])
                else:
                    print(f"Invalid index: {idx}")
                    break
            else:
                token_u = token.upper()
                if token_u in MSG_TYPES:
                    chosen.append(token_u)
                else:
                    print(f"Unknown type: {token}")
                    break
        else:
            if chosen:
                return sorted(set(chosen))
            print("Please select at least one valid type.")


def _build_ranges(start: datetime, end: datetime, granularity: str) -> list[tuple[datetime, datetime]]:
    if granularity == "daily":
        ranges = []
        cur = start
        while cur < end:
            nxt = min(cur + timedelta(days=1), end)
            ranges.append((cur, nxt))
            cur = nxt
        return ranges
    if granularity == "weekly":
        ranges = []
        cur = start
        while cur < end:
            nxt = min(cur + timedelta(days=7), end)
            ranges.append((cur, nxt))
            cur = nxt
        return ranges
    return [(start, end)]


def _sanitize_filename(value: str) -> str:
    return "".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in value).strip().replace(" ", "_")


def main() -> int:
    args = parse_args()
    if not args.cert.exists():
        raise FileNotFoundError(f"Certificate not found: {args.cert}")
    if not args.key.exists():
        raise FileNotFoundError(f"Private key not found: {args.key}")

    print("BaSE+ Interactive Downloader")
    print("SOAP ops: list = MessageList, download = Any")
    msg_types = _prompt_msg_types()
    granularity = input("Granularity [daily/weekly/custom] (default custom): ").strip().lower() or "custom"
    start = _prompt_date("Start date (YYYY-MM-DD): ")
    end = _prompt_date("End date   (YYYY-MM-DD): ")
    if end <= start:
        raise ValueError("End date must be after start date.")

    client = BaseWSClient(
        cert_path=str(args.cert),
        key_path=str(args.key),
        endpoint_url=args.endpoint,
        verify_ssl=not args.insecure,
        timeout=args.timeout,
    )

    args.outdir.mkdir(parents=True, exist_ok=True)
    ranges = _build_ranges(start, end, granularity if granularity in {"daily", "weekly"} else "custom")

    for idx, (rng_start, rng_end) in enumerate(ranges, start=1):
        print(f"\nRange {idx}/{len(ranges)}: {rng_start.isoformat()} -> {rng_end.isoformat()}")
        response = client.list_message_results(
            start_time=rng_start.strftime("%Y-%m-%dT00:00:00"),
            end_time=rng_end.strftime("%Y-%m-%dT00:00:00"),
            msg_types=msg_types,
        )
        print(f"status={response.status_code}")
        if response.status_code != 200:
            print(response.text)
            continue
        try:
            messages = client.parse_message_list(response.text)
        except Exception as exc:
            print(f"Failed to parse message list: {exc}")
            print(response.text)
            continue

        print("Messages:")
        for m in messages:
            print(m)

        for m in messages:
            code = m.get("Code")
            if not code:
                continue
            target_name = m.get("MessageIdentification", f"{code}.bin")
            out_path = args.outdir / f"{_sanitize_filename(target_name)}_{code}.bin"
            print(f"Downloading code={code} -> {out_path.name}")
            dl = client.download_result(code)
            if dl.status_code != 200:
                print(f"download failed status={dl.status_code}")
                print(dl.text)
                continue
            blob = client.extract_compressed_payload(dl.text)
            data = client.decode_payload(blob)
            out_path.write_bytes(data)
            print(f"saved={out_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
