#!/usr/bin/env python3
"""
IntouchCX Philippines referral upload automation.

Reads eligible candidates from the Google Sheets tracker, maps fields to the
IntouchCX Google Form, submits sequentially, then writes "Executive Team" into
the Intouch column only after a confirmed successful submission.

Default mode is dry-run (no form posts, no sheet writes).

Usage:
  # Dry run (default): identify eligible candidates and validate mappings
  python intouch_uploader.py

  # Live submissions (requires Google service-account credentials)
  python intouch_uploader.py --live

  # Limit how many live submissions to attempt in one run
  python intouch_uploader.py --live --max-submit 5

Environment:
  GOOGLE_SERVICE_ACCOUNT_JSON     Preferred: full service-account JSON key contents
  GOOGLE_APPLICATION_CREDENTIALS  Alternative: path to a service-account JSON file
  INTOUCH_SPREADSHEET_ID          Optional override (default: tracker below)
  INTOUCH_SHEET_GID               Optional override (default: 112514731)
  INTOUCH_DRY_RUN                 Set to "0" / "false" to allow live mode via env
"""

from __future__ import annotations

import argparse
import csv
import fcntl
import json
import logging
import os
import re
import sys
import time
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

MANILA = ZoneInfo("Asia/Manila")
CUTOFF_DATE = date(2026, 10, 2)
INTOUCH_SUCCESS_MARK = "Executive Team"

DEFAULT_SPREADSHEET_ID = "1NoRX955F0dpxMReiC-6lcd3hgxFDccS3H9hzabTQghE"
DEFAULT_SHEET_GID = 112514731  # "VALID CONSENT - TP, FVR & EXL"

FORM_VIEW_URL = (
    "https://docs.google.com/forms/d/e/"
    "1FAIpQLScb7bYxPgxnJ_zLfIq8zRnA0sVCFW3TDEeM2SOpCYQzN-nG0Q/viewform"
)
FORM_RESPONSE_URL = (
    "https://docs.google.com/forms/d/e/"
    "1FAIpQLScb7bYxPgxnJ_zLfIq8zRnA0sVCFW3TDEeM2SOpCYQzN-nG0Q/formResponse"
)

# Verified against the live form structure (stop if these change).
FORM_ENTRIES = {
    "privacy_consent": "857581117",
    "referrer": "630903126",
    "email": "599356917",
    "full_name": "210162759",
    "contact": "1878718037",
    "preferred_site": "532848587",
    "bpo_experience": "801612430",
    "residence": "1880550055",
    "sales_experience": "1207245034",
    "chat_support": "182540642",
    "work_onsite": "169954946",
    "tsr_experience": "381462576",
    "seasonal": "2119671638",
    "campaign": "1834113923",
}

FIXED_ANSWERS = {
    "privacy_consent": "Yes",
    "referrer": "Olympuz - CXTRA006",
    "sales_experience": "No, I do not",
    "chat_support": "No, I do not have chat support experience",
    "work_onsite": "Yes, I can work on site",
    "tsr_experience": "No",
    "seasonal": "Yes",
    "campaign": "English Campaign",
}

PREFERRED_SITES = [
    "Cyberpark Tower One, Araneta City, Quezon City",
    "U.P. Town Center, Katipunan Ave., Quezon City",
    "One Montage Tower, Cebu City",
    "City Clark Front Tower 1,  Angeles City Pampanga",
    "Greenhills Tower,  Ortigas Ave. San Juan City",
]

RESIDENCE_OPTIONS = [
    "Within NCR",
    "Rizal/Antipolo",
    "Bulacan",
    "Cavite",
    "Laguna",
    "Within Metro Cebu (Cebu City, Compostela, Consolacion, Talisay, Mandaue, Lapu-Lapu, Liloan)",
    "Danao City",
    "Pampanga",
    "Tarlac",
]

BPO_OPTIONS = [
    "No Experience",
    "<6 Months",
    "6+ Mos",
    "12+ Mos",
    "18+ Mos",
    "24+ Mos",
    "36+ Mos",
    "48+ Mos",
    "60+ Mos",
]

# Form-listed Metro Cebu localities only (do not broaden).
METRO_CEBU_LOCALITIES = [
    "cebu city",
    "compostela",
    "consolacion",
    "talisay",
    "mandaue",
    "lapu-lapu",
    "lapu lapu",
    "liloan",
]

NCR_MARKERS = [
    "manila",
    "quezon city",
    "qc",
    "pasig",
    "makati",
    "pasay",
    "paranaque",
    "parañaque",
    "caloocan",
    "taguig",
    "las pinas",
    "las piñas",
    "mandaluyong",
    "san juan",
    "valenzuela",
    "malabon",
    "navotas",
    "marikina",
    "muntinlupa",
    "pateros",
    "cubao",
    "fairview",
    "novaliches",
    "monumento",
    "malibay",
    "ortigas",
    "bgc",
    "metro manila",
    "ncr",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

ROOT = Path(__file__).resolve().parent
LOG_DIR = ROOT / "logs"
STATE_DIR = ROOT / "data"
SUBMISSION_LOG = STATE_DIR / "submission_log.jsonl"
RUN_LOCK = STATE_DIR / "run.lock"

logger = logging.getLogger("intouch_uploader")


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------


@dataclass
class Candidate:
    sheet_row: int  # 1-based spreadsheet row number
    last_name: str
    first_name: str
    contact: str
    email: str
    location: str
    bpo_raw: str
    timestamp_raw: str
    record_date: date
    intouch: str
    status: str = ""
    preferred_site_raw: str = ""

    @property
    def applicant_name(self) -> str:
        return f"{self.last_name}, {self.first_name}"

    @property
    def candidate_id(self) -> str:
        """Stable id: sheet row + email + contact (not name alone)."""
        return f"row:{self.sheet_row}|email:{self.email.lower()}|contact:{self.contact}"


@dataclass
class MappedPayload:
    privacy_consent: str
    referrer: str
    email: str
    full_name: str
    contact: str
    preferred_site: str
    bpo_experience: str
    residence: str
    sales_experience: str
    chat_support: str
    work_onsite: str
    tsr_experience: str
    seasonal: str
    campaign: str


@dataclass
class AttemptResult:
    candidate_id: str
    sheet_row: int
    applicant_name: str
    eligibility: str
    submission: str = "not_attempted"
    sheet_update: str = "not_attempted"
    reason: str = ""
    mapped: Optional[dict[str, str]] = None


@dataclass
class RunSummary:
    run_at_manila: str
    mode: str
    total_records_scanned: int = 0
    total_eligible: int = 0
    total_submitted: int = 0
    total_intouch_updated: int = 0
    total_skipped: int = 0
    total_failed_or_unconfirmed: int = 0
    total_manual_review: int = 0
    attempts: list[AttemptResult] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def now_manila() -> datetime:
    return datetime.now(tz=MANILA)


def is_blank(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def col_letter(index_zero_based: int) -> str:
    """Convert 0-based column index to A1 column letters."""
    n = index_zero_based + 1
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def parse_record_date(raw: str) -> Optional[date]:
    raw = (raw or "").strip()
    if not raw:
        return None
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).date()
    except ValueError:
        pass
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None


def find_header(headers: list[str], *candidates: str) -> Optional[int]:
    normalized = {h.strip().lower(): i for i, h in enumerate(headers)}
    for name in candidates:
        idx = normalized.get(name.strip().lower())
        if idx is not None:
            return idx
    return None


# ---------------------------------------------------------------------------
# Field mapping
# ---------------------------------------------------------------------------


def map_bpo_experience(raw: str) -> tuple[Optional[str], Optional[str]]:
    """Map free-text BPO EXP to a form option. Skip if not confidently mappable."""
    s = norm(raw)
    if not s:
        return None, "missing BPO experience"

    no_exp = {
        "none",
        "non",
        "n/a",
        "na",
        "n.a",
        "n.a.",
        "0",
        "no",
        "no experience",
        "no bpo exp",
        "no bpo experience",
        "none yet",
        "-",
        "nil",
    }
    if s in no_exp or "no bpo" in s or s == "no experience":
        return "No Experience", None

    years = None
    months = None
    m_year = re.search(r"(\d+(?:\.\d+)?)\s*(?:years?|yrs?)\b", s)
    if m_year:
        years = float(m_year.group(1))
    m_month = re.search(r"(\d+)\s*(?:months?|mos?|mo)\b", s)
    if m_month:
        months = int(m_month.group(1))
    if months is None:
        m_glued = re.search(r"(\d+)months?\b", s)
        if m_glued:
            months = int(m_glued.group(1))

    if years is not None and months is not None:
        total_months = int(years * 12) + months
    elif years is not None:
        total_months = int(years * 12)
    elif months is not None:
        total_months = months
    else:
        m_over = re.search(r"over\s+(\d+)\s*(?:years?|yrs?)", s)
        if m_over:
            total_months = int(m_over.group(1)) * 12
        elif re.fullmatch(r"\d+", s) or "unit not specified" in s:
            return None, f"BPO experience unit ambiguous: {raw!r}"
        else:
            return None, f"BPO experience unmappable: {raw!r}"

    if total_months <= 0:
        return "No Experience", None
    if total_months < 6:
        return "<6 Months", None
    if total_months < 12:
        return "6+ Mos", None
    if total_months < 18:
        return "12+ Mos", None
    if total_months < 24:
        return "18+ Mos", None
    if total_months < 36:
        return "24+ Mos", None
    if total_months < 48:
        return "36+ Mos", None
    if total_months < 60:
        return "48+ Mos", None
    return "60+ Mos", None


def map_residence_and_site(
    location: str, preferred_site_raw: str = ""
) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Return (residence, preferred_site, error_reason).

    Only maps when the nearest site/residence can be determined reliably from
    the form's actual options. Ambiguous NCR multi-site cases are skipped.
    """
    # Explicit preferred-site column wins when it exactly matches a form option.
    explicit = (preferred_site_raw or "").strip()
    if explicit:
        for site in PREFERRED_SITES:
            if norm(explicit) == norm(site):
                # Still need residence from location.
                residence, _, err = map_residence_and_site(location, preferred_site_raw="")
                if err:
                    return None, None, err
                return residence, site, None
        return None, None, f"preferred site not in form options: {explicit!r}"

    loc = norm(location)
    if not loc:
        return None, None, "missing location / city of residence"

    # Danao
    if "danao" in loc:
        return (
            "Danao City",
            "One Montage Tower, Cebu City",
            None,
        )

    # Tarlac
    if "tarlac" in loc:
        return (
            "Tarlac",
            "City Clark Front Tower 1,  Angeles City Pampanga",
            None,
        )

    # Pampanga / Clark / Angeles
    if any(k in loc for k in ("pampanga", "angeles", "clark", "mabalacat")):
        return (
            "Pampanga",
            "City Clark Front Tower 1,  Angeles City Pampanga",
            None,
        )

    # Cavite
    if any(
        k in loc
        for k in (
            "cavite",
            "imus",
            "bacoor",
            "dasmariñas",
            "dasmarinas",
            "general trias",
            "gen. trias",
            "naic",
            "silang",
            "tagaytay",
            "tanza",
        )
    ):
        # Multiple NCR sites exist; nearest among them is not uniquely reliable.
        return None, None, f"Cavite location maps to residence but nearest NCR site is ambiguous: {location!r}"

    # Laguna
    if any(
        k in loc
        for k in (
            "laguna",
            "santa rosa",
            "sta. rosa",
            "sta rosa",
            "calamba",
            "biñan",
            "binan",
            "san pedro",
            "cabuyao",
        )
    ):
        return None, None, f"Laguna location maps to residence but nearest NCR site is ambiguous: {location!r}"

    # Bulacan
    if any(
        k in loc
        for k in (
            "bulacan",
            "meycauayan",
            "marilao",
            "malolos",
            "baliuag",
            "san jose del monte",
        )
    ):
        return None, None, f"Bulacan location maps to residence but nearest NCR site is ambiguous: {location!r}"

    # Rizal / Antipolo
    if any(
        k in loc
        for k in ("rizal", "antipolo", "cainta", "taytay", "binangonan", "angono", "rodriguez")
    ):
        return None, None, f"Rizal location maps to residence but nearest NCR site is ambiguous: {location!r}"

    # Metro Cebu — only form-listed localities
    metro_cebu_hit = any(k in loc for k in METRO_CEBU_LOCALITIES)
    if metro_cebu_hit:
        # Require Cebu context for Talisay (name exists outside Cebu).
        if "talisay" in loc and "cebu" not in loc:
            return None, None, f"Talisay without Cebu context; residence/site ambiguous: {location!r}"
        return (
            "Within Metro Cebu (Cebu City, Compostela, Consolacion, Talisay, Mandaue, Lapu-Lapu, Liloan)",
            "One Montage Tower, Cebu City",
            None,
        )

    # City of Naga, Cebu (listed under broader Cebu city area usage in tracker)
    if "naga" in loc and "cebu" in loc:
        return (
            "Within Metro Cebu (Cebu City, Compostela, Consolacion, Talisay, Mandaue, Lapu-Lapu, Liloan)",
            "One Montage Tower, Cebu City",
            None,
        )

    # High-confidence NCR site pins (unique nearest site)
    if any(k in loc for k in ("katipunan", "up town", "diliman", "loyola")):
        return "Within NCR", "U.P. Town Center, Katipunan Ave., Quezon City", None
    if any(k in loc for k in ("san juan", "greenhills")):
        return "Within NCR", "Greenhills Tower,  Ortigas Ave. San Juan City", None
    if any(k in loc for k in ("quezon city", "qc", "cubao", "araneta")):
        return "Within NCR", "Cyberpark Tower One, Araneta City, Quezon City", None
    if "pasig" in loc or "ortigas" in loc or "mandaluyong" in loc:
        return "Within NCR", "Greenhills Tower,  Ortigas Ave. San Juan City", None

    # Other NCR cities: multiple sites (Cyberpark / UP / Greenhills) — skip
    if any(k in loc for k in NCR_MARKERS):
        return None, None, f"NCR location has multiple possible preferred sites; nearest not unique: {location!r}"

    return None, None, f"location not mappable to form residence/site options: {location!r}"


def build_mapped_payload(candidate: Candidate) -> tuple[Optional[MappedPayload], Optional[str]]:
    if is_blank(candidate.last_name) or is_blank(candidate.first_name):
        return None, "missing first/last name"
    if is_blank(candidate.email) or "@" not in candidate.email:
        return None, "missing/invalid email"
    if is_blank(candidate.contact):
        return None, "missing contact number"

    bpo, bpo_err = map_bpo_experience(candidate.bpo_raw)
    if bpo_err:
        return None, bpo_err
    if bpo not in BPO_OPTIONS:
        return None, f"BPO option not in form: {bpo!r}"

    residence, site, loc_err = map_residence_and_site(
        candidate.location, candidate.preferred_site_raw
    )
    if loc_err:
        return None, loc_err
    if residence not in RESIDENCE_OPTIONS:
        return None, f"residence option not in form: {residence!r}"
    if site not in PREFERRED_SITES:
        return None, f"preferred site not in form: {site!r}"

    # Preserve contact as string (leading zeros).
    contact = str(candidate.contact).strip()
    # If Sheets dropped a leading 0 on an 10-digit PH mobile, keep as-is (do not invent).

    payload = MappedPayload(
        privacy_consent=FIXED_ANSWERS["privacy_consent"],
        referrer=FIXED_ANSWERS["referrer"],
        email=candidate.email.strip(),
        full_name=candidate.applicant_name,
        contact=contact,
        preferred_site=site,
        bpo_experience=bpo,
        residence=residence,
        sales_experience=FIXED_ANSWERS["sales_experience"],
        chat_support=FIXED_ANSWERS["chat_support"],
        work_onsite=FIXED_ANSWERS["work_onsite"],
        tsr_experience=FIXED_ANSWERS["tsr_experience"],
        seasonal=FIXED_ANSWERS["seasonal"],
        campaign=FIXED_ANSWERS["campaign"],
    )
    return payload, None


# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------


def _write_service_account_file(parsed: dict[str, Any]) -> Path:
    if parsed.get("type") != "service_account":
        raise RuntimeError(
            "Google credentials JSON must be a service-account key (type=service_account)"
        )
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    out = STATE_DIR / "google_service_account.json"
    out.write_text(json.dumps(parsed), encoding="utf-8")
    os.chmod(out, 0o600)
    return out


def resolve_credentials_path() -> Path:
    """
    Resolve Google credentials from either:
      - GOOGLE_SERVICE_ACCOUNT_JSON (raw JSON secret contents), or
      - GOOGLE_APPLICATION_CREDENTIALS (file path OR raw JSON contents)
    """
    raw_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_json:
        try:
            parsed = json.loads(raw_json)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "GOOGLE_SERVICE_ACCOUNT_JSON is set but is not valid JSON"
            ) from exc
        if not isinstance(parsed, dict):
            raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON must be a JSON object")
        return _write_service_account_file(parsed)

    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if not creds_path:
        raise RuntimeError(
            "Missing Google credentials. Set secret GOOGLE_SERVICE_ACCOUNT_JSON "
            "(service-account JSON contents) or GOOGLE_APPLICATION_CREDENTIALS "
            "(path to the JSON key file)."
        )

    # Support secrets that put raw JSON into GOOGLE_APPLICATION_CREDENTIALS.
    if creds_path.lstrip().startswith("{"):
        try:
            parsed = json.loads(creds_path)
        except json.JSONDecodeError as exc:
            raise RuntimeError(
                "GOOGLE_APPLICATION_CREDENTIALS looks like JSON but failed to parse"
            ) from exc
        if not isinstance(parsed, dict):
            raise RuntimeError("GOOGLE_APPLICATION_CREDENTIALS JSON must be an object")
        return _write_service_account_file(parsed)

    path = Path(creds_path)
    if path.exists():
        return path

    raise RuntimeError(
        "GOOGLE_APPLICATION_CREDENTIALS is set but is neither a valid file path "
        "nor service-account JSON contents."
    )


class SheetsClient:
    def __init__(self, spreadsheet_id: str, credentials_path: str | Path):
        creds = service_account.Credentials.from_service_account_file(
            str(credentials_path), scopes=SCOPES
        )
        self.service = build("sheets", "v4", credentials=creds, cache_discovery=False)
        self.spreadsheet_id = spreadsheet_id
        self._sheet_title: Optional[str] = None
        self._sheet_gid: Optional[int] = None

    def resolve_sheet(self, gid: int) -> str:
        meta = (
            self.service.spreadsheets()
            .get(spreadsheetId=self.spreadsheet_id, fields="sheets(properties(sheetId,title))")
            .execute()
        )
        for sheet in meta.get("sheets", []):
            props = sheet.get("properties", {})
            if props.get("sheetId") == gid:
                self._sheet_title = props["title"]
                self._sheet_gid = gid
                return self._sheet_title
        raise RuntimeError(f"No sheet tab found with gid={gid}")

    def read_all_values(self) -> list[list[str]]:
        if not self._sheet_title:
            raise RuntimeError("Call resolve_sheet() first")
        result = (
            self.service.spreadsheets()
            .values()
            .get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{self._sheet_title}'",
                majorDimension="ROWS",
                valueRenderOption="FORMATTED_VALUE",
            )
            .execute()
        )
        return result.get("values", [])

    def get_cell(self, a1: str) -> str:
        result = (
            self.service.spreadsheets()
            .values()
            .get(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{self._sheet_title}'!{a1}",
                valueRenderOption="FORMATTED_VALUE",
            )
            .execute()
        )
        values = result.get("values", [])
        if not values or not values[0]:
            return ""
        return str(values[0][0])

    def update_cell(self, a1: str, value: str) -> None:
        self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{self._sheet_title}'!{a1}",
            valueInputOption="RAW",
            body={"values": [[value]]},
        ).execute()


def load_candidates(values: list[list[str]]) -> tuple[list[Candidate], dict[str, int], list[str]]:
    if not values:
        raise RuntimeError("Spreadsheet is empty or inaccessible")

    headers = [str(h) for h in values[0]]
    idx = {
        "last_name": find_header(headers, "Last Name"),
        "first_name": find_header(headers, "First Name"),
        "contact": find_header(headers, "Contact Num", "Active Mobile Number"),
        "email": find_header(headers, "Email", "Email Address"),
        "location": find_header(headers, "City of Residence", "Location", "location"),
        "bpo": find_header(headers, "BPO EXP", "BPO or Callcenter Experience"),
        "timestamp": find_header(headers, "Timestamp", "Date", "Record Date"),
        "intouch": find_header(headers, "Intouch"),
        "status": find_header(headers, "Status"),
        "preferred_site": find_header(
            headers, "Preferred Site", "Site", "Intouch Site", "Site Assignment"
        ),
    }
    required = ["last_name", "first_name", "contact", "email", "location", "bpo", "timestamp", "intouch"]
    missing = [k for k in required if idx[k] is None]
    if missing:
        raise RuntimeError(
            f"Required headers not found: {missing}. Actual headers: {headers}"
        )

    candidates: list[Candidate] = []
    for row_offset, row in enumerate(values[1:], start=2):
        def cell(key: str) -> str:
            i = idx[key]
            if i is None or i >= len(row):
                return ""
            return str(row[i]) if row[i] is not None else ""

        # Contact preference: Contact Num column already preferred via header order.
        contact = cell("contact")
        ts_raw = cell("timestamp")
        record_date = parse_record_date(ts_raw)
        if record_date is None:
            continue

        candidates.append(
            Candidate(
                sheet_row=row_offset,
                last_name=cell("last_name").strip(),
                first_name=cell("first_name").strip(),
                contact=contact.strip(),
                email=cell("email").strip(),
                location=cell("location").strip(),
                bpo_raw=cell("bpo").strip(),
                timestamp_raw=ts_raw.strip(),
                record_date=record_date,
                intouch=cell("intouch").strip(),
                status=cell("status").strip(),
                preferred_site_raw=cell("preferred_site").strip(),
            )
        )
    return candidates, {k: v for k, v in idx.items() if v is not None}, headers


# ---------------------------------------------------------------------------
# Form submission
# ---------------------------------------------------------------------------


def verify_form_available(session: requests.Session) -> None:
    resp = session.get(FORM_VIEW_URL, timeout=45)
    resp.raise_for_status()
    if "formResponse" not in resp.text and "FB_PUBLIC_LOAD_DATA_" not in resp.text:
        raise RuntimeError("Form does not appear to be accepting responses / structure unrecognized")
    # Sanity-check known entry ids still present
    for key, entry_id in FORM_ENTRIES.items():
        if f"entry.{entry_id}" not in resp.text and entry_id not in resp.text:
            # Multi-page forms may only expose page-1 entries in HTML.
            if key != "privacy_consent":
                continue
            raise RuntimeError(
                f"Form structure changed: missing entry for {key} ({entry_id}). Stopping."
            )


def payload_to_form_data(payload: MappedPayload) -> dict[str, str]:
    data = {}
    mapping = asdict(payload)
    for field_name, value in mapping.items():
        entry_id = FORM_ENTRIES[field_name]
        data[f"entry.{entry_id}"] = value
    return data


def submit_form(session: requests.Session, payload: MappedPayload) -> tuple[bool, str]:
    """
    POST to Google Forms formResponse endpoint.
    Returns (success, detail).
    """
    form_data = payload_to_form_data(payload)
    headers = {
        "User-Agent": "IntouchUploader/1.0 (+authorized recruitment automation)",
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": FORM_VIEW_URL,
    }
    try:
        resp = session.post(
            FORM_RESPONSE_URL,
            data=form_data,
            headers=headers,
            timeout=60,
            allow_redirects=True,
        )
    except requests.RequestException as exc:
        return False, f"form request error: {exc}"

    body = resp.text or ""
    # Google typically returns 200 with confirmation text, or redirects.
    body_l = body.lower()
    has_confirmation_text = (
        "your response has been recorded" in body_l
        or "response has been recorded" in body_l
    )
    # Some Google Form posts return a sparse 200 after accept.
    sparse_ok = resp.status_code == 200 and len(body) < 50
    confirmed = resp.status_code in (200, 302) and (has_confirmation_text or sparse_ok)
    if confirmed:
        return True, f"http {resp.status_code}"
    if resp.status_code >= 400:
        return False, f"http {resp.status_code}"
    # Uncertain outcome — do NOT treat as success and do not auto-retry
    return False, f"unconfirmed submission (http {resp.status_code}, body_len={len(body)})"


# ---------------------------------------------------------------------------
# Submission log / locking
# ---------------------------------------------------------------------------


def load_successful_ids(path: Path) -> set[str]:
    success: set[str] = set()
    if not path.exists():
        return success
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if rec.get("submission") == "success":
                cid = rec.get("candidate_id")
                if cid:
                    success.add(cid)
    return success


def append_log(path: Path, record: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, ensure_ascii=False) + "\n")


class RunLock:
    def __init__(self, path: Path):
        self.path = path
        self._fh = None

    def __enter__(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._fh = self.path.open("w")
        try:
            fcntl.flock(self._fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as exc:
            raise RuntimeError("Another automation run is already in progress") from exc
        self._fh.write(str(os.getpid()))
        self._fh.flush()
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._fh:
            fcntl.flock(self._fh.fileno(), fcntl.LOCK_UN)
            self._fh.close()
            try:
                self.path.unlink(missing_ok=True)
            except OSError:
                pass


# ---------------------------------------------------------------------------
# Core run
# ---------------------------------------------------------------------------


def evaluate_eligibility(
    candidate: Candidate,
    prior_success: set[str],
) -> tuple[str, Optional[MappedPayload], str]:
    """
    Returns (eligibility, payload_or_none, reason).
    eligibility: eligible | skip | manual_review
    """
    if candidate.record_date < CUTOFF_DATE:
        return "skip", None, f"record date {candidate.record_date} before cutoff {CUTOFF_DATE}"
    if not is_blank(candidate.intouch):
        return "skip", None, "Intouch cell is not empty"
    if candidate.candidate_id in prior_success:
        return "skip", None, "already successfully submitted per automation log"

    payload, err = build_mapped_payload(candidate)
    if err:
        return "manual_review", None, err
    return "eligible", payload, ""


def run(dry_run: bool = True, max_submit: Optional[int] = None) -> RunSummary:
    load_dotenv()
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    STATE_DIR.mkdir(parents=True, exist_ok=True)

    run_ts = now_manila()
    summary = RunSummary(
        run_at_manila=run_ts.isoformat(),
        mode="dry_run" if dry_run else "live",
    )

    spreadsheet_id = os.getenv("INTOUCH_SPREADSHEET_ID", DEFAULT_SPREADSHEET_ID)
    sheet_gid = int(os.getenv("INTOUCH_SHEET_GID", str(DEFAULT_SHEET_GID)))
    creds_path = resolve_credentials_path()

    with RunLock(RUN_LOCK):
        sheets = SheetsClient(spreadsheet_id, creds_path)
        sheet_title = sheets.resolve_sheet(sheet_gid)
        logger.info("Resolved sheet gid=%s title=%r", sheet_gid, sheet_title)

        try:
            values = sheets.read_all_values()
        except HttpError as exc:
            raise RuntimeError(f"Spreadsheet inaccessible; aborting run: {exc}") from exc

        if not values:
            raise RuntimeError("Spreadsheet returned no data; aborting without submissions")

        candidates, header_idx, headers = load_candidates(values)
        summary.total_records_scanned = max(0, len(values) - 1)
        logger.info("Headers: %s", headers)
        logger.info("Header indices: %s", header_idx)

        intouch_col = header_idx["intouch"]
        prior_success = load_successful_ids(SUBMISSION_LOG)
        session = requests.Session()

        if not dry_run:
            verify_form_available(session)

        submitted_count = 0
        for candidate in candidates:
            eligibility, payload, reason = evaluate_eligibility(candidate, prior_success)

            attempt = AttemptResult(
                candidate_id=candidate.candidate_id,
                sheet_row=candidate.sheet_row,
                applicant_name=candidate.applicant_name,
                eligibility=eligibility,
                reason=reason,
                mapped=asdict(payload) if payload else None,
            )

            if eligibility == "skip":
                summary.total_skipped += 1
                # Only log skips that are Oct 2+ (reduce noise from pre-cutoff rows)
                if candidate.record_date >= CUTOFF_DATE:
                    _persist_attempt(summary, attempt, dry_run)
                continue

            if eligibility == "manual_review":
                summary.total_manual_review += 1
                summary.total_skipped += 1
                _persist_attempt(summary, attempt, dry_run)
                continue

            # eligible
            summary.total_eligible += 1

            if dry_run:
                attempt.submission = "dry_run"
                attempt.sheet_update = "dry_run"
                _persist_attempt(summary, attempt, dry_run)
                continue

            if max_submit is not None and submitted_count >= max_submit:
                attempt.eligibility = "skip"
                attempt.reason = "max-submit limit reached for this run"
                summary.total_skipped += 1
                _persist_attempt(summary, attempt, dry_run)
                continue

            # Pre-submit recheck of Intouch cell
            intouch_a1 = f"{col_letter(intouch_col)}{candidate.sheet_row}"
            current_intouch = sheets.get_cell(intouch_a1).strip()
            if current_intouch:
                attempt.eligibility = "skip"
                attempt.reason = f"Intouch filled before submit: {current_intouch!r}"
                summary.total_skipped += 1
                _persist_attempt(summary, attempt, dry_run)
                continue

            if candidate.candidate_id in load_successful_ids(SUBMISSION_LOG):
                attempt.eligibility = "skip"
                attempt.reason = "duplicate prevented by submission log recheck"
                summary.total_skipped += 1
                _persist_attempt(summary, attempt, dry_run)
                continue

            assert payload is not None
            ok, detail = submit_form(session, payload)
            if not ok:
                if "unconfirmed" in detail:
                    attempt.submission = "unconfirmed"
                    attempt.reason = detail
                    summary.total_failed_or_unconfirmed += 1
                else:
                    attempt.submission = "failed"
                    attempt.reason = detail
                    summary.total_failed_or_unconfirmed += 1
                _persist_attempt(summary, attempt, dry_run)
                # Do not immediately retry uncertain outcomes
                time.sleep(1)
                continue

            attempt.submission = "success"
            # Update Intouch only after confirmed success
            try:
                sheets.update_cell(intouch_a1, INTOUCH_SUCCESS_MARK)
                verified = sheets.get_cell(intouch_a1).strip()
                if verified != INTOUCH_SUCCESS_MARK:
                    attempt.sheet_update = "failed_verify"
                    attempt.reason = (
                        f"form submitted but Intouch verify failed "
                        f"(wanted {INTOUCH_SUCCESS_MARK!r}, got {verified!r}); reconcile before retry"
                    )
                    summary.total_failed_or_unconfirmed += 1
                    summary.total_submitted += 1
                    prior_success.add(candidate.candidate_id)
                    _persist_attempt(summary, attempt, dry_run)
                    continue
                attempt.sheet_update = "success"
                summary.total_submitted += 1
                summary.total_intouch_updated += 1
                prior_success.add(candidate.candidate_id)
                submitted_count += 1
            except HttpError as exc:
                attempt.sheet_update = "failed"
                attempt.reason = (
                    f"form submitted but sheet update failed: {exc}; "
                    "reconcile before retrying form submission"
                )
                summary.total_submitted += 1
                summary.total_failed_or_unconfirmed += 1
                prior_success.add(candidate.candidate_id)
            _persist_attempt(summary, attempt, dry_run)
            time.sleep(1.25)  # sequential pacing

    return summary


def _persist_attempt(summary: RunSummary, attempt: AttemptResult, dry_run: bool) -> None:
    summary.attempts.append(attempt)
    # Avoid logging full PII mapped payloads into durable logs; keep identifiers only.
    record = {
        "run_at_manila": summary.run_at_manila,
        "mode": summary.mode,
        "candidate_id": attempt.candidate_id,
        "sheet_row": attempt.sheet_row,
        "applicant_name": attempt.applicant_name,
        "eligibility": attempt.eligibility,
        "submission": attempt.submission,
        "sheet_update": attempt.sheet_update,
        "reason": attempt.reason,
    }
    if attempt.mapped and dry_run:
        # Dry-run report may include mapping for review (local file only).
        record["mapped"] = attempt.mapped
    append_log(SUBMISSION_LOG if not dry_run else LOG_DIR / "dry_run.jsonl", record)
    logger.info(
        "row=%s eligibility=%s submission=%s sheet_update=%s reason=%s",
        attempt.sheet_row,
        attempt.eligibility,
        attempt.submission,
        attempt.sheet_update,
        attempt.reason,
    )


def print_summary(summary: RunSummary) -> None:
    print("\n=== IntouchCX Upload Run Summary ===")
    print(f"Run (Asia/Manila): {summary.run_at_manila}")
    print(f"Mode:              {summary.mode}")
    print(f"Records scanned:   {summary.total_records_scanned}")
    print(f"Eligible:          {summary.total_eligible}")
    print(f"Submitted:         {summary.total_submitted}")
    print(f"Intouch updated:   {summary.total_intouch_updated}")
    print(f"Skipped:           {summary.total_skipped}")
    print(f"Failed/unconfirmed:{summary.total_failed_or_unconfirmed}")
    print(f"Manual review:     {summary.total_manual_review}")

    eligible = [a for a in summary.attempts if a.eligibility == "eligible"]
    manual = [a for a in summary.attempts if a.eligibility == "manual_review"]

    if eligible:
        print(f"\n--- Eligible candidates ({len(eligible)}) ---")
        for a in eligible:
            site = (a.mapped or {}).get("preferred_site", "")
            bpo = (a.mapped or {}).get("bpo_experience", "")
            residence = (a.mapped or {}).get("residence", "")
            print(
                f"  row {a.sheet_row}: {a.applicant_name} | "
                f"{residence} | {site} | BPO={bpo} | submission={a.submission}"
            )

    if manual:
        print(f"\n--- Manual review ({len(manual)}) ---")
        for a in manual[:50]:
            print(f"  row {a.sheet_row}: {a.applicant_name} | {a.reason}")
        if len(manual) > 50:
            print(f"  ... and {len(manual) - 50} more")

    out_path = LOG_DIR / f"summary_{now_manila().strftime('%Y%m%d_%H%M%S')}.json"
    # Compact summary JSON without repeating every mapped PII field in live mode
    dump = asdict(summary)
    if summary.mode == "live":
        for item in dump.get("attempts", []):
            item.pop("mapped", None)
    out_path.write_text(json.dumps(dump, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSummary written to {out_path}")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="IntouchCX candidate upload automation")
    p.add_argument(
        "--live",
        action="store_true",
        help="Enable live form submission and Intouch sheet updates (default: dry-run)",
    )
    p.add_argument(
        "--max-submit",
        type=int,
        default=None,
        help="Max live submissions in this run (ignored in dry-run)",
    )
    p.add_argument(
        "--verbose",
        action="store_true",
        help="Debug logging",
    )
    return p


def main(argv: Optional[list[str]] = None) -> int:
    load_dotenv()
    args = build_parser().parse_args(argv)

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(
                LOG_DIR / f"run_{now_manila().strftime('%Y%m%d')}.log",
                encoding="utf-8",
            ),
        ],
    )

    env_live = os.getenv("INTOUCH_DRY_RUN", "1").strip().lower() in {"0", "false", "no"}
    dry_run = not (args.live or env_live)

    if not dry_run:
        logger.warning("LIVE MODE enabled — form submissions and sheet updates will run")
    else:
        logger.info("Dry-run mode — no form submissions and no spreadsheet modifications")

    try:
        summary = run(dry_run=dry_run, max_submit=args.max_submit)
    except Exception as exc:
        logger.error("Run aborted: %s", exc)
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print_summary(summary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
