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
  GOOGLE_SIGNIN_EMAIL             Google account for form sign-in
                                  (default: hrcareers@olympuz-org.com)
  GOOGLE_SIGNIN_PASSWORD          Google password (secret; never commit)
  GOOGLE_STORAGE_STATE_JSON       Optional Playwright storage_state JSON from a
                                  trusted signed-in browser (bypasses 2FA prompts)
  INTOUCH_SPREADSHEET_ID          Optional override (default: tracker below)
  INTOUCH_SHEET_GID               Optional: process only this sheet gid
                                  (default: both VALID CONSENT + MYNIMO)
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
DEFAULT_GOOGLE_SIGNIN_EMAIL = "hrcareers@olympuz-org.com"

# Daily routine sources: VALID CONSENT + MYNIMO (each has an Intouch column).
SOURCE_SHEETS = [
    {"gid": 112514731, "key": "valid_consent"},  # VALID CONSENT - TP, FVR & EXL
    {"gid": 360587240, "key": "mynimo"},  # MYNIMO
]


def resolve_source_sheets() -> list[dict[str, Any]]:
    """Return sheets to process. INTOUCH_SHEET_GID limits to a single tab when set."""
    override = os.getenv("INTOUCH_SHEET_GID", "").strip()
    if override:
        gid = int(override)
        for src in SOURCE_SHEETS:
            if src["gid"] == gid:
                return [src]
        return [{"gid": gid, "key": f"gid_{gid}"}]
    return list(SOURCE_SHEETS)

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
    source_sheet: str = ""
    source_gid: int = 0

    @property
    def applicant_name(self) -> str:
        return f"{self.last_name}, {self.first_name}"

    @property
    def candidate_id(self) -> str:
        """Stable id: source sheet + row + email + contact (not name alone)."""
        sheet = self.source_sheet or "unknown"
        return (
            f"sheet:{sheet}|row:{self.sheet_row}|"
            f"email:{self.email.lower()}|contact:{self.contact}"
        )

    @property
    def legacy_candidate_id(self) -> str:
        """Pre-multi-source id format (VALID CONSENT-era submission log)."""
        return (
            f"row:{self.sheet_row}|"
            f"email:{self.email.lower()}|contact:{self.contact}"
        )


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
    source_sheet: str = ""


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


def repair_mojibake(text: str) -> str:
    """
    Fix common UTF-8→GBK mojibake seen in the tracker (e.g. ñ stored as 帽).

    Example: 'Ca帽oneo' → 'Cañoneo'
    """
    if not text:
        return text

    # Targeted: UTF-8 bytes of ñ (C3 B1) mis-decoded as GBK become 帽 (U+5E3D).
    targeted = (
        text.replace("\u5e3d", "ñ")
        .replace("Ã±", "ñ")
        .replace("Ã‘", "Ñ")
    )

    has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in targeted)
    has_latin = any(("A" <= ch <= "Z") or ("a" <= ch <= "z") for ch in targeted)
    if has_cjk and has_latin:
        try:
            repaired = text.encode("gbk").decode("utf-8")
            if repaired.isprintable():
                return repaired
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return targeted


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

    # MYNIMO-style ranges: "1 - 5 Months", "6 - 11 Months"
    m_month_range = re.search(r"(\d+)\s*-\s*(\d+)\s*(?:months?|mos?)\b", s)
    if m_month_range:
        total_months = int(m_month_range.group(1))
        years = None
        months = total_months
    else:
        years = None
        months = None
        # MYNIMO-style: "2 Years & Above", "3 Years & Above"
        m_years_above = re.search(r"(\d+)\s*(?:years?|yrs?)\s*&\s*above\b", s)
        if m_years_above:
            years = float(m_years_above.group(1))
        else:
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

    if years is not None and months is not None and not m_month_range:
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


def normalize_contact_number(raw: str) -> tuple[Optional[str], Optional[str]]:
    """
    Return a single phone-number string for the form.

    Prefers the first digit group when cells contain multiple numbers
    (e.g. '09636715467/ 09651168536'). Preserves leading zeros.
    """
    text = str(raw or "").strip()
    if not text:
        return None, "missing contact number"
    # Split on common separators for multiple numbers.
    parts = re.split(r"[/,;|]|\bor\b|\band\b", text, flags=re.I)
    for part in parts:
        digits = re.sub(r"\D", "", part)
        if len(digits) >= 10:
            # Keep original leading-zero style from the chosen segment when present.
            seg = part.strip()
            m = re.search(r"0?\d{10,15}", seg)
            if m:
                return m.group(0), None
            return digits, None
    # Fallback: first 10–15 digit run anywhere in the raw string.
    m = re.search(r"0?\d{10,15}", text)
    if m:
        return m.group(0), None
    return None, f"contact number not usable for form (must be a number): {raw!r}"


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

    contact, contact_err = normalize_contact_number(candidate.contact)
    if contact_err:
        return None, contact_err

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


def _infer_contact_column(
    headers: list[str],
    values: list[list[str]],
    idx: dict[str, Optional[int]],
) -> Optional[int]:
    """Infer contact column when the header label is missing or corrupted."""
    first_i = idx.get("first_name")
    loc_i = idx.get("location")
    email_i = idx.get("email")
    right_bound = None
    for bound in (loc_i, email_i):
        if bound is not None:
            right_bound = bound if right_bound is None else min(right_bound, bound)
    candidates: list[int] = []
    for i, _h in enumerate(headers):
        if first_i is not None and i <= first_i:
            continue
        if right_bound is not None and i >= right_bound:
            continue
        candidates.append(i)
    if not candidates:
        used = {v for v in idx.values() if v is not None}
        candidates = [i for i in range(len(headers)) if i not in used]

    best_i: Optional[int] = None
    best_hits = 0
    sample_rows = values[1:31]
    for i in candidates:
        hits = 0
        seen = 0
        for row in sample_rows:
            if i >= len(row):
                continue
            cell = str(row[i] or "").strip()
            if not cell:
                continue
            seen += 1
            digits = re.sub(r"\D", "", cell)
            if len(digits) >= 10:
                hits += 1
        if seen and hits >= max(3, (seen + 1) // 2) and hits > best_hits:
            best_hits = hits
            best_i = i
    if best_i is not None:
        logger.warning(
            "Contact header missing/corrupted; inferred column index %s header=%r",
            best_i,
            headers[best_i],
        )
    return best_i


def load_candidates(
    values: list[list[str]],
    *,
    source_sheet: str = "",
    source_gid: int = 0,
) -> tuple[list[Candidate], dict[str, int], list[str]]:
    if not values:
        raise RuntimeError("Spreadsheet is empty or inaccessible")

    headers = [str(h) for h in values[0]]
    idx = {
        "last_name": find_header(headers, "Last Name"),
        "first_name": find_header(headers, "First Name"),
        "contact": find_header(
            headers,
            "Contact Num",
            "Contact Number",
            "Contact",
            "Active Mobile Number",
            "Mobile Number",
            "Mobile",
            "Phone",
            "Phone Number",
        ),
        # Prefer exact "Email" over "Email Address" (MYNIMO has both-style headers).
        "email": find_header(headers, "Email", "Email Address"),
        "location": find_header(
            headers, "City of Residence", "City Of Residence", "Location", "location"
        ),
        "bpo": find_header(headers, "BPO EXP", "BPO or Callcenter Experience"),
        "timestamp": find_header(headers, "Timestamp", "Date", "Record Date"),
        "intouch": find_header(headers, "Intouch"),
        "status": find_header(headers, "Status"),
        "preferred_site": find_header(
            headers, "Preferred Site", "Site", "Intouch Site", "Site Assignment"
        ),
    }
    # VALID CONSENT contact header has been observed corrupted (e.g. "x`");
    # recover by content: column between First Name and Location/Email with phone-like values.
    if idx["contact"] is None:
        idx["contact"] = _infer_contact_column(headers, values, idx)

    required = ["last_name", "first_name", "contact", "email", "location", "bpo", "timestamp", "intouch"]
    missing = [k for k in required if idx[k] is None]
    if missing:
        raise RuntimeError(
            f"Required headers not found on {source_sheet or 'sheet'}: {missing}. "
            f"Actual headers: {headers}"
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
                # Repair tracker mojibake (ñ→帽) before form submit; do not rewrite the sheet name cells.
                last_name=repair_mojibake(cell("last_name").strip()),
                first_name=repair_mojibake(cell("first_name").strip()),
                contact=contact.strip(),
                email=cell("email").strip(),
                location=repair_mojibake(cell("location").strip()),
                bpo_raw=cell("bpo").strip(),
                timestamp_raw=ts_raw.strip(),
                record_date=record_date,
                intouch=cell("intouch").strip(),
                status=cell("status").strip(),
                preferred_site_raw=cell("preferred_site").strip(),
                source_sheet=source_sheet,
                source_gid=source_gid,
            )
        )
    return candidates, {k: v for k, v in idx.items() if v is not None}, headers


# ---------------------------------------------------------------------------
# Form submission (browser + Google sign-in)
# ---------------------------------------------------------------------------


def get_google_signin_credentials() -> tuple[str, str]:
    email = (
        os.getenv("GOOGLE_SIGNIN_EMAIL", "").strip()
        or DEFAULT_GOOGLE_SIGNIN_EMAIL
    )
    password = os.getenv("GOOGLE_SIGNIN_PASSWORD", "").strip()
    if not password:
        raise RuntimeError(
            "Missing secret GOOGLE_SIGNIN_PASSWORD. Add it in Cursor secrets "
            f"(account defaults to {DEFAULT_GOOGLE_SIGNIN_EMAIL}). "
            "Do not paste the password into chat."
        )
    return email, password


class BrowserFormSession:
    """
    Persistent Chromium session that signs into Google once, then submits
    IntouchCX form responses one candidate at a time.
    """

    def __init__(self) -> None:
        from playwright.sync_api import sync_playwright

        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        storage_state = self._load_storage_state()
        context_kwargs: dict[str, Any] = {
            "viewport": {"width": 1280, "height": 900},
            "user_agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            ),
            "locale": "en-US",
        }
        if storage_state is not None:
            context_kwargs["storage_state"] = storage_state
            logger.info("Loaded GOOGLE_STORAGE_STATE_JSON for trusted browser session")
        self._context = self._browser.new_context(**context_kwargs)
        self._page = self._context.new_page()
        self._signed_in = False
        self._used_storage_state = storage_state is not None

    @staticmethod
    def _load_storage_state() -> Optional[dict[str, Any]]:
        raw = os.getenv("GOOGLE_STORAGE_STATE_JSON", "").strip()
        if not raw:
            path = os.getenv("GOOGLE_STORAGE_STATE_PATH", "").strip()
            candidates = []
            if path:
                candidates.append(Path(path))
            candidates.append(STATE_DIR / "google_storage_state.json")
            for candidate in candidates:
                if candidate.exists():
                    raw = candidate.read_text(encoding="utf-8")
                    break
            else:
                return None
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise RuntimeError("GOOGLE_STORAGE_STATE_JSON is not valid JSON") from exc
        if not isinstance(data, dict):
            raise RuntimeError("GOOGLE_STORAGE_STATE_JSON must be a JSON object")
        return data

    def close(self) -> None:
        try:
            # Persist refreshed cookies for reuse within this environment disk.
            out = STATE_DIR / "google_storage_state.json"
            self._context.storage_state(path=str(out))
        except Exception:
            pass
        try:
            self._context.close()
        except Exception:
            pass
        try:
            self._browser.close()
        except Exception:
            pass
        try:
            self._playwright.stop()
        except Exception:
            pass

    def _form_accessible(self) -> bool:
        body = self._page.inner_text("body").lower()
        return (
            "do you agree with intouchcx" in body
            or "data privacy notice" in body
            or "preferred site" in body
        ) and "couldn’t sign you in" not in body and "couldn't sign you in" not in body

    def ensure_signed_in(self) -> None:
        if self._signed_in:
            return

        page = self._page
        # If a trusted storage state was provided, try the form directly first.
        if self._used_storage_state:
            page.goto(FORM_VIEW_URL, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(2000)
            if self._form_accessible():
                self._signed_in = True
                logger.info("Form accessible via saved Google session")
                return
            logger.warning("Saved session did not unlock the form; falling back to password sign-in")

        email, password = get_google_signin_credentials()
        logger.info("Signing into Google as %s (password from secret)", email)
        page.goto(
            "https://accounts.google.com/v3/signin/identifier?flowName=GlifWebSignIn&flowEntry=ServiceLogin&hl=en",
            wait_until="domcontentloaded",
            timeout=60000,
        )
        page.wait_for_timeout(1500)

        email_box = page.locator("#identifierId, input[type='email']").first
        email_box.wait_for(state="visible", timeout=30000)
        email_box.fill(email)
        page.locator("#identifierNext").first.click()
        page.wait_for_timeout(3000)

        body = page.inner_text("body").lower()
        if any(
            p in body
            for p in (
                "couldn't find your google account",
                "couldn’t find your google account",
                "enter a valid email",
            )
        ):
            raise RuntimeError(
                f"Google sign-in rejected the email {email!r}. "
                "Check GOOGLE_SIGNIN_EMAIL / account access."
            )

        pwd_box = page.locator('input[type="password"]').first
        try:
            pwd_box.wait_for(state="visible", timeout=30000)
        except Exception as exc:
            snippet = page.inner_text("body")[:400].replace("\n", " ")
            raise RuntimeError(
                "Google did not show the password field "
                f"(possible CAPTCHA / block). Page text: {snippet!r}"
            ) from exc
        pwd_box.fill(password)
        page.locator("#passwordNext").first.click()
        page.wait_for_timeout(4000)

        # Handle 2-Step Verification: prefer phone tap and wait for approval.
        for _ in range(24):  # up to ~2 minutes
            body = page.inner_text("body").lower()
            url = page.url
            if "myaccount.google.com" in url or ("accounts.google.com" not in url and "signin" not in url):
                break
            if "couldn’t sign you in" in body or "couldn't sign you in" in body:
                raise RuntimeError(
                    "Google blocked sign-in from this unrecognized device/location. "
                    "Approve a phone prompt from a prior attempt, sign in once from a "
                    "trusted browser and provide GOOGLE_STORAGE_STATE_JSON, or use an "
                    "account that allows this environment."
                )
            if "wrong password" in body or "incorrect password" in body:
                raise RuntimeError("Google rejected the password in GOOGLE_SIGNIN_PASSWORD.")
            if "2-step verification" in body or "challenge" in url:
                # Choose phone prompt if listed
                phone = page.get_by_text(re.compile("tap yes on your phone", re.I))
                if phone.count():
                    try:
                        phone.first.click(timeout=3000)
                        logger.info(
                            "Waiting for Google 2FA phone approval "
                            "(tap Yes on the phone for %s)...",
                            email,
                        )
                    except Exception:
                        pass
                page.wait_for_timeout(5000)
                continue
            # Signed in enough to leave challenge pages
            if "signin" not in url and "challenge" not in url:
                break
            page.wait_for_timeout(5000)

        # Open the form while authenticated
        page.goto(FORM_VIEW_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(2500)
        if "sign in to continue" in page.inner_text("body").lower():
            try:
                page.get_by_role("button", name=re.compile("sign in", re.I)).first.click(
                    timeout=5000
                )
                page.wait_for_timeout(4000)
            except Exception:
                pass

        if not self._form_accessible():
            snippet = page.inner_text("body")[:300].replace("\n", " ")
            raise RuntimeError(
                "Form still not accessible after Google sign-in. "
                f"Complete phone 2FA approval or provide GOOGLE_STORAGE_STATE_JSON. "
                f"Page preview: {snippet!r}"
            )
        self._signed_in = True
        logger.info("Google sign-in complete; form is accessible")

    def _click_option(self, text: str, exact: bool = True) -> None:
        page = self._page
        # Prefer radio/option by aria-label or visible text
        radio = page.locator(f'[role="radio"][aria-label="{text}"]')
        if radio.count():
            radio.first.click()
            return
        if exact:
            page.get_by_text(text, exact=True).first.click()
        else:
            page.get_by_text(text, exact=False).first.click()

    def _fill_question_input(self, question_substring: str, value: str) -> None:
        item = self._page.locator("div[role=listitem]").filter(
            has_text=re.compile(re.escape(question_substring), re.I)
        )
        box = item.locator("input[type=text], input[type=email], textarea").first
        box.fill(value)

    def _check_collect_email(self) -> None:
        """Page 1 requires recording the signed-in Google email with the response."""
        page = self._page
        checkbox = page.locator('[role="checkbox"][aria-label*="Record" i]')
        if not checkbox.count():
            checkbox = page.get_by_role(
                "checkbox", name=re.compile("record .* email", re.I)
            )
        if checkbox.count():
            checked = checkbox.first.get_attribute("aria-checked")
            if checked != "true":
                checkbox.first.click()
                page.wait_for_timeout(300)

    def _click_radio_value(self, value: str) -> None:
        page = self._page
        radio = page.locator(f'[role="radio"][aria-label="{value}"]')
        if radio.count():
            radio.first.click()
            return
        radio = page.locator(f'[role="radio"][data-value="{value}"]')
        if radio.count():
            radio.first.click()
            return
        page.get_by_text(value, exact=True).first.click()

    def submit(self, payload: MappedPayload) -> tuple[bool, str]:
        """Fill and submit the multi-page IntouchCX form for one candidate."""
        try:
            self.ensure_signed_in()
        except Exception as exc:
            return False, f"google sign-in failed: {exc}"

        page = self._page
        try:
            page.goto(FORM_VIEW_URL, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(1500)

            # Page 1: collect signed-in email + privacy consent
            self._check_collect_email()
            self._click_radio_value(payload.privacy_consent)
            page.wait_for_timeout(300)
            next_btn = page.get_by_role("button", name=re.compile("^next$", re.I))
            if next_btn.count():
                next_btn.first.click()
            else:
                page.locator("div[role=button]", has_text=re.compile("^next$", re.I)).first.click()
            page.wait_for_timeout(2500)

            # Confirm we left page 1
            if page.locator('[role="radio"][aria-label="Olympuz - CXTRA006"]').count() == 0:
                # Retry Next once if still on page 1 due to validation
                self._check_collect_email()
                if next_btn.count():
                    next_btn.first.click()
                page.wait_for_timeout(2500)
            if page.locator('[role="radio"][aria-label="Olympuz - CXTRA006"]').count() == 0:
                shot = LOG_DIR / f"page1_stuck_{now_manila().strftime('%Y%m%d_%H%M%S')}.png"
                page.screenshot(path=str(shot), full_page=True)
                return False, f"could not advance past form page 1; screenshot={shot}"

            # Page 2: applicant fields
            self._click_radio_value(payload.referrer)
            self._fill_question_input("Email address (Applicant)", payload.email)
            self._fill_question_input("Full Name (Applicant)", payload.full_name)
            self._fill_question_input("Contact Number (Applicant)", payload.contact)
            self._click_radio_value(payload.preferred_site)
            self._click_radio_value(payload.bpo_experience)
            # Residence option text is long; match via aria-label/data-value first.
            self._click_radio_value(payload.residence)

            sales_q = page.locator("div[role=listitem]").filter(has_text="sales experience")
            sales_q.locator(f'[role=radio][aria-label="{payload.sales_experience}"]').first.click()

            chat_q = page.locator("div[role=listitem]").filter(has_text="chat support")
            chat_q.locator(f'[role=radio][aria-label="{payload.chat_support}"]').first.click()

            onsite_q = page.locator("div[role=listitem]").filter(has_text="work on site")
            onsite_q.locator(f'[role=radio][aria-label="{payload.work_onsite}"]').first.click()

            tsr_q = page.locator("div[role=listitem]").filter(has_text="TSR")
            tsr_q.locator(f'[role=radio][aria-label="{payload.tsr_experience}"]').first.click()

            seasonal_q = page.locator("div[role=listitem]").filter(has_text="seasonal")
            seasonal_q.locator(f'[role=radio][aria-label="{payload.seasonal}"]').first.click()

            campaign_q = page.locator("div[role=listitem]").filter(has_text="Campaign")
            campaign_q.locator(f'[role=radio][aria-label="{payload.campaign}"]').first.click()

            page.wait_for_timeout(500)
            submit = page.get_by_role("button", name=re.compile("^submit$", re.I))
            if submit.count():
                submit.first.click()
            else:
                page.locator("div[role=button]", has_text=re.compile("^submit$", re.I)).first.click()
            page.wait_for_timeout(4000)

            body = page.inner_text("body")
            body_l = body.lower()
            if "your response has been recorded" in body_l or "response has been recorded" in body_l:
                return True, "browser confirmation recorded"
            # Uncertain — do not treat as success
            shot = LOG_DIR / f"unconfirmed_{now_manila().strftime('%Y%m%d_%H%M%S')}.png"
            try:
                page.screenshot(path=str(shot), full_page=True)
            except Exception:
                shot = None
            return False, (
                f"unconfirmed browser submission; url={page.url}; "
                f"screenshot={shot}; body_preview={body[:240]!r}"
            )
        except Exception as exc:
            return False, f"browser form error: {exc}"


def verify_form_reachable() -> None:
    """Lightweight check that the public form URL loads."""
    resp = requests.get(FORM_VIEW_URL, timeout=45)
    resp.raise_for_status()
    if "FB_PUBLIC_LOAD_DATA_" not in resp.text and "IntouchCX" not in resp.text:
        raise RuntimeError("Form does not appear available / structure unrecognized")


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
    if candidate.candidate_id in prior_success or (
        candidate.legacy_candidate_id in prior_success
    ):
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
    sources = resolve_source_sheets()
    creds_path = resolve_credentials_path()

    with RunLock(RUN_LOCK):
        sheets = SheetsClient(spreadsheet_id, creds_path)
        prior_success = load_successful_ids(SUBMISSION_LOG)
        browser_session: Optional[BrowserFormSession] = None

        if not dry_run:
            verify_form_reachable()
            # Fail fast if password secret is missing before any candidate loop.
            get_google_signin_credentials()
            browser_session = BrowserFormSession()

        attempt_count = 0
        try:
            for source in sources:
                sheet_gid = int(source["gid"])
                source_key = str(source["key"])
                sheet_title = sheets.resolve_sheet(sheet_gid)
                logger.info(
                    "Resolved sheet key=%s gid=%s title=%r",
                    source_key,
                    sheet_gid,
                    sheet_title,
                )

                try:
                    values = sheets.read_all_values()
                except HttpError as exc:
                    raise RuntimeError(
                        f"Spreadsheet inaccessible for {source_key}; aborting run: {exc}"
                    ) from exc

                if not values:
                    logger.warning("Sheet %s returned no data; skipping tab", source_key)
                    continue

                candidates, header_idx, headers = load_candidates(
                    values,
                    source_sheet=source_key,
                    source_gid=sheet_gid,
                )
                summary.total_records_scanned += max(0, len(values) - 1)
                logger.info("[%s] Headers: %s", source_key, headers)
                logger.info("[%s] Header indices: %s", source_key, header_idx)

                intouch_col = header_idx["intouch"]

                for candidate in candidates:
                    eligibility, payload, reason = evaluate_eligibility(
                        candidate, prior_success
                    )

                    attempt = AttemptResult(
                        candidate_id=candidate.candidate_id,
                        sheet_row=candidate.sheet_row,
                        applicant_name=candidate.applicant_name,
                        eligibility=eligibility,
                        reason=reason,
                        mapped=asdict(payload) if payload else None,
                        source_sheet=source_key,
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

                    if max_submit is not None and attempt_count >= max_submit:
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

                    success_ids = load_successful_ids(SUBMISSION_LOG)
                    if (
                        candidate.candidate_id in success_ids
                        or candidate.legacy_candidate_id in success_ids
                    ):
                        attempt.eligibility = "skip"
                        attempt.reason = "duplicate prevented by submission log recheck"
                        summary.total_skipped += 1
                        _persist_attempt(summary, attempt, dry_run)
                        continue

                    assert payload is not None
                    assert browser_session is not None
                    attempt_count += 1
                    ok, detail = browser_session.submit(payload)
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
                                f"(wanted {INTOUCH_SUCCESS_MARK!r}, got {verified!r}); "
                                "reconcile before retry"
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
        finally:
            if browser_session is not None:
                browser_session.close()

    return summary


def _persist_attempt(summary: RunSummary, attempt: AttemptResult, dry_run: bool) -> None:
    summary.attempts.append(attempt)
    # Avoid logging full PII mapped payloads into durable logs; keep identifiers only.
    record = {
        "run_at_manila": summary.run_at_manila,
        "mode": summary.mode,
        "candidate_id": attempt.candidate_id,
        "source_sheet": attempt.source_sheet,
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
        "sheet=%s row=%s eligibility=%s submission=%s sheet_update=%s reason=%s",
        attempt.source_sheet or "-",
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
                f"  [{a.source_sheet or '-'}] row {a.sheet_row}: {a.applicant_name} | "
                f"{residence} | {site} | BPO={bpo} | submission={a.submission}"
            )

    if manual:
        print(f"\n--- Manual review ({len(manual)}) ---")
        for a in manual[:50]:
            print(
                f"  [{a.source_sheet or '-'}] row {a.sheet_row}: "
                f"{a.applicant_name} | {a.reason}"
            )
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
