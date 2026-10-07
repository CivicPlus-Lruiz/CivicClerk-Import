import os
import re
import pandas as pd
from datetime import datetime
from urllib.parse import unquote
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

INPUT_XLSX  = "/mnt/user-data/uploads/Nassau.xlsx"
OUTPUT_XLSX = "/mnt/user-data/outputs/Meetings_Output.xlsx"

OUTPUT_COLUMNS = [
    "Event Name", "Event Date", "Event Time", "Event Category",
    "Meeting Type", "Event Description", "Video File Name",
    "Agenda File Name", "Minutes File Name", "Agenda Packet File Name",
    "Close Caption File Name", "Notice File Name", "Other File Name",
    "External Media URL", "Agenda is Published", "Minutes is Published",
    "Agenda Packet is Published"
]

# ================================================================
# SUBTYPE KEYWORDS  (ordered longest-first to catch multi-word first)
# ================================================================
SUBTYPE_KEYWORDS = [
    # Multi-word spaced variants (longest first)
    "winter strategic",
    "executive special",
    "executive closed",
    "reorganizational meeting",
    "reorganizational",
    "shade meeting",
    "joint meeting",
    "joint school board",
    "joint planning and zoning board",
    "joint with planning and zoning board",
    "joint with school board",
    "with school board",
    "with city of fernandina beach",
    "with the city of fernandina beach",
    "with the nassau county school board",
    "special meeting",
    "emergency meeting",
    "budget meeting",
    "budget meetings",
    "strategic planning",
    "study session",
    "work session",
    "workshop",
    "emergency",
    "special",
    "closed",
    "shade",
    # CamelCase / no-space variants (customers who name files without spaces)
    "specialmeeting",
    "emergencymeeting",
    "budgetmeeting",
    "strategicplanning",
    "studysession",
    "worksession",
    "jointmeeting",
    "shademeeting",
]

# Noise tokens to strip INTERNALLY for matching only
NOISE_PATTERNS = [
    r"\(copy\)",
    r"\bcopy\b",
    r"generated agenda",
    r"uploaded minutes",
    r"audio only\.?",
    r"audio only",
    r"video only",
    r"partial video",
    r"video \d+",
    r"part \d+",
    r"\bpart\d+\b",
    r"zoom video for.*",
    r"granicus video.*",
    r"granicus video from.*",
    r"\bjt\b",
    r"\d+%3a\d+\s*p\.?m\.?.*",   # URL-encoded timestamps
    r"from \d+.*",
    r"\.doc\b",
    r"\.txt\b",
    r"\.aspx\b",
    r"\(saissa\)",
    r"\bsaissa\b",
    r"adivsory",   # typo normalisation handled in alias
]

# ================================================================
# BODY ALIAS MAP  →  canonical name (without "Nassau County" prefix)
# ================================================================
BODY_ALIASES = {
    # Board of County Commissioners variants
    "nassau county board of county commissioners": "Board Of County Commissioners",
    "nassau county board of county commissioner":  "Board Of County Commissioners",
    "nassau county board of county commisioners":  "Board Of County Commissioners",
    "nassau county board of county commisioner":   "Board Of County Commissioners",
    "nassau county board of county commmissioners":"Board Of County Commissioners",
    "nassau county board of county commissoners":  "Board Of County Commissioners",
    "nassau county board of county commisssioners":"Board Of County Commissioners",
    "nassau board of county commissioners":        "Board Of County Commissioners",
    "nassau board of county commissioner":         "Board Of County Commissioners",
    "board of county commissioners":               "Board Of County Commissioners",
    "board of county commissioner":                "Board Of County Commissioners",
    "nassa county board of county commissioners":  "Board Of County Commissioners",
    "nassay county board of county commissioners": "Board Of County Commissioners",
    "bocc":                                        "Board Of County Commissioners",
    # Development Review Committee
    "nassau county development review committee":  "Development Review Committee",
    "development review committee":                "Development Review Committee",
    "development reveiw committee":                "Development Review Committee",
    "development review commitee":                 "Development Review Committee",
    "drc":                                         "Development Review Committee",
    # Planning and Zoning Board
    "nassau county planning and zoning board":     "Planning And Zoning Board",
    "nassau county planning and zoning":           "Planning And Zoning Board",
    # Conservation Land
    "conservation land acquisition and management committee":       "Conservation Land Acquisition And Management Committee",
    "conservation land and acquisition management committee":       "Conservation Land Acquisition And Management Committee",
    "conservation land aquisition management committee":            "Conservation Land Acquisition And Management Committee",
    "conservation land and aquisition management committee":        "Conservation Land Acquisition And Management Committee",
    "conservation lands acquisition and management committee":      "Conservation Land Acquisition And Management Committee",
    "conservation lands and aquisition management committee":       "Conservation Land Acquisition And Management Committee",
    "conservation land management and acquisition committee":       "Conservation Land Acquisition And Management Committee",
    # SAISSA
    "south amelia island shore stabilization association":          "South Amelia Island Shore Stabilization Association",
    "south amelia island shore stabilization association (saissa)": "South Amelia Island Shore Stabilization Association",
    "saissa":                                                       "South Amelia Island Shore Stabilization Association",
    # Code Enforcement
    "nassau county code enforcement board":        "Code Enforcement Board",
    "code enforcement board":                      "Code Enforcement Board",
    "nassau county code enforcement magistrate":   "Code Enforcement Magistrate",
    "code enforcement magistrate":                 "Code Enforcement Magistrate",
    # Conditional Use and Variance Board
    "nassau county conditional use and variance board": "Conditional Use And Variance Board",
    "conditional use and variance board":               "Conditional Use And Variance Board",
    # Amelia Island Local Planning Agency
    "amelia island local planning agency":         "Amelia Island Local Planning Agency",
    # Amelia Island Tourist Development Council
    "amelia island tourist development council":   "Amelia Island Tourist Development Council",
    # American Beach Water and Sewer District Advisory Board
    "american beach water and sewer district advisory board":  "American Beach Water And Sewer District Advisory Board",
    "american beach water and sewer district adivsory board":  "American Beach Water And Sewer District Advisory Board",
    # Affordable Housing Advisory Committee
    "affordable housing advisory committee":       "Affordable Housing Advisory Committee",
    # Essential Housing Advisory Committee
    "essential housing advisory committee":        "Essential Housing Advisory Committee",
    # Construction Board of Adjustments and Appeals
    "construction board of adjustments and appeals": "Construction Board Of Adjustments And Appeals",
    # Nassau County Amelia Island Tree Commission
    "nassau county amelia island tree commission": "Amelia Island Tree Commission",
    "amelia island tree commission":               "Amelia Island Tree Commission",
    # Nassau County Legislative Delegation
    "nassau county legislative delegation":        "Legislative Delegation",
    # Nassau County Opioid Settlement Task Force Advisory Committee
    "nassau county opioid settlement task force advisory committee": "Opioid Settlement Task Force Advisory Committee",
    "opioid settlement task force advisory committee":               "Opioid Settlement Task Force Advisory Committee",
    # Nassau County Planning Advisory Committee
    "nassau county planning advisory committee":   "Planning Advisory Committee",
    "nassau county plan advisory committee":       "Planning Advisory Committee",
    "planning advisory committee":                 "Planning Advisory Committee",
    # Nassau County Board of Commissioners (missing "County" in middle)
    "nassau county board of commissioners":        "Board Of County Commissioners",
    "nassau county board of commissioners meeting":"Board Of County Commissioners",
    "board of commissioners":                      "Board Of County Commissioners",
    # American Beach Water and Sewer District (advisory board typo variant)
    "american beach water and sewer district board":        "American Beach Water And Sewer District Advisory Board",
    "american beach water and sewer district advisory board": "American Beach Water And Sewer District Advisory Board",
    # Amelia Island Joint Local Planning Agency
    "amelia island joint local planning agency":   "Amelia Island Local Planning Agency",
    # Strategic Session / Audio Only variants → BOCC
    "strategic session":                           "Board Of County Commissioners",
    "strategic planning audio only":               "Board Of County Commissioners",
    "strategic audio only":                        "Board Of County Commissioners",
    # Fiscal Year Budget Meetings → BOCC
    "fiscal year budget meeting":                  "Board Of County Commissioners",
    "fiscal year budget meetings":                 "Board Of County Commissioners",
    # Winter Strategic → BOCC
    "winter strategic":                            "Board Of County Commissioners",
    "winter strategic day one":                    "Board Of County Commissioners",
    "winter strategic day two":                    "Board Of County Commissioners",
    "day one":                                     "Board Of County Commissioners",
    "day two":                                     "Board Of County Commissioners",
    # Code Enforcement Special Magistrate variant
    "code enforcement special magistrate":         "Code Enforcement Magistrate",
    "nassau county code enforcement special magistrate": "Code Enforcement Magistrate",
}

# Bodies that should go straight to error log (unrecognized)
UNRECOGNIZED_PATTERNS = [
    r"gran ilegislate",
    r"livemanager",
    r"nassau clerk live stream",
    r"web training",
    r"^\d{8}\s",          # filenames starting with raw date like 20250806
    r"^\d{8}$",
]

# ================================================================
# DATE PARSING
# ================================================================
MONTHS = {
    "january":1,"jan":1,"february":2,"feb":2,"march":3,"mar":3,
    "april":4,"apr":4,"may":5,"june":6,"jun":6,"july":7,"jul":7,
    "august":8,"aug":8,"september":9,"sept":9,"sep":9,
    "october":10,"oct":10,"november":11,"nov":11,"december":12,"dec":12,
}

def pivot_year(y):
    if y < 100:
        return 2000 + y if y < 50 else 1900 + y
    return y

WEEKDAYS = [
    "monday","tuesday","wednesday","thursday","friday","saturday","sunday",
    "mon","tues","tue","wed","thu","thur","thurs","fri","sat","sun"
]

# Explicit format strings supported by extract_date.
# Each maps to (year_pos, month_pos, day_pos) for a 3-segment date (0-indexed).
_EXPLICIT_FORMATS = {
    "MM-DD-YYYY": (2, 0, 1), "MM-DD-YY": (2, 0, 1),
    "MMDDYYYY":   (2, 0, 1), "MMDDYY":   (2, 0, 1),
    "DD-MM-YYYY": (2, 1, 0), "DD-MM-YY": (2, 1, 0),
    "DDMMYYYY":   (2, 1, 0), "DDMMYY":   (2, 1, 0),
    "YYYY-MM-DD": (0, 1, 2), "YYYYMMDD":  (0, 1, 2),
    "YY-MM-DD":   (0, 1, 2), "YYMMDD":    (0, 1, 2),
}

def extract_date(filename, date_format_pref="US DATE"):
    """
    Extract meeting date from filename.
    Returns (datetime, error_str) — error_str is "" on success.
    date_format_pref: "US DATE" (month-first), "EU DATE" (day-first),
                      or an explicit format key from _EXPLICIT_FORMATS.
    """
    explicit = _EXPLICIT_FORMATS.get(date_format_pref)
    if explicit is None and date_format_pref not in ("US DATE", "EU DATE"):
        date_format_pref = "US DATE"

    text = re.sub(r"\.(pdf|mp4|mp3|srt|vtt)$", "", filename, flags=re.IGNORECASE)
    text = text.lower()
    text = re.sub(r"[,\-_.]+", " ", text)        # normalise all separators incl dots
    text = re.sub(r"(\d)(st|nd|rd|th)\b", r"\1", text)
    parts = [p for p in text.split() if p not in WEEKDAYS]
    text  = " ".join(parts)

    valid = []

    # 1) Written month — month first (May 17 2023)
    for m in re.finditer(
        r"(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[^0-9a-z]*?(\d{1,2})(?:[^0-9a-z]*?(\d{2,4}))?",
        text):
        mon_t, d, y = m.groups()
        month = MONTHS.get(mon_t[:3], 0)
        day   = int(d)
        year  = int(y) if y else datetime.now().year
        if year < 100: year = pivot_year(year)
        if 1800 <= year <= datetime.now().year + 1:
            try: valid.append(datetime(year, month, day))
            except ValueError: pass

    # 1b) Written month — day first (27 Jan 2026)
    for m in re.finditer(
        r"(\d{1,2})[^0-9a-z]*?(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*(?:[^0-9a-z]*?(\d{2,4}))?",
        text):
        d, mon_t, y = m.groups()
        month = MONTHS.get(mon_t[:3], 0)
        day   = int(d)
        year  = int(y) if y else datetime.now().year
        if year < 100: year = pivot_year(year)
        if 1800 <= year <= datetime.now().year + 1:
            try: valid.append(datetime(year, month, day))
            except ValueError: pass

    # 2) Numeric triplets (MM DD YYYY or DD MM YYYY)
    for m in re.finditer(r"(?<!\d)(?:\d{3,4}[._\-\s])?(\d{1,4})[._\-\s](\d{1,4})[._\-\s](\d{2,4})(?!\d)", text):
        a, b, c = m.groups()
        try:
            segs = [int(a), int(b), int(c)]
            # Flag any 3-digit number in the triplet as an invalid year
            if any(100 <= int(x) < 1000 for x in (a, b, c)):
                bad = next(x for x in (a, b, c) if 100 <= int(x) < 1000)
                return None, f"Invalid 3-digit year in filename: {bad}"
            if explicit:
                # User specified an exact format — use it directly
                yi, mi, di = explicit
                yy, mm, dd = segs[yi], segs[mi], segs[di]
                if yy < 100: yy = pivot_year(yy)
                if yy > datetime.now().year + 1: continue
                try: valid.append(datetime(yy, mm, dd))
                except ValueError: pass
            else:
                mm, dd, yy = segs[0], segs[1], segs[2]
                if 100 <= yy < 1000:
                    return None, f"Invalid 3-digit year in filename: {yy}"
                if yy < 100: yy = pivot_year(yy)
                if yy > datetime.now().year + 1: continue
                d_us = d_eu = None
                try: d_us = datetime(yy, mm, dd)
                except ValueError: pass
                try: d_eu = datetime(yy, dd, mm)
                except ValueError: pass
                if date_format_pref == "EU DATE" and d_eu:   valid.append(d_eu)
                elif d_us:                                     valid.append(d_us)
                elif d_eu:                                     valid.append(d_eu)
        except Exception:
            continue

    # 3) Year-first (2023_05_17, 19740304)
    for m in re.finditer(r"((?:19|20)\d{2})[^\d]?(\d{1,2})[^\d]?(\d{1,2})(?!\d)", text):
        y, mo, da = map(int, m.groups())
        if 1 <= mo <= 12 and 1 <= da <= 31 and 1900 <= y <= datetime.now().year + 1:
            try: valid.append(datetime(y, mo, da))
            except ValueError: pass

    # 3.5) Pure 8-digit MMDDYYYY
    for m in re.finditer(r"\b(\d{8})\b", text):
        raw = m.group(1)
        mm = int(raw[:2]); dd = int(raw[2:4]); yyyy = int(raw[4:])
        if 1 <= mm <= 12 and 1 <= dd <= 31:
            try: valid.append(datetime(yyyy, mm, dd)); continue
            except ValueError: pass

    # 4) Compact 6-digit — honour explicit format, otherwise try YYMMDD first
    for m in re.finditer(r"\b(\d{6})\b", text):
        s = m.group(1)
        if explicit:
            # Explicit format: slice according to position map
            # For 6-digit compact the segments are always 2 digits each
            segs = [int(s[0:2]), int(s[2:4]), int(s[4:6])]
            yi, mi, di = explicit
            yy, mm, dd = segs[yi], segs[mi], segs[di]
            yy = pivot_year(yy)
            if 1 <= mm <= 12 and 1 <= dd <= 31 and yy <= datetime.now().year + 1:
                try: valid.append(datetime(yy, mm, dd))
                except ValueError: pass
            continue
        matched = False
        # Try YYMMDD (e.g. 220531 = 2022-05-31)
        try:
            y2, mo, da = int(s[:2]), int(s[2:4]), int(s[4:])
            y_full = pivot_year(y2)
            if 1 <= mo <= 12 and 1 <= da <= 31 and y_full <= datetime.now().year + 1:
                valid.append(datetime(y_full, mo, da))
                matched = True
        except ValueError:
            pass
        if matched:
            continue
        # Fallback: MMDDYY / DDMMYY
        try:
            if date_format_pref == "EU DATE":
                da, mo, y2 = int(s[:2]), int(s[2:4]), int(s[4:])
            else:
                mo, da, y2 = int(s[:2]), int(s[2:4]), int(s[4:])
            y_full = pivot_year(y2)
            if 1 <= mo <= 12 and 1 <= da <= 31 and y_full <= datetime.now().year + 1:
                valid.append(datetime(y_full, mo, da))
        except ValueError:
            continue

    if valid:
        return valid[-1], ""
    return None, ""

# ================================================================
# BODY + SUBTYPE PARSING
# ================================================================
def extract_filename_from_url(value: str):
    """
    If value is a URL (http/https/ftp), return (full_url, bare_filename).
    If it is already a bare filename, return (None, value).
    The full_url is stored in the output; bare_filename is used for all parsing.
    """
    stripped = value.strip()
    if re.match(r"^(https?|ftp)://", stripped, re.IGNORECASE):
        from urllib.parse import urlparse
        path = urlparse(stripped).path          # e.g. /Minutes/Some_File.pdf
        filename = unquote(path.split("/")[-1]) # e.g. Some_File.pdf
        return stripped, filename               # (full_url, filename_for_parsing)
    return None, stripped                       # (no url, bare filename)


def clean_for_matching(filename):
    """Strip extension, underscores, noise tokens — for matching only."""
    name = re.sub(r"\.(pdf|mp4|mp3|srt|vtt)$", "", filename, flags=re.IGNORECASE)
    name = unquote(name).replace("_", " ").replace("-", " ")
    # Remove URL-encoded parens
    name = name.replace("%28", "(").replace("%29", ")")
    name = re.sub(r"\s+", " ", name).strip()

    # Strip leading 8-digit date prefix (e.g. "20250806 SAISSA ...")
    name = re.sub(r"^\d{8}\s+", "", name)

    # Strip noise
    for pat in NOISE_PATTERNS:
        name = re.sub(pat, " ", name, flags=re.IGNORECASE)

    # Remove date patterns
    name = re.sub(r"\b20\d{2}\s+\d{1,2}\s+\d{1,2}\b", " ", name)
    name = re.sub(r"\b\d{1,2}\s+\d{1,2}\s+20\d{2}\b", " ", name)
    name = re.sub(r"\b\d{4}\b", " ", name)   # lone years
    name = re.sub(r"\s+", " ", name).strip()
    return name

def extract_subtype(cleaned_name):
    """Return (subtype_str, name_without_subtype) from cleaned name."""
    # Maps keyword (lowercase) → display name shown in output
    _DISPLAY = {
        "specialmeeting":    "Special Meeting",
        "emergencymeeting":  "Emergency Meeting",
        "budgetmeeting":     "Budget Meeting",
        "strategicplanning": "Strategic Planning",
        "studysession":      "Study Session",
        "worksession":       "Work Session",
        "jointmeeting":      "Joint Meeting",
        "shademeeting":      "Shade Meeting",
        "special meeting":   "Special Meeting",
        "emergency meeting": "Emergency Meeting",
        "budget meeting":    "Budget Meeting",
        "study session":     "Study Session",
        "work session":      "Work Session",
        "joint meeting":     "Joint Meeting",
        "shade meeting":     "Shade Meeting",
    }
    lower = cleaned_name.lower()
    for kw in SUBTYPE_KEYWORDS:
        pattern = r"\b" + re.escape(kw) + r"\b"
        if re.search(pattern, lower):
            remainder = re.sub(pattern, " ", cleaned_name, flags=re.IGNORECASE)
            remainder = re.sub(r"\s+", " ", remainder).strip()
            display = _DISPLAY.get(kw, kw.title())
            return display, remainder
    return "", cleaned_name

def resolve_body(name_without_subtype):
    """Map to canonical body name via alias dict."""
    key = name_without_subtype.strip().lower()
    # Direct lookup
    if key in BODY_ALIASES:
        return BODY_ALIASES[key]
    # Prefix-strip "nassau county " and retry
    stripped = re.sub(r"^nassau county\s+", "", key).strip()
    if stripped in BODY_ALIASES:
        return BODY_ALIASES[stripped]
    # Suffix-strip trailing numbers/labels e.g. "bocc 1", "bocc 2"
    stripped2 = re.sub(r"\s+\d+$", "", key).strip()
    if stripped2 in BODY_ALIASES:
        return BODY_ALIASES[stripped2]
    # Partial match — if key starts with a known alias
    for alias, canon in BODY_ALIASES.items():
        if key.startswith(alias) or alias.startswith(key):
            return canon
    return None

def is_unrecognized(cleaned_name):
    lower = cleaned_name.lower()
    for pat in UNRECOGNIZED_PATTERNS:
        if re.search(pat, lower):
            return True
    return False

def parse_filename(raw_value: str, date_format_pref: str = "US DATE"):
    """
    Accepts a raw cell value — either a bare filename or a full URL.
    date_format_pref: "US DATE" (month-first) or "EU DATE" (day-first).
    Returns dict: stored_value, date, subtype, body, match_key, is_error, error_reason
    """
    url, filename = extract_filename_from_url(raw_value)
    stored_value  = url if url else filename

    cleaned = clean_for_matching(filename)

    if is_unrecognized(cleaned):
        return {"stored_value": stored_value, "date": None, "subtype": "",
                "body": None, "match_key": None, "is_error": True,
                "error_reason": "Unrecognized meeting body"}

    date, date_err = extract_date(filename, date_format_pref)
    if date_err:
        return {"stored_value": stored_value, "date": None, "subtype": "",
                "body": None, "match_key": None, "is_error": True,
                "error_reason": date_err}
    if not date:
        return {"stored_value": stored_value, "date": None, "subtype": "",
                "body": None, "match_key": None, "is_error": True,
                "error_reason": "Could not parse date from filename"}

    subtype, name_no_sub = extract_subtype(cleaned)
    body = resolve_body(name_no_sub)

    if not body:
        return {"stored_value": stored_value, "date": date, "subtype": subtype,
                "body": None, "match_key": None, "is_error": True,
                "error_reason": f"Unrecognized meeting body: '{name_no_sub.strip()}'"}

    match_key = (body.lower(), date.strftime("%Y-%m-%d"), subtype.lower())
    return {"stored_value": stored_value, "date": date, "subtype": subtype,
            "body": body, "match_key": match_key, "is_error": False, "error_reason": ""}

# ================================================================
# EVENT NAME BUILDER
# ================================================================
NO_APPEND = ["meeting","session","hearing","workshop","retreat",
             "conference","townhall","board","commission","committee",
             "council","delegation","agency"]

def build_event_name(body, subtype):
    base = body.strip()
    if subtype:
        # avoid double "Meeting" from subtype
        st = subtype.strip()
        base = f"{base} {st}"
    lower_end = base.lower().split()[-1] if base else ""
    if lower_end not in NO_APPEND:
        base = f"{base} Meeting"
    return base

# ================================================================
# MAIN
# ================================================================
def process():
    df_in = pd.read_excel(INPUT_XLSX, dtype=str).fillna("")

    agendas  = [v for v in df_in["Agendas"].tolist() if v.strip()]
    minutes  = [v for v in df_in["Minutes"].tolist() if v.strip()]
    videos   = [v for v in df_in["Videos"].tolist()  if v.strip()]

    # Close Captions column — paired by row index with Videos.
    # The GUI pre-matched CC filenames to video filenames by date, so we just
    # read them positionally: cc_by_video[raw_video_value] = cc_filename.
    cc_col_raw = (df_in["Close Captions"].tolist()
                  if "Close Captions" in df_in.columns else [])
    videos_raw = df_in["Videos"].tolist()   # full list including blanks, for index alignment
    cc_by_video: dict = {}
    for raw_vid, raw_cc in zip(videos_raw, cc_col_raw):
        raw_vid = str(raw_vid).strip()
        raw_cc  = str(raw_cc).strip()
        if raw_vid and raw_cc:
            cc_by_video[raw_vid] = raw_cc

    # Buckets: match_key → {agenda, minutes, video, meta}
    meetings = {}   # match_key → record dict
    errors   = []   # list of error dicts
    # Track seen filenames per column for duplicate detection
    seen = {"Agendas": {}, "Minutes": {}, "Videos": {}, "Close Captions": {}}
    # Track which match_keys have a duplicate flagged (for highlighting)
    keys_with_duplicates = set()

    def process_column(filelist, col_name, file_field):
        for raw_value in filelist:
            raw_value = raw_value.strip()
            if not raw_value:
                continue

            # Exact value duplicate within same column
            if raw_value in seen[col_name]:
                errors.append({
                    "Filename": raw_value,
                    "Source Column": col_name,
                    "Issue": f"Duplicate — already seen in {col_name}"
                })
                continue
            seen[col_name][raw_value] = True

            parsed = parse_filename(raw_value)
            stored = parsed.get("stored_value", raw_value)  # full URL or bare filename

            if parsed["is_error"]:
                errors.append({
                    "Filename": stored,
                    "Source Column": col_name,
                    "Issue": parsed["error_reason"]
                })
                continue

            key = parsed["match_key"]

            if key not in meetings:
                meetings[key] = {
                    "Event Name":    build_event_name(parsed["body"], parsed["subtype"]),
                    "Event Date":    parsed["date"].strftime("%m/%d/%Y"),
                    "Event Time":    "",
                    "Event Category": parsed["body"],
                    "Meeting Type":  "",
                    "Event Description": "",
                    "Video File Name": "",
                    "Agenda File Name": "",
                    "Minutes File Name": "",
                    "Agenda Packet File Name": "",
                    "Close Caption File Name": "",
                    "Notice File Name": "",
                    "Other File Name": "",
                    "External Media URL": "",
                    "Agenda is Published": "",
                    "Minutes is Published": "",
                    "Agenda Packet is Published": "",
                    "_key": key,
                }

            rec = meetings[key]
            existing = rec[file_field]

            if existing:
                errors.append({
                    "Filename": stored,
                    "Source Column": col_name,
                    "Issue": f"Duplicate — another {col_name.rstrip('s')} already matched to this meeting ({existing})"
                })
                keys_with_duplicates.add(key)
            else:
                rec[file_field] = stored   # store URL or filename as-is

                # ── If this is a video, attach the pre-matched CC file ──
                if file_field == "Video File Name":
                    cc_fname = cc_by_video.get(raw_value, "")
                    if cc_fname and not rec["Close Caption File Name"]:
                        rec["Close Caption File Name"] = cc_fname

    process_column(agendas, "Agendas", "Agenda File Name")
    process_column(minutes, "Minutes", "Minutes File Name")
    process_column(videos,  "Videos",  "Video File Name")

    # Build output dataframe
    rows = [v for v in meetings.values()]
    df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS + ["_key"])
    # Sort by body then date
    df_out["_sort_date"] = pd.to_datetime(df_out["Event Date"], format="%m/%d/%Y", errors="coerce")
    df_out = df_out.sort_values(["Event Category", "_sort_date"])
    df_out = df_out.drop(columns=["_sort_date"])

    keys_list = df_out["_key"].tolist()
    df_out = df_out.drop(columns=["_key"])

    df_errors = pd.DataFrame(errors, columns=["Filename", "Source Column", "Issue"]) \
        if errors else pd.DataFrame(columns=["Filename", "Source Column", "Issue"])

    # Write with openpyxl for formatting
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="CivicClerk Import", index=False)
        df_errors.to_excel(writer, sheet_name="Error Log", index=False)

    # Apply formatting
    wb = load_workbook(OUTPUT_XLSX)

    YELLOW = PatternFill("solid", fgColor="FFFF00")
    HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
    HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    DATA_FONT   = Font(name="Arial", size=10)
    ERR_HEADER_FILL = PatternFill("solid", fgColor="7B2D00")

    # ── Main sheet ──
    ws = wb["CivicClerk Import"]
    for col_idx, cell in enumerate(ws[1], 1):
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 30

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        key = keys_list[row_idx - 2] if row_idx - 2 < len(keys_list) else None
        is_dup_row = key in keys_with_duplicates
        for cell in row:
            cell.font = DATA_FONT
            cell.alignment = Alignment(vertical="center")
            if is_dup_row:
                cell.fill = YELLOW

    # Auto-width columns
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

    ws.freeze_panes = "A2"

    # ── Error Log sheet ──
    we = wb["Error Log"]
    for cell in we[1]:
        cell.fill = ERR_HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    we.row_dimensions[1].height = 25

    for row in we.iter_rows(min_row=2):
        for cell in row:
            cell.font = DATA_FONT
            cell.alignment = Alignment(vertical="center")

    for col in we.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        we.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 80)

    we.freeze_panes = "A2"

    wb.save(OUTPUT_XLSX)

    print(f"Done. {len(df_out)} meeting rows, {len(df_errors)} errors.")
    print(f"Output: {OUTPUT_XLSX}")

if __name__ == "__main__":
    try:
        process()
    except Exception as e:
        import traceback
        traceback.print_exc()
