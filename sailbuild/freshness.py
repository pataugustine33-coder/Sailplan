"""
Freshness gate — single authority on whether a fetched source is operational.

Used by weather_pull.py to decide:
  - whether to accept a fetch result (parses out the issuance time and tiers it)
  - whether to fall through to the next URL in the fallback chain
  - whether build.py should run (`--require` blocks if any required source is STALE)

Tiers (cadence-aware):
  FRESH    age < refresh_hr × 2     (within 1-2 cycles)
  STALE    age < refresh_hr × 12    (multiple cycles late, may still be useful)
  ARCHIVED age >= refresh_hr × 12   (don't use as live forecast)
  UNKNOWN  could not parse issuance timestamp

Note: this duplicates some logic in sailbuild/tabs/support.py (classify_freshness),
but the support.py version is workbook-render-time on already-assembled data.
This module operates earlier in the pipeline — on raw fetched text — so it can
gate the assembly itself.
"""
from datetime import datetime, timezone, timedelta
import re


_MONTH = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
          "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}

_TZ_OFFSET = {"EDT": -4, "EST": -5, "CDT": -5, "CST": -6, "MDT": -6, "MST": -7,
              "PDT": -7, "PST": -8, "AKDT": -8, "AKST": -9, "HDT": -9, "HST": -10}

# Regex for NWS bulletin headers — "1118 AM EDT Tue May 12 2026" or "530 PM EDT Tue May 12 2026"
_NWS_TIME_RE = re.compile(
    r"(\d{3,4})\s+(AM|PM)\s+(EDT|EST|CDT|CST|MDT|MST|PDT|PST|AKDT|AKST|HDT|HST)"
    r"\s+\w{3}\s+(\w{3})\s+(\d{1,2})\s+(\d{4})",
    re.IGNORECASE,
)

# Buoy reading time — "as of (11:40 am EDT)" or "11:08 am EDT"
_BUOY_TIME_RE = re.compile(
    r"(?:as\s+of\s*\(?\s*)?(\d{1,2}):(\d{2})\s*(am|pm|AM|PM)\s+(\w{3,4})",
    re.IGNORECASE,
)
_BUOY_UTC_RE = re.compile(
    r"(\d{2})(\d{2})\s+GMT(?:\s+on)?\s+(\d{2})/(\d{2})/(\d{2,4})",
    re.IGNORECASE,
)


def extract_issuance(text: str, kind: str) -> datetime | None:
    """Parse issuance time from a fetched source's text.

    Returns aware UTC datetime, or None if not parseable.
    """
    if not text:
        return None

    if kind in ("cwf", "afd", "offshore", "high_seas", "nhc"):
        return _extract_nws_bulletin_time(text)
    elif kind == "buoy":
        return _extract_buoy_time(text)
    return None


def _extract_nws_bulletin_time(text: str) -> datetime | None:
    """NWS bulletins put issuance in the first ~10 lines as e.g.
    '1118 AM EDT Tue May 12 2026' or '530 PM EDT Tue May 12 2026'."""
    head = "\n".join(text.splitlines()[:15])
    m = _NWS_TIME_RE.search(head)
    if not m:
        return None
    try:
        time_str, ampm, tz_abbr = m.group(1), m.group(2).upper(), m.group(3).upper()
        month_name, day_str, year_str = m.group(4), m.group(5), m.group(6)
        minute = int(time_str[-2:])
        hour = int(time_str[:-2])
        if ampm == "PM" and hour != 12:
            hour += 12
        elif ampm == "AM" and hour == 12:
            hour = 0
        month = _MONTH.get(month_name.capitalize())
        if month is None:
            return None
        offset_h = _TZ_OFFSET.get(tz_abbr, 0)
        local = datetime(int(year_str), month, int(day_str), hour, minute,
                         tzinfo=timezone(timedelta(hours=offset_h)))
        return local.astimezone(timezone.utc)
    except (ValueError, KeyError, IndexError):
        return None


def _extract_buoy_time(text: str) -> datetime | None:
    """Buoy pages can give us either GMT (preferred, unambiguous) or local time."""
    # Prefer GMT — "1540 GMT on 04/15/2026"
    m = _BUOY_UTC_RE.search(text)
    if m:
        try:
            hh, mm = int(m.group(1)), int(m.group(2))
            mo, dd = int(m.group(3)), int(m.group(4))
            yy_raw = m.group(5)
            yy = int(yy_raw) if len(yy_raw) == 4 else 2000 + int(yy_raw)
            return datetime(yy, mo, dd, hh, mm, tzinfo=timezone.utc)
        except (ValueError, IndexError):
            pass
    # Fall back to local time — less precise about date, assume today UTC
    m = _BUOY_TIME_RE.search(text)
    if m:
        try:
            hh, mm, ampm, tz_abbr = int(m.group(1)), int(m.group(2)), m.group(3).upper(), m.group(4).upper()
            if ampm == "PM" and hh != 12:
                hh += 12
            elif ampm == "AM" and hh == 12:
                hh = 0
            offset_h = _TZ_OFFSET.get(tz_abbr, 0)
            now = datetime.now(timezone.utc)
            local_now = now.astimezone(timezone(timedelta(hours=offset_h)))
            local = local_now.replace(hour=hh, minute=mm, second=0, microsecond=0)
            # If derived local time is more than 12 hr in the future, assume yesterday
            if local > local_now + timedelta(hours=12):
                local -= timedelta(days=1)
            return local.astimezone(timezone.utc)
        except (ValueError, KeyError, IndexError):
            pass
    return None


def tier_freshness(issued_utc: datetime | None, refresh_hr: float,
                   now_utc: datetime | None = None) -> tuple[str, str]:
    """Return (tier, age_str) for a given issuance time.

    Tier is one of: FRESH, STALE, ARCHIVED, UNKNOWN.
    age_str is human-readable: "8.1 hr" / "2.2 days" / "11.3 weeks".
    """
    if issued_utc is None:
        return ("UNKNOWN", "?")
    if now_utc is None:
        now_utc = datetime.now(timezone.utc)
    age_hr = max(0.0, (now_utc - issued_utc).total_seconds() / 3600.0)
    # Format age
    if age_hr < 24:
        age_str = f"{age_hr:.1f} hr"
    elif age_hr < 14 * 24:
        age_str = f"{age_hr / 24:.1f} days"
    else:
        age_str = f"{age_hr / (24 * 7):.1f} weeks"
    # Tier — refresh_hr is the expected cycle (e.g., 6 hr for NWS, 1 hr for buoys)
    if age_hr < refresh_hr * 2:
        return ("FRESH", age_str)
    if age_hr < refresh_hr * 12:
        return ("STALE", age_str)
    return ("ARCHIVED", age_str)


def assess(text: str, kind: str, refresh_hr: float,
           now_utc: datetime | None = None) -> dict:
    """Combined parse + tier. Returns:
        {
            "issued_utc": datetime or None,
            "tier": "FRESH" | "STALE" | "ARCHIVED" | "UNKNOWN",
            "age_str": "8.1 hr",
        }
    """
    issued = extract_issuance(text, kind)
    tier, age_str = tier_freshness(issued, refresh_hr, now_utc)
    return {"issued_utc": issued, "tier": tier, "age_str": age_str}
