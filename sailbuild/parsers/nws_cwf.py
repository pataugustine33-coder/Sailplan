"""
NWS Coastal Waters Forecast (CWF) parser.

Input: raw text of an NWS CWF bulletin (e.g., FZUS52 KCHS).
Output: dict of {zone: {period: {wind, sea, wave_detail, SCA, ...}}}

Bulletin structure:
  FZUS52 KCHS 121518          ← WMO header (UTC)
  CWFCHS
  Coastal Waters Forecast
  ...
  1118 AM EDT Tue May 12 2026  ← local issuance
  ...
  AMZ300-130530-               ← zone header (zone + UGC expiration)
  1118 AM EDT Tue May 12 2026
  .Synopsis for ...            ← (zone 300 is synopsis-only)
  ...
  $$
  AMZ360-130530-               ← next zone
  ...SMALL CRAFT ADVISORY IN EFFECT...
  .TODAY...
  NE winds 20 to 25 kt...
  .TONIGHT...
  E winds 15 to 20 kt...
  ...
  $$

Each "$$" terminates a zone block. Each period within a zone starts with a
".PERIOD..." line (e.g., ".TODAY...", ".WEDNESDAY...", ".WED NIGHT...").
"""
import re
from datetime import datetime


# Wind direction word → compass degrees (centered).
WIND_DIR_DEG = {
    "N": 360, "NNE": 22, "NE": 45, "ENE": 67,
    "E": 90, "ESE": 112, "SE": 135, "SSE": 157,
    "S": 180, "SSW": 202, "SW": 225, "WSW": 247,
    "W": 270, "WNW": 292, "NW": 315, "NNW": 337,
}

# Common period name normalizations (NWS day-of-week → our short period codes).
PERIOD_NAMES = {
    "TODAY": "TODAY", "TONIGHT": "TONIGHT", "REST OF TODAY": "TODAY",
    "THIS AFTERNOON": "TODAY", "THIS MORNING": "TODAY", "REST OF TONIGHT": "TONIGHT",
    "OVERNIGHT": "TONIGHT", "THIS EVENING": "TONIGHT",
    "MON": "MON", "MONDAY": "MON", "MON NIGHT": "MON_NIGHT", "MONDAY NIGHT": "MON_NIGHT",
    "TUE": "TUE", "TUESDAY": "TUE", "TUE NIGHT": "TUE_NIGHT", "TUESDAY NIGHT": "TUE_NIGHT",
    "WED": "WED", "WEDNESDAY": "WED", "WED NIGHT": "WED_NIGHT", "WEDNESDAY NIGHT": "WED_NIGHT",
    "THU": "THU", "THURSDAY": "THU", "THU NIGHT": "THU_NIGHT", "THURSDAY NIGHT": "THU_NIGHT",
    "FRI": "FRI", "FRIDAY": "FRI", "FRI NIGHT": "FRI_NIGHT", "FRIDAY NIGHT": "FRI_NIGHT",
    "SAT": "SAT", "SATURDAY": "SAT", "SAT NIGHT": "SAT_NIGHT", "SATURDAY NIGHT": "SAT_NIGHT",
    "SUN": "SUN", "SUNDAY": "SUN", "SUN NIGHT": "SUN_NIGHT", "SUNDAY NIGHT": "SUN_NIGHT",
}


def parse_cwf(text: str) -> dict:
    """Parse a CWF bulletin into a structured dict.

    Returns:
        {
            "issuance": {"utc_wmo": "121518", "local_label": "1118 AM EDT Tue May 12 2026", ...},
            "office": "CHS",
            "synopsis": {"zone": "AMZ300", "text": "..."},
            "zones": {
                "AMZ360": {
                    "description": "...",
                    "SCA": {"in_effect_until": "Wed 3 AM EDT", "text": "..."} or None,
                    "periods": {"WED": {...}, "WED_NIGHT": {...}, ...},
                },
                ...
            }
        }
    """
    out = {"issuance": _parse_issuance(text), "office": _parse_office(text),
           "synopsis": None, "zones": {}}

    # Split into blocks by $$ terminator
    blocks = text.split("$$")
    for block in blocks:
        block = block.strip()
        if not block:
            continue

        # Synopsis block: has ".SYNOPSIS..." but may not have an AMZ zone header
        # (the synopsis is typically the first block before any zone).
        if ".SYNOPSIS" in block.upper():
            zone_match = re.search(r"^(AMZ\d{3})", block, re.MULTILINE)
            zone = zone_match.group(1) if zone_match else "synopsis"
            out["synopsis"] = {
                "zone": zone,
                "text": _extract_synopsis(block),
            }
            continue

        zone_match = re.search(r"^(AMZ\d{3})-\d+-", block, re.MULTILINE)
        if not zone_match:
            continue
        zone = zone_match.group(1)

        zone_data = _parse_zone_block(block, zone)
        if zone_data:
            out["zones"][zone] = zone_data

    return out


def _parse_issuance(text: str) -> dict:
    """Extract issuance timestamp info."""
    out = {"utc_wmo": None, "local_label": None}
    m = re.search(r"^FZUS\d{2} (K\w{3}) (\d{6})", text, re.MULTILINE)
    if m:
        out["utc_wmo"] = m.group(2)
        out["office_id"] = m.group(1)
    m = re.search(r"^(\d{1,4} (?:AM|PM) [A-Z]{3} \w{3} \w{3} \d+ \d{4})$", text, re.MULTILINE)
    if m:
        out["local_label"] = m.group(1)
    return out


def _parse_office(text: str) -> str:
    """Return office code (CHS, JAX, MHX, etc.) from product header."""
    m = re.search(r"^CWF(\w{3})$", text, re.MULTILINE)
    if m:
        return m.group(1)
    m = re.search(r"^FZUS\d{2} K(\w{3})", text, re.MULTILINE)
    if m:
        return m.group(1)
    return "?"


def _extract_synopsis(block: str) -> str:
    """Pull synopsis text body — only the text after .Synopsis... marker."""
    # Look for ".Synopsis...<text>" (NWS uses both ".SYNOPSIS..." and ".Synopsis for...")
    m = re.search(r"\.Synopsis[^.]*?\.\.\.\s*(.+?)(?=\$\$|\Z)", block, re.DOTALL | re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return ""


def _parse_zone_block(block: str, zone: str) -> dict:
    """Parse a single zone block into description, SCA, periods."""
    # First line(s) after zone header form description
    lines = block.split("\n")
    desc_lines = []
    description = ""
    for i, line in enumerate(lines):
        if line.startswith(zone):
            # next several non-empty lines until we hit timestamp or SCA or period
            j = i + 1
            while j < len(lines):
                ln = lines[j].strip()
                if not ln:
                    j += 1
                    continue
                if ln.startswith(".") or "SMALL CRAFT" in ln.upper() or re.match(r"^\d+ (AM|PM)", ln):
                    break
                desc_lines.append(ln)
                j += 1
            break
    description = " ".join(desc_lines).strip("-").strip()

    # Hazard detection: SCA / GALE WARNING / STORM WARNING / HFWW
    sca = None
    hazard_match = re.search(
        r"\.\.\.((?:SMALL CRAFT ADVISORY|GALE WARNING|STORM WARNING|HURRICANE FORCE WIND WARNING)[^.]*?)\.\.\.",
        block, re.IGNORECASE,
    )
    if hazard_match:
        hazard_text = hazard_match.group(1).strip()
        sca = {
            "text": hazard_text,
            "in_effect_until": _extract_sca_expiry(hazard_text),
            "type": _classify_hazard(hazard_text),
        }

    # Period blocks: ".PERIOD NAME...text"
    periods = {}
    period_pattern = re.compile(r"\.([A-Z][A-Z\s]+?)\.\.\.(.+?)(?=\.[A-Z][A-Z\s]+?\.\.\.|\Z)", re.DOTALL)
    for m in period_pattern.finditer(block):
        period_name_raw = m.group(1).strip()
        period_body = m.group(2).strip()
        if "SYNOPSIS" in period_name_raw.upper():
            continue
        period_key = PERIOD_NAMES.get(period_name_raw.upper())
        if not period_key:
            continue
        periods[period_key] = _parse_period_body(period_body)

    return {
        "description": description,
        "SCA": sca,
        "periods": periods,
    }


def _extract_sca_expiry(sca_text: str) -> str:
    """Extract 'until X' phrase from SCA text."""
    m = re.search(r"in effect (until [^\.]+)", sca_text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"in effect (through [^\.]+)", sca_text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"in effect (from [^\.]+)", sca_text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return sca_text


def _classify_hazard(text: str) -> str:
    """Classify hazard into severity tier: SCA, GALE, STORM, HFWW."""
    t = text.upper()
    if "HURRICANE FORCE WIND" in t:
        return "HFWW"
    if "STORM WARNING" in t:
        return "STORM"
    if "GALE WARNING" in t:
        return "GALE"
    if "SMALL CRAFT" in t:
        return "SCA"
    return "OTHER"


def _parse_period_body(body: str) -> dict:
    """Parse one period's forecast text into structured fields.

    Example body:
      "NE winds 20 to 25 kt with gusts to 30 kt. Seas 4 to 6 ft, building to 5 to 7 ft this afternoon.
       Wave Detail: E 6 ft at 6 seconds and E 1 foot at 12 seconds. ..."
    """
    body = " ".join(body.split())  # normalize whitespace
    out = {"raw": body}

    # Wind: "<DIR> winds <LOW> to <HIGH> kt" or variations
    wind = _parse_wind(body)
    if wind:
        out.update(wind)

    # Seas: "Seas <LOW> to <HIGH> ft" (with optional "occasionally to <occ> ft")
    sea = _parse_sea(body)
    if sea:
        out.update(sea)

    # Wave Detail
    wd = _parse_wave_detail(body)
    if wd:
        out["wave_detail"] = wd

    return out


def _parse_wind(body: str):
    """Extract wind direction and speed range from period text."""
    # Pattern: "NE winds 20 to 25 kt" or "Northeast winds 20 to 25 knots"
    pattern = re.compile(
        r"\b("
        r"north(?:east|west)?(?:east|west|northeast|northwest)?|"
        r"south(?:east|west)?(?:east|west|southeast|southwest)?|"
        r"east(?:northeast|southeast)?|"
        r"west(?:northwest|southwest)?|"
        r"N|NNE|NE|ENE|E|ESE|SE|SSE|S|SSW|SW|WSW|W|WNW|NW|NNW"
        r")\s+winds?\s+"
        r"(?:around\s+)?"
        r"(\d+)(?:\s+to\s+(\d+))?\s*(?:kt|knots?)",
        re.IGNORECASE,
    )
    m = pattern.search(body)
    if not m:
        return None
    dir_word = m.group(1).upper()
    dir_abbrev = _wind_word_to_abbrev(dir_word)
    lo = int(m.group(2))
    hi = int(m.group(3)) if m.group(3) else lo
    return {
        "wind_dir": dir_abbrev,
        "wind_dir_deg": WIND_DIR_DEG.get(dir_abbrev, 0),
        "wind_kt": [lo, hi],
        "wind_text_raw": m.group(0),
    }


def _wind_word_to_abbrev(word: str) -> str:
    """Convert 'northeast' / 'NE' / 'east northeast' etc. → 'NE' / 'ENE'."""
    word = word.upper().strip()
    if word in WIND_DIR_DEG:
        return word
    # Handle compound words like "EAST NORTHEAST"
    mapping = {
        "NORTH": "N", "SOUTH": "S", "EAST": "E", "WEST": "W",
        "NORTHEAST": "NE", "NORTHWEST": "NW", "SOUTHEAST": "SE", "SOUTHWEST": "SW",
        "NORTHNORTHEAST": "NNE", "EASTNORTHEAST": "ENE",
        "EASTSOUTHEAST": "ESE", "SOUTHSOUTHEAST": "SSE",
        "SOUTHSOUTHWEST": "SSW", "WESTSOUTHWEST": "WSW",
        "WESTNORTHWEST": "WNW", "NORTHNORTHWEST": "NNW",
    }
    clean = word.replace(" ", "")
    return mapping.get(clean, word)


def _parse_sea(body: str):
    """Extract sea height range and 'occasionally to X' value.

    Accepts both 'Seas N to M ft' (offshore zones) and 'Waves N to M ft'
    (harbor zones). NWS uses 'Waves' for inland/harbor zones since the
    fetch isn't long enough for proper swell.
    """
    pattern = re.compile(
        r"(?:Seas|Waves)\s+(?:around\s+)?(\d+)(?:\s+to\s+(\d+))?\s*(?:ft|feet)"
        r"(?:.*?occasionally\s+to\s+(\d+)\s*(?:ft|feet))?",
        re.IGNORECASE,
    )
    m = pattern.search(body)
    if not m:
        return None
    lo = int(m.group(1))
    hi = int(m.group(2)) if m.group(2) else lo
    occ = int(m.group(3)) if m.group(3) else None
    out = {"sea_ft": [lo, hi]}
    if occ:
        out["sea_ft_occ"] = occ
    return out


def _parse_wave_detail(body: str):
    """Parse 'Wave Detail: X 6 ft at 7 seconds and Y 2 ft at 11 seconds' field."""
    m = re.search(r"Wave Detail:\s*(.+?)(?:\.\s+[A-Z]|$)", body)
    if not m:
        return None
    wd_text = m.group(1)
    # Find each component
    component_pattern = re.compile(
        r"\b("
        r"north(?:east|west)?(?:east|west|northeast|northwest)?|"
        r"south(?:east|west)?(?:east|west|southeast|southwest)?|"
        r"east(?:northeast|southeast)?|"
        r"west(?:northwest|southwest)?|"
        r"N|NNE|NE|ENE|E|ESE|SE|SSE|S|SSW|SW|WSW|W|WNW|NW|NNW"
        r")\s+(\d+)\s*(?:ft|feet|foot)\s+at\s+(\d+)\s*seconds?",
        re.IGNORECASE,
    )
    components = []
    for cm in component_pattern.finditer(wd_text):
        dir_word = cm.group(1).upper()
        dir_abbrev = _wind_word_to_abbrev(dir_word)
        components.append({
            "direction": dir_abbrev,
            "direction_deg": WIND_DIR_DEG.get(dir_abbrev, 0),
            "height_ft": int(cm.group(2)),
            "period_s": int(cm.group(3)),
        })
    if not components:
        return None
    out = {"primary": components[0]}
    if len(components) > 1:
        out["secondary"] = components[1]
    return out
