"""
NWS Area Forecast Discussion (AFD) parser.

Input: raw text of an NWS AFD bulletin (e.g., FXUS62 KCHS).
Output: dict with key messages, marine section, near-term, short-term, long-term.

Bulletin structure:
  FXUS62 KCHS 121114
  AFDCHS
  Area Forecast Discussion
  National Weather Service Charleston SC
  714 AM EDT Tue May 12 2026

  .KEY MESSAGES...
  - 1) ...
  - 2) ...

  &&

  .DISCUSSION...
  ...

  &&

  .NEAR TERM /THROUGH TONIGHT/...
  ...

  &&

  .MARINE...
  Today and tonight: ...
  Extended Marine: ...

  &&

  .WATCHES/WARNINGS/ADVISORIES...
  ...

  $$
"""
import re


def parse_afd(text: str) -> dict:
    """Parse AFD into sections.

    Returns:
        {
            "issuance": {"utc_wmo": "...", "local_label": "...", "office_id": "KCHS"},
            "office": "CHS",
            "key_messages": ["...", "..."],
            "near_term": "...",
            "short_term": "...",
            "long_term": "...",
            "marine": "...",
            "aviation": "...",
            "watches_warnings_advisories": "...",
        }
    """
    out = {
        "issuance": _parse_issuance(text),
        "office": _parse_office(text),
        "key_messages": _extract_key_messages(text),
        "near_term": _extract_section(text, ["NEAR TERM"]),
        "short_term": _extract_section(text, ["SHORT TERM"]),
        "long_term": _extract_section(text, ["LONG TERM"]),
        "marine": _extract_section(text, ["MARINE"]),
        "aviation": _extract_section(text, ["AVIATION"]),
        "fire_weather": _extract_section(text, ["FIRE WEATHER"]),
        "watches_warnings_advisories": _extract_section(text, ["WATCHES/WARNINGS/ADVISORIES"]),
    }
    return out


def _parse_issuance(text: str) -> dict:
    out = {"utc_wmo": None, "local_label": None, "office_id": None}
    m = re.search(r"^FXUS\d{2} (K\w{3}) (\d{6})", text, re.MULTILINE)
    if m:
        out["office_id"] = m.group(1)
        out["utc_wmo"] = m.group(2)
    m = re.search(r"^(\d{1,4} (?:AM|PM) EDT \w{3} \w{3} \d+ \d{4})$", text, re.MULTILINE)
    if m:
        out["local_label"] = m.group(1)
    return out


def _parse_office(text: str) -> str:
    m = re.search(r"^AFD(\w{3})$", text, re.MULTILINE)
    if m:
        return m.group(1)
    m = re.search(r"^FXUS\d{2} K(\w{3})", text, re.MULTILINE)
    if m:
        return m.group(1)
    return "?"


def _extract_key_messages(text: str) -> list:
    """Extract bulleted KEY MESSAGES list."""
    m = re.search(r"\.KEY MESSAGES\.\.\.(.+?)(?=&&|\Z)", text, re.DOTALL)
    if not m:
        return []
    body = m.group(1)
    msgs = []
    # KEY MESSAGES are typically formatted as "- 1)" or "- " items
    bullet_pattern = re.compile(r"-\s*(?:\d+\))?\s*(.+?)(?=\n-|\n\n|\Z)", re.DOTALL)
    for bm in bullet_pattern.finditer(body):
        msg = " ".join(bm.group(1).split())
        if msg:
            msgs.append(msg)
    return msgs


def _extract_section(text: str, section_names: list) -> str:
    """Extract a section by name (e.g., 'MARINE', 'NEAR TERM').

    NWS AFDs sometimes prefix section names with the office code, e.g.
    '.CHS WATCHES/WARNINGS/ADVISORIES...' rather than '.WATCHES/...'. We
    accept an optional 3-letter office prefix before the section name.
    """
    for name in section_names:
        pattern = re.compile(
            r"\.(?:[A-Z]{3}\s+)?" + re.escape(name) + r"\s*(?:/[^/]+/)?\s*\.\.\.(.+?)(?=&&|\$\$|\Z)",
            re.DOTALL,
        )
        m = pattern.search(text)
        if m:
            body = m.group(1).strip()
            return " ".join(body.split())
    return ""
