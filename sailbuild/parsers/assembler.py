"""
Forecast bundle assembler.

Takes parsed CWF + AFD outputs and produces a forecast.yaml dict that the
build script can consume.

The assembler does the structural work of combining multiple parsed inputs
into the schema build.py expects. It does NOT make weather judgments — those
remain in the human-edited forecast.yaml fields (waypoint_assignments,
plan_commentary, notes_addendum). The assembler emits a skeleton; the analyst
edits in the route-specific narrative.
"""
from datetime import datetime


def assemble_forecast(cwfs: dict, afds: dict, gulf_stream: dict = None) -> dict:
    """Build a forecast bundle dict.

    Args:
        cwfs: {"CHS": parsed_cwf_chs, "JAX": parsed_cwf_jax, ...}
        afds: {"CHS": parsed_afd_chs, "JAX": parsed_afd_jax, ...}
        gulf_stream: optional {"effective_date": ..., "positions": [...]}
    """
    bundle = {
        "cycle": _assemble_cycle(cwfs, afds),
        "zones": _assemble_zones(cwfs),
        "synopses": _assemble_synopses(cwfs),
        "afd_marine": _assemble_afd_marine(afds),
    }
    if gulf_stream:
        bundle["gulf_stream"] = gulf_stream
    return bundle


def _assemble_cycle(cwfs, afds) -> dict:
    """Build cycle metadata block."""
    cycle = {"label": ""}
    for office, parsed in cwfs.items():
        key = f"cwf{office.lower()}"
        issuance = parsed.get("issuance", {})
        cycle[key] = {
            "issued": _to_iso(issuance.get("local_label", "")),
            "label_short": _short_label(issuance.get("local_label", "")),
            "label_full": issuance.get("local_label", ""),
        }
    for office, parsed in afds.items():
        key = f"afd{office.lower()}"
        issuance = parsed.get("issuance", {})
        cycle[key] = {
            "issued": _to_iso(issuance.get("local_label", "")),
            "label_short": _short_label(issuance.get("local_label", "")),
            "label_full": issuance.get("local_label", ""),
        }
    # Build human cycle label
    parts = []
    for office, parsed in cwfs.items():
        s = parsed.get("issuance", {}).get("local_label", "")
        if s:
            parts.append(f"CWF{office} {_short_label(s)}")
    cycle["label"] = " / ".join(parts) + " cycle"
    return cycle


def _assemble_zones(cwfs) -> dict:
    """Merge zone forecasts from all CWFs into a single zones dict."""
    zones = {}
    for office, parsed in cwfs.items():
        for zone_id, zone_data in parsed.get("zones", {}).items():
            entry = {
                "description": zone_data.get("description", ""),
                "office": office,
                "periods": zone_data.get("periods", {}),
            }
            if zone_data.get("SCA"):
                entry["SCA"] = {
                    "in_effect_until": zone_data["SCA"].get("in_effect_until", ""),
                    "reason": zone_data["SCA"].get("text", ""),
                    "type": zone_data["SCA"].get("type", "OTHER"),
                }
            # Strip "raw" field from periods to keep the YAML clean
            cleaned_periods = {}
            for period_name, period_data in entry["periods"].items():
                cleaned = {k: v for k, v in period_data.items() if k != "raw" and k != "wind_text_raw"}
                cleaned_periods[period_name] = cleaned
            entry["periods"] = cleaned_periods
            zones[zone_id] = entry
    return zones


def _assemble_synopses(cwfs) -> dict:
    """Pull synopsis text from each CWF."""
    out = {}
    for office, parsed in cwfs.items():
        syn = parsed.get("synopsis")
        if syn:
            out[office] = {
                "zone": syn["zone"],
                "text": syn["text"],
            }
    return out


def _assemble_afd_marine(afds) -> dict:
    """Pull marine section from each AFD."""
    return {office: parsed.get("marine", "") for office, parsed in afds.items()}


def _to_iso(local_label: str) -> str:
    """Convert '1118 AM EDT Tue May 12 2026' to ISO datetime str."""
    if not local_label:
        return ""
    try:
        # Try common formats
        for fmt in ["%I%M %p %Z %a %b %d %Y", "%I %p %Z %a %b %d %Y"]:
            try:
                dt = datetime.strptime(local_label, fmt)
                return dt.isoformat() + "-04:00"
            except ValueError:
                pass
        # Manual parse for "1118 AM EDT Tue May 12 2026"
        parts = local_label.split()
        if len(parts) >= 6:
            time_str = parts[0]
            ampm = parts[1]
            month_name = parts[3]
            day = int(parts[4])
            year = int(parts[5])
            if len(time_str) == 3:
                h = int(time_str[0])
                m = int(time_str[1:])
            else:
                h = int(time_str[:-2])
                m = int(time_str[-2:])
            if ampm == "PM" and h != 12:
                h += 12
            elif ampm == "AM" and h == 12:
                h = 0
            month_map = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                         "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
            month = month_map.get(month_name, 1)
            return f"{year:04d}-{month:02d}-{day:02d}T{h:02d}:{m:02d}:00-04:00"
    except Exception:
        return local_label
    return local_label


def _short_label(local_label: str) -> str:
    """Build short label like 'Tue 5/12 1118 AM' from full local label.

    Input format: '1118 AM EDT Tue May 12 2026' → parts:
        [0]=1118, [1]=AM, [2]=EDT, [3]=Tue, [4]=May, [5]=12, [6]=2026
    """
    if not local_label:
        return ""
    parts = local_label.split()
    if len(parts) >= 7:
        time_str = parts[0]
        ampm = parts[1]
        dow = parts[3]
        month_name = parts[4]
        day = parts[5]
        month_map = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                     "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
        month = month_map.get(month_name, 0)
        if month and day.isdigit():
            return f"{dow} {month}/{int(day)} {time_str} {ampm}"
    return local_label
