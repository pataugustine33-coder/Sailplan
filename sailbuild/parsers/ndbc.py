"""
NDBC buoy observation parser.

Supports two endpoint formats:
  (1) latest_obs/{station}.txt — compact most-recent reading
  (2) realtime2/{station}.txt — last ~45 days of 10-minute obs

For now, parses latest_obs format which looks like:
    Station 41033
    32° 16.7' N  80° 24.4' W

    11:08 am EDT
    1508 GMT 05/12/26
    Wind: ENE (70°), 23.3 kt
    Gust: 31.1 kt
    Pres: 30.17
    Air Temp: 69.4 °F
    Water Temp: 71.2 °F
    Dew Point: 58.8 °F

And wave-only stations (no wind/pressure sensor):
    Station 41112
    30° 42.5' N  81° 17.5' W

    11:56 am EDT
    1556 GMT 05/12/26
    Seas: 6.6 ft
    Peak Period: 7 sec
    ...

    Wave Summary
    11:56 am EDT
    1500 GMT 05/12/26
    Swell: 1.0 ft
    Period: 10.5 sec
    Direction: E
    Wind Wave: 6.6 ft
    Period: 7.1 sec
    Direction: NE
"""
import re


def parse_latest_obs(text: str) -> dict:
    """Parse one latest_obs response.

    Returns:
        {
            "station_id": "41033",
            "lat": 32.279, "lon": -80.407,
            "reading_time": "11:08 am EDT",
            "reading_utc": "1508 GMT 05/12/26",
            "wind_dir": "ENE (70°)" or None,
            "wind_kt": 23.3 or None,
            "gust_kt": 31.1 or None,
            "pressure_inhg": 30.17 or None,
            "seas_ft": 6.6 or None,
            "peak_period_s": 7 or None,
            "wave_summary": {
                "swell": {"ft": 1.0, "period_s": 10.5, "direction": "E"},
                "wind_wave": {"ft": 6.6, "period_s": 7.1, "direction": "NE"},
            } or None,
        }
    """
    out = {}

    # Station ID
    m = re.search(r"Station\s+(\d+)", text)
    if m:
        out["station_id"] = m.group(1)

    # Lat/lon — "32° 16.7' N  80° 24.4' W"
    ll = re.search(r"(\d+)°\s*(\d+\.?\d*)'?\s*([NS])\s+(\d+)°\s*(\d+\.?\d*)'?\s*([EW])", text)
    if ll:
        lat = int(ll.group(1)) + float(ll.group(2)) / 60.0
        if ll.group(3) == "S":
            lat = -lat
        lon = int(ll.group(4)) + float(ll.group(5)) / 60.0
        if ll.group(6) == "W":
            lon = -lon
        out["lat"] = round(lat, 4)
        out["lon"] = round(lon, 4)

    # Reading time
    m = re.search(r"(\d{1,2}:\d{2}\s*(?:am|pm|AM|PM)\s+EDT)", text)
    if m:
        out["reading_time"] = m.group(1)
    m = re.search(r"(\d{4}\s+GMT\s+\d{2}/\d{2}/\d{2})", text)
    if m:
        out["reading_utc"] = m.group(1)

    # Wind
    m = re.search(r"Wind:\s+([^,]+),\s+([\d.]+)\s+kt", text)
    if m:
        out["wind_dir"] = m.group(1).strip()
        out["wind_kt"] = float(m.group(2))

    # Gust
    m = re.search(r"Gust:\s+([\d.]+)\s+kt", text)
    if m:
        out["gust_kt"] = float(m.group(1))

    # Pressure
    m = re.search(r"Pres:\s+([\d.]+)", text)
    if m:
        out["pressure_inhg"] = float(m.group(1))

    # Seas / peak period
    m = re.search(r"Seas:\s+([\d.]+)\s+ft", text)
    if m:
        out["seas_ft"] = float(m.group(1))
    m = re.search(r"Peak Period:\s+([\d.]+)\s+sec", text)
    if m:
        out["peak_period_s"] = float(m.group(1))

    # Wave Summary block (swell + wind wave components)
    if "Wave Summary" in text:
        ws_block = text[text.index("Wave Summary"):]
        wave_summary = {}
        swell = _extract_wave_component(ws_block, "Swell")
        if swell:
            wave_summary["swell"] = swell
        wind_wave = _extract_wave_component(ws_block, "Wind Wave")
        if wind_wave:
            wave_summary["wind_wave"] = wind_wave
        if wave_summary:
            out["wave_summary"] = wave_summary

    return out


def _extract_wave_component(text: str, label: str) -> dict:
    """Extract one wave component (Swell or Wind Wave)."""
    # Pattern:
    #   Swell: 1.0 ft
    #   Period: 10.5 sec
    #   Direction: E
    pattern = re.compile(
        re.escape(label) + r":\s+([\d.]+)\s+ft\s+Period:\s+([\d.]+)\s+sec\s+Direction:\s+(\w+)",
        re.IGNORECASE,
    )
    # Also try multiline version
    if not pattern.search(text):
        pattern = re.compile(
            re.escape(label) + r":\s+([\d.]+)\s+ft\s*\n+\s*Period:\s+([\d.]+)\s+sec\s*\n+\s*Direction:\s+(\w+)",
            re.IGNORECASE | re.MULTILINE,
        )
    m = pattern.search(text)
    if not m:
        return None
    return {
        "ft": float(m.group(1)),
        "period_s": float(m.group(2)),
        "direction": m.group(3).strip(),
    }


def format_wave_summary(parsed: dict) -> str:
    """Render a wave-only buoy reading as human-readable string."""
    parts = []
    if "seas_ft" in parsed and "peak_period_s" in parsed:
        parts.append(f"Seas {parsed['seas_ft']} ft @ {parsed['peak_period_s']:.0f}s")
    ws = parsed.get("wave_summary", {})
    if "wind_wave" in ws:
        ww = ws["wind_wave"]
        parts.append(f"Wind Wave: {ww['direction']} {ww['ft']} ft @ {ww['period_s']:.1f}s")
    if "swell" in ws:
        sw = ws["swell"]
        parts.append(f"Swell: {sw['direction']} {sw['ft']} ft @ {sw['period_s']:.1f}s")
    return " | ".join(parts)


def parse_station_page(text: str, station_id: str = None) -> dict:
    """Parse an NDBC station_page.php response (HTML or markdown-rendered).

    Tolerant of both the raw HTML and the markdown-rendered version that
    web fetchers produce. Looks for table rows with labels like:
        Wind Direction (WDIR): | SSW ( 200 deg true )
        Wind Speed (WSPD):     | 5.8 kts
        Atmospheric Pressure (PRES): | 30.20 in
        Significant Wave Height (WVHT): | 2.6 ft

    Also tries to extract the latest observation timestamp from headers
    like "Conditions at 41008 as of (11:40 am EDT)".

    Returns the same dict shape as parse_latest_obs.
    """
    out = {}
    if station_id:
        out["station_id"] = station_id
    else:
        m = re.search(r"Station\s+(\d+)", text)
        if m:
            out["station_id"] = m.group(1)

    # Reading time — "Conditions at 41008 as of (11:40 am EDT)"
    m = re.search(r"as of\s*\(?\s*(\d{1,2}:\d{2}\s*(?:am|pm|AM|PM)\s+\w+)", text)
    if m:
        out["reading_time"] = m.group(1).strip()
    # GMT timestamp — "1540 GMT on 04/15/2026"
    m = re.search(r"(\d{4}\s+GMT(?:\s+on)?\s+\d{2}/\d{2}/\d{2,4})", text)
    if m:
        out["reading_utc"] = m.group(1)

    # Wind direction — "SSW ( 200 deg true )" or "SSW (200 deg true)"
    m = re.search(r"Wind Direction[^|]*\|\s*([NEWS]+)\s*\(\s*(\d+)\s*deg", text, re.IGNORECASE)
    if m:
        out["wind_dir"] = m.group(1)
        out["wind_dir_deg"] = int(m.group(2))

    # Wind speed
    m = re.search(r"Wind Speed[^|]*\|\s*([\d.]+)\s*kts?", text, re.IGNORECASE)
    if m:
        out["wind_kt"] = float(m.group(1))

    # Wind gust
    m = re.search(r"Wind Gust[^|]*\|\s*([\d.]+)\s*kts?", text, re.IGNORECASE)
    if m:
        out["gust_kt"] = float(m.group(1))

    # Pressure
    m = re.search(r"Atmospheric Pressure[^|]*\|\s*([\d.]+)\s*in", text, re.IGNORECASE)
    if m:
        out["pressure_inhg"] = float(m.group(1))

    # Pressure tendency (3-hour) — useful for derive_pressure_trends
    m = re.search(r"Pressure Tendency[^|]*\|\s*([+\-][\d.]+)", text, re.IGNORECASE)
    if m:
        out["pressure_tendency_3hr"] = float(m.group(1))

    # Significant wave height
    m = re.search(r"Significant Wave Height[^|]*\|\s*([\d.]+)\s*ft", text, re.IGNORECASE)
    if m:
        out["seas_ft"] = float(m.group(1))

    # Wave summary block (Swell + Wind Wave + direction)
    wave_summary = {}
    m = re.search(r"Swell Height[^|]*\|\s*([\d.]+)\s*ft.*?Swell Period[^|]*\|\s*([\d.]+)\s*sec.*?Swell Direction[^|]*\|\s*(\w+)",
                  text, re.IGNORECASE | re.DOTALL)
    if m:
        wave_summary["swell"] = {
            "ft": float(m.group(1)),
            "period_s": float(m.group(2)),
            "direction": m.group(3).strip(),
        }
    m = re.search(r"Wind Wave Height[^|]*\|\s*([\d.]+)\s*ft.*?Wind Wave Period[^|]*\|\s*([\d.]+)\s*sec.*?Wind Wave Direction[^|]*\|\s*(\w+)",
                  text, re.IGNORECASE | re.DOTALL)
    if m:
        wave_summary["wind_wave"] = {
            "ft": float(m.group(1)),
            "period_s": float(m.group(2)),
            "direction": m.group(3).strip(),
        }
    if wave_summary:
        out["wave_summary"] = wave_summary

    # Detect offline status — station_page shows this when no data <8hr
    if re.search(r"no data in last 8 hours", text, re.IGNORECASE):
        # Only flag offline if the latest observation timestamp is also old.
        # The boilerplate legend always contains this phrase, so we need
        # the reading_time absence too.
        if "wind_kt" not in out and "seas_ft" not in out:
            out["status"] = "offline"

    return out


def parse_buoy(text: str, station_id: str = None) -> dict:
    """Auto-detect format and dispatch to the right parser.

    Discriminator: latest_obs format starts with "Station N" alone,
    has terse "Wind: ..." lines. station_page.php has Markdown table
    rows with pipes.
    """
    if "|" in text and ("Wind Direction" in text or "Atmospheric Pressure" in text):
        return parse_station_page(text, station_id)
    return parse_latest_obs(text)
