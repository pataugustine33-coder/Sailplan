"""
Compute engine. Takes a passage + forecast + plan ID, produces a list of
fully-resolved leg dictionaries with all derived values.

A "leg" here means the row for one waypoint on a Plan tab. Each leg has:
  - WP id, course, cum_nm
  - ETA (str + decimal hour for color), color code
  - Wind dir (text, degrees), TWS range, sea height range, period
  - Sea From (deg, label), Sea Angle, source tag (WD = Wave Detail, SE = Synoptic Estimate)
  - TWA
  - Polar speed (pure), Boat speed (adjusted), sail mode
  - AWS, AWA
  - Cumulative time / sailing / motoring
  - Notes, Weather Risk, Sea Source
"""
from dataclasses import dataclass, field
from typing import Optional
import math
from .polar import polar_speed, select_sea_factor, apparent_wind


def _haversine_nm(lat1, lon1, lat2, lon2):
    """Great-circle distance in nautical miles between two lat/lon points."""
    R = 3440.065  # earth radius in NM
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


DAY_ORDER = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]


@dataclass
class Leg:
    """One row on a Plan tab."""
    wp_id: str
    wp_name: str
    cum_nm: float
    course_out: Optional[int]
    eta_str: str
    eta_decimal_hour: float
    eta_color: str          # "C6EFCE" green, "FFEB9C" yellow, "FFC7CE" red
    wind_dir_text: str
    wind_dir_deg: int
    wind_kt_low: float
    wind_kt_high: float
    wind_kt_avg: float
    gust_kt: Optional[float]   # forecast gust ceiling for this WP period, None if not in YAML
    sea_ft_low: float
    sea_ft_high: float
    sea_period_s: float
    sea_from_deg: int
    sea_from_label: str
    sea_source_tag: str     # "WD" or "SE"
    sea_angle: int
    twa: int
    polar_speed: float
    boat_speed: float
    sail_mode: str
    is_motor: bool
    aws: float
    awa: int
    cum_time_hr: float
    cum_sailing_hr: float
    cum_motoring_hr: float
    notes: str
    weather_risk: str
    sea_source: str
    risk_color: str = "C6EFCE"   # green / yellow / red for Weather Risk cell
    row_band: Optional[str] = None  # None, "amber", "salmon", "light_green", "yellow"
    pressure_inhg: Optional[float] = None   # synoptic forecast pressure at this WP/time
    pressure_trend: Optional[str] = None    # vocab: Rising fast/Rising/Rising slow/Steady/Falling slow/Falling/Falling fast/Bottoming
    sea_position: str = ""      # HD/PB/PQ/ST/SQ/SB — which side the boat feels the sea
    leg_distance_nm: float = 0.0
    leg_hours: float = 0.0
    chart_label: str = ""       # Short city/location label for chart x-axis (e.g. "Vero Beach"); fallback to heuristic on wp_name if not set


MIN_SAILING_TWA = 45  # below this, boat must tack (VMG ~5 kt < motor 6 kt) → motor


RISK_COLORS = {
    "green": "C6EFCE",
    "yellow": "FFEB9C",
    "red": "FFC7CE",
}

# Whole-row banding palette (matches manual workbook conventions)
ROW_BAND_COLORS = {
    "amber": "FFE699",        # pre-frontal / pressure-falling
    "salmon": "F8CBAD",       # tactical hazard (wind on nose, tack required)
    "light_green": "E2EFDA",  # light-air recovery / favorable
    "yellow": "FFEB9C",       # high-attention departure into rough conditions
}


def auto_risk_level(*, wind_kt_high, sea_ft_high, eta_color, wind_dir_text):
    """Default risk level if not explicitly specified in YAML.

    Rules:
      - red    : peak wind ≥ 25 OR peak sea ≥ 8 OR ETA color is red (night arrival)
      - yellow : peak wind ≥ 18 OR peak sea ≥ 6 OR ETA color is yellow (twilight)
                 OR wind direction has transition (→) indicating shift
      - green  : otherwise
    """
    if wind_kt_high >= 25 or sea_ft_high >= 8 or eta_color == "FFC7CE":
        return "red"
    if wind_kt_high >= 18 or sea_ft_high >= 6 or eta_color == "FFEB9C":
        return "yellow"
    if "→" in str(wind_dir_text):
        return "yellow"
    return "green"


def _to_int_deg(value) -> int:
    """Coerce a bearing value to int. YAML strings like '080' (leading zero)
    parse as strings, so we accept either type."""
    if value is None:
        return 0
    if isinstance(value, int):
        return value
    return int(str(value).strip())


def compute_twa(course_deg: int, wind_dir_deg: int) -> int:
    """True wind angle (0-180) between course and wind direction."""
    course_deg = _to_int_deg(course_deg)
    wind_dir_deg = _to_int_deg(wind_dir_deg)
    diff = abs(course_deg - wind_dir_deg)
    return round(min(diff, 360 - diff))


def compute_sea_angle(course_deg: int, sea_from_deg: int) -> int:
    """Sea angle (0-180) between course and sea-from direction."""
    course_deg = _to_int_deg(course_deg)
    sea_from_deg = _to_int_deg(sea_from_deg)
    diff = abs(course_deg - sea_from_deg)
    return round(min(diff, 360 - diff))


def sea_position_label(course_deg: int, sea_from_deg: int) -> str:
    """Position abbreviation describing how the boat feels the sea.

    Returns one of: HD, PB, PQ, ST, SQ, SB.

    The signed relative bearing tells which side:
      diff = sea_from - course, normalized to -180..+180
      negative = port, positive = starboard
    """
    course_deg = _to_int_deg(course_deg)
    sea_from_deg = _to_int_deg(sea_from_deg)
    diff = sea_from_deg - course_deg
    while diff > 180:
        diff -= 360
    while diff < -180:
        diff += 360
    abs_a = abs(diff)
    if abs_a <= 30:
        return "HD"
    if abs_a >= 150:
        return "ST"
    is_port = diff < 0
    if is_port:
        return "PB" if abs_a <= 90 else "PQ"
    return "SB" if abs_a <= 90 else "SQ"


def eta_color_for_hour(hour_24: float, arrival_timing: dict) -> str:
    """Color code an ETA cell based on time of day at the destination."""
    h = hour_24 % 24
    day_start, day_end = arrival_timing["day_window"]
    if day_start <= h <= day_end:
        return "C6EFCE"  # green
    for tw_start, tw_end in arrival_timing["twilight_windows"]:
        if tw_start <= h <= tw_end:
            return "FFEB9C"  # yellow
    return "FFC7CE"  # red


def hours_to_eta(total_hr: float, dep_day: str, dep_hour: float) -> tuple[str, float]:
    """Convert hours-from-departure to ('Day H:MM AM/PM', decimal_hour_of_day)."""
    abs_h = dep_hour + total_hr
    day_offset = int(abs_h // 24)
    rem = abs_h - day_offset * 24
    h = int(rem)
    m = int(round((rem - h) * 60))
    if m == 60:
        h += 1
        m = 0
    if h >= 24:
        h -= 24
        day_offset += 1
    dep_idx = DAY_ORDER.index(dep_day)
    day = DAY_ORDER[(dep_idx + day_offset) % 7]
    if h == 0:
        h12, ampm = 12, "AM"
    elif h < 12:
        h12, ampm = h, "AM"
    elif h == 12:
        h12, ampm = 12, "PM"
    else:
        h12, ampm = h - 12, "PM"
    return f"{day} {h12}:{m:02d} {ampm}", (abs_h % 24)


def avg(lo, hi):
    return (lo + hi) / 2


def get_cycle_label_for_office(forecast: dict, office: str) -> str:
    """Look up the CWF cycle label_short for a given office, e.g. 'CHS' → 'Tue 5/12 1118 AM'.

    Works for any office (CHS, JAX, ILM, MHX, etc.) by checking cycle[cwf<office>].
    Falls back to the first cwf* entry if the specific office isn't found.
    """
    cycle = forecast.get("cycle", {})
    key = f"cwf{office.lower()}"
    if key in cycle:
        return cycle[key].get("label_short", "")
    # Fallback: any cwf* entry
    for k, v in cycle.items():
        if k.startswith("cwf"):
            return v.get("label_short", "")
    return ""


def iter_office_cycles(cycle: dict, prefix: str = "cwf"):
    """Yield (office_uppercase, cycle_dict) for each office in cycle metadata.

    Args:
        cycle: forecast["cycle"] dict
        prefix: "cwf" for CWFs, "afd" for AFDs

    Example: cycle = {"cwfchs": {...}, "cwfjax": {...}, "afdchs": {...}}
             iter_office_cycles(cycle, "cwf") yields ("CHS", {...}), ("JAX", {...})
    """
    for key, val in cycle.items():
        if key.startswith(prefix) and isinstance(val, dict):
            office = key[len(prefix):].upper()
            yield office, val


def get_period_data(forecast: dict, zone: str, period: str) -> dict:
    """Look up wind/sea data for a (zone, period) tuple in the forecast bundle."""
    if zone not in forecast["zones"]:
        raise KeyError(
            f"Zone {zone!r} not defined in forecast.yaml. Defined zones: "
            f"{sorted(forecast['zones'].keys())}"
        )
    zone_data = forecast["zones"][zone]
    if period not in zone_data["periods"]:
        raise KeyError(
            f"Period {period!r} not defined in zone {zone}. Defined periods: "
            f"{sorted(zone_data['periods'].keys())}. "
            f"Check waypoint_assignments — every (zone, period) pair must exist in zones."
        )
    return zone_data["periods"][period]


def resolve_leg_conditions(forecast: dict, assignment: dict) -> dict:
    """Combine zone-period forecast with any per-WP overrides into a flat dict.

    Defensive against terse bulletins: extended-range periods (Wed onward in
    typical NWS CWF) often drop the 'Wave Detail:' line, leaving only wind
    and sea height. When wave_detail is missing or empty, synthesize a
    primary component from the wind direction and full sea height, mark the
    sea_source as 'synoptic', and continue. This matches the operational
    convention where a deckhand reading the bulletin would assume the seas
    are wind-driven from the prevailing wind direction.
    """
    zone = assignment["zone"]
    period = assignment["period"]
    pd = get_period_data(forecast, zone, period)

    # Defensive wave_detail: synthesize from wind dir if bulletin was terse
    wave_detail = pd.get("wave_detail")
    wave_detail_synoptic = False
    if not wave_detail or not wave_detail.get("primary"):
        wind_dir_deg_pd = _to_int_deg(pd["wind_dir_deg"])
        sea_ft = pd["sea_ft"]
        sea_high = sea_ft[1] if isinstance(sea_ft, (list, tuple)) else sea_ft
        wave_detail = {
            "primary": {
                "direction": pd["wind_dir"],
                "direction_deg": wind_dir_deg_pd,
                "height_ft": sea_high,
                # Wind-driven seas at moderate wind: rule of thumb ~ TWS/3
                # In a real refresh the analyst can override this.
                "period_s": 5,
            },
            "secondary": None,
        }
        wave_detail_synoptic = True

    out = {
        "wind_dir_deg": _to_int_deg(pd["wind_dir_deg"]),
        "wind_dir_text": pd["wind_dir"],
        "wind_kt": pd["wind_kt"],
        "wind_gust_kt": pd.get("wind_gust_kt"),
        "sea_ft": pd["sea_ft"],
        "wave_detail": wave_detail,
        "wave_detail_synoptic": wave_detail_synoptic,
        "zone": zone,
        "period": period,
    }
    override = assignment.get("override", {})
    for key in ("wind_dir_text", "wind_kt", "sea_ft"):
        if key in override:
            out[key] = override[key]
    if "wind_dir_deg" in override:
        out["wind_dir_deg"] = _to_int_deg(override["wind_dir_deg"])
    return out


def build_legs_for_plan(passage: dict, forecast: dict, plan_id: str) -> list[Leg]:
    """Resolve every WP into a Leg with all derived values computed."""
    plan = next(p for p in passage["plans"] if p["id"] == plan_id)
    assignments = forecast["waypoint_assignments"][plan_id]
    waypoints = passage["waypoints"]
    calibration = passage["calibration"]
    arrival_timing = passage["arrival_timing"]
    motor_speed = passage["vessel"]["motor_speed_kt"]
    motor_crossover = passage["vessel"]["motor_crossover_tws_kt"]
    polar_design = passage["vessel"].get("design_number", "D1170")

    # Underway mid-route start: when a plan declares `underway_start`, drop the
    # waypoints already passed and begin the leg sequence at the boat's current
    # fix. depart_hour for such a plan is the FIX time (not the original
    # departure), so the remaining legs are timed forward from the fix.
    us = plan.get("underway_start")
    if us:
        # Determine the waypoints AHEAD of the boat. Prefer an explicit
        # next_wp_id (handles a corner-cut, where the boat has stood off the
        # rhumb and bypassed a waypoint that is still "ahead" by cum). Else
        # fall back to the cum_nm filter.
        next_id = us.get("next_wp_id")
        if next_id:
            idx = next((j for j, wp in enumerate(waypoints) if wp["id"] == next_id), None)
            ahead = list(waypoints[idx:]) if idx is not None else []
        else:
            start_cum = float(us["cum_nm"])
            ahead = [wp for wp in waypoints if wp["cum_nm"] > start_cum + 0.05]
        if ahead:
            # First leg distance from ACTUAL fix coords to the next waypoint
            # (not a cum difference — the boat may be well off the rhumb).
            d0 = _haversine_nm(us["lat"], us["lon"], ahead[0]["lat"], ahead[0]["lon"])
            boat_cum = ahead[0]["cum_nm"] - d0
            boat_node = {
                "id": us.get("label", "BOAT"),
                "name": us.get("name", "Current Position"),
                "lat": us["lat"],
                "lon": us["lon"],
                "cum_nm": boat_cum,
                "course_out": us.get("course_out", ahead[0].get("course_out")),
                "chart_label": us.get("name", "Current Position"),
            }
            # The leg departing the boat now is best represented by the
            # conditions forecast for the waypoint it is currently approaching.
            boat_assign = dict(assignments.get(ahead[0]["id"], {}))
            remaining = passage["passage"].get("total_nm", ahead[-1]["cum_nm"]) - boat_cum
            boat_assign["notes_addendum"] = (
                f"CURRENT FIX {us.get('time', '')} — {boat_cum:.0f} NM run, "
                f"{remaining:.0f} NM remaining. SOG {us.get('sog_kt', '?')} kt, "
                f"COG {us.get('cog_deg', us.get('course_out', '?'))}°T. "
                f"{d0:.0f} NM to {ahead[0]['id']}; leg conditions per {ahead[0]['id']} forecast."
            )
            boat_assign["weather_risk"] = (
                "Underway from the current fix; legs ahead timed from this position."
            )
            assignments = dict(assignments)
            assignments[boat_node["id"]] = boat_assign
            waypoints = [boat_node] + ahead

    legs: list[Leg] = []
    cum_time = 0.0
    cum_sailing = 0.0
    cum_motoring = 0.0

    # First pass: compute polar/boat/TWA for each WP using its own conditions.
    # The leg from WP[i] to WP[i+1] uses WP[i]'s wind/sea (the conditions DEPARTING
    # the waypoint, which dominate the upcoming leg). WP6 has no outgoing leg.
    for i, wp in enumerate(waypoints):
        wp_id = wp["id"]
        course_out = wp.get("course_out")
        if course_out is not None:
            course_out = _to_int_deg(course_out)
        assignment = assignments.get(wp_id, {})
        notes_addendum = assignment.get("notes_addendum", "")
        weather_risk = assignment.get("weather_risk", "")

        if course_out is None or wp_id not in assignments:
            # Terminal WP (e.g. WP6). Use the arrival-zone conditions but no leg.
            # Pull the assignment data anyway for wind/sea display at arrival.
            if wp_id in assignments:
                cond = resolve_leg_conditions(forecast, assignments[wp_id])
                wind_text = cond["wind_dir_text"]
                wind_deg = cond["wind_dir_deg"]
                tws_lo, tws_hi = cond["wind_kt"]
                sea_lo, sea_hi = cond["sea_ft"]
                wave = cond["wave_detail"]["primary"]
                period_s = wave["period_s"]
                cycle_label = get_cycle_label_for_office(forecast, forecast["zones"][cond["zone"]]["office"])
                sea_source = f"Wave Detail {forecast['zones'][cond['zone']]['office']} {cond['zone']} {cond['period']} ({cycle_label} cycle)"
            else:
                wind_text = "—"
                wind_deg = 0
                tws_lo = tws_hi = 0
                sea_lo = sea_hi = 0
                period_s = 0
                sea_source = "—"
                wave = None

            eta_str, eta_h = hours_to_eta(cum_time, plan["depart_day"], plan["depart_hour"])
            eta_color_arr = eta_color_for_hour(eta_h, arrival_timing)

            # Risk color: explicit YAML override or auto-compute
            risk_level = assignment.get("risk_level")
            if not risk_level:
                risk_level = auto_risk_level(
                    wind_kt_high=tws_hi, sea_ft_high=sea_hi,
                    eta_color=eta_color_arr, wind_dir_text=wind_text,
                )
            risk_color = RISK_COLORS.get(risk_level, RISK_COLORS["green"])
            row_band = assignment.get("row_band")
            pressure_inhg = assignment.get("pressure")
            pressure_trend = assignment.get("pressure_trend")

            legs.append(Leg(
                wp_id=wp_id, wp_name=wp["name"], cum_nm=wp["cum_nm"],
                course_out=None,
                eta_str=eta_str, eta_decimal_hour=eta_h,
                eta_color=eta_color_arr,
                wind_dir_text=wind_text, wind_dir_deg=wind_deg,
                wind_kt_low=tws_lo, wind_kt_high=tws_hi, wind_kt_avg=avg(tws_lo, tws_hi),
                gust_kt=cond.get("wind_gust_kt") if wp_id in assignments else None,
                sea_ft_low=sea_lo, sea_ft_high=sea_hi, sea_period_s=period_s,
                sea_from_deg=0, sea_from_label="—", sea_source_tag="",
                sea_angle=0, twa=0,
                polar_speed=0, boat_speed=0,
                sail_mode="Inlet entry", is_motor=False,
                aws=0, awa=0,
                cum_time_hr=cum_time, cum_sailing_hr=cum_sailing, cum_motoring_hr=cum_motoring,
                notes=_format_arrival_notes(wp_id, eta_str, arrival_timing, cond if wp_id in assignments else {}, notes_addendum),
                weather_risk=weather_risk,
                sea_source=sea_source,
                risk_color=risk_color,
                row_band=row_band,
                pressure_inhg=pressure_inhg,
                pressure_trend=pressure_trend,
                chart_label=wp.get("chart_label", ""),
            ))
            continue

        cond = resolve_leg_conditions(forecast, assignment)
        wind_text = cond["wind_dir_text"]
        wind_deg = cond["wind_dir_deg"]
        tws_lo, tws_hi = cond["wind_kt"]
        tws_avg = avg(tws_lo, tws_hi)
        sea_lo, sea_hi = cond["sea_ft"]
        sea_avg = avg(sea_lo, sea_hi)
        wave_primary = cond["wave_detail"]["primary"]
        period_s = wave_primary["period_s"]
        sea_from_deg = _to_int_deg(wave_primary["direction_deg"])
        sea_from_label = f"{sea_from_deg:03d}° {wave_primary['direction']}"

        twa = compute_twa(course_out, wind_deg)
        sea_angle = compute_sea_angle(course_out, sea_from_deg)
        sea_pos = sea_position_label(course_out, sea_from_deg)
        pure_polar = polar_speed(tws_avg, twa, polar_design)
        sf, sf_label = select_sea_factor(twa, sea_avg, period_s, calibration["sea_factors"])
        modeled = pure_polar * sf

        # Motor decision
        if twa < MIN_SAILING_TWA:
            # Wind on the nose — boat can't fetch this angle without tacking.
            # VMG when beating drops below motor speed for cruising yachts, so motor.
            is_motor = True
            boat_spd = motor_speed
            sail_mode = f"Motor (wind on nose, TWA {twa}°)"
        elif tws_avg < motor_crossover and modeled < motor_speed:
            is_motor = True
            boat_spd = motor_speed
            sail_mode = "Motor / light sail"
        elif modeled < motor_speed:
            is_motor = True
            boat_spd = motor_speed
            sail_mode = "Motor / light sail"
        else:
            is_motor = False
            boat_spd = modeled
            sail_mode = _sail_mode_for_twa(twa)

        aws_kt, awa_deg = apparent_wind(tws_avg, twa, boat_spd)

        eta_str, eta_h = hours_to_eta(cum_time, plan["depart_day"], plan["depart_hour"])
        eta_color_leg = eta_color_for_hour(eta_h, arrival_timing)

        # Risk color: explicit YAML override or auto-compute
        risk_level = assignment.get("risk_level")
        if not risk_level:
            risk_level = auto_risk_level(
                wind_kt_high=tws_hi, sea_ft_high=sea_hi,
                eta_color=eta_color_leg, wind_dir_text=wind_text,
            )
        risk_color = RISK_COLORS.get(risk_level, RISK_COLORS["green"])
        row_band = assignment.get("row_band")
        pressure_inhg = assignment.get("pressure")
        pressure_trend = assignment.get("pressure_trend")

        cycle_label = get_cycle_label_for_office(forecast, forecast["zones"][cond["zone"]]["office"])
        sea_source = f"Wave Detail {forecast['zones'][cond['zone']]['office']} {cond['zone']} {cond['period']} ({cycle_label} cycle)"
        notes = _format_notes(
            zone=cond["zone"], period=cond["period"],
            cycle_label=cycle_label,
            wind_text=wind_text, tws_lo=tws_lo, tws_hi=tws_hi,
            sea_lo=sea_lo, sea_hi=sea_hi, wave_primary=wave_primary,
            wave_secondary=cond["wave_detail"].get("secondary"),
            twa=twa, polar=pure_polar, sf=sf, boat=boat_spd,
            is_motor=is_motor, addendum=notes_addendum,
            vessel_label=passage["vessel"]["designation"],
        )

        legs.append(Leg(
            wp_id=wp_id, wp_name=wp["name"], cum_nm=wp["cum_nm"],
            course_out=course_out,
            eta_str=eta_str, eta_decimal_hour=eta_h,
            eta_color=eta_color_leg,
            wind_dir_text=wind_text, wind_dir_deg=wind_deg,
            wind_kt_low=tws_lo, wind_kt_high=tws_hi, wind_kt_avg=tws_avg,
            gust_kt=cond.get("wind_gust_kt"),
            sea_ft_low=sea_lo, sea_ft_high=sea_hi, sea_period_s=period_s,
            sea_from_deg=sea_from_deg, sea_from_label=sea_from_label,
            sea_source_tag="WD",
            sea_angle=sea_angle, twa=twa,
            polar_speed=round(pure_polar, 2), boat_speed=round(boat_spd, 2),
            sail_mode=sail_mode, is_motor=is_motor,
            aws=round(aws_kt, 1), awa=round(awa_deg),
            cum_time_hr=cum_time, cum_sailing_hr=cum_sailing, cum_motoring_hr=cum_motoring,
            notes=notes, weather_risk=weather_risk, sea_source=sea_source,
            risk_color=risk_color, row_band=row_band,
            pressure_inhg=pressure_inhg, pressure_trend=pressure_trend,
            sea_position=sea_pos,
            chart_label=wp.get("chart_label", ""),
        ))

        # Advance cumulative for next leg
        if i < len(waypoints) - 1:
            next_wp = waypoints[i + 1]
            leg_nm = next_wp["cum_nm"] - wp["cum_nm"]
            leg_hr = leg_nm / boat_spd
            cum_time += leg_hr
            if is_motor:
                cum_motoring += leg_hr
            else:
                cum_sailing += leg_hr
            legs[-1].leg_distance_nm = leg_nm
            legs[-1].leg_hours = leg_hr

    # Post-process: auto-derive pressure_trend from the pressure sequence.
    # YAML-specified trends win on conflict (drift warning printed); blanks
    # are auto-filled; pressure_trend_auto:true forces derived value.
    derive_pressure_trends(legs, assignments, plan_id=plan_id)

    return legs


def _classify_pressure_rate(rate_inhg_per_hr: float) -> str:
    """Convert a rate of change (inHg per hour) into the trend vocabulary.

    Thresholds calibrated against typical synoptic-scale pressure changes:
      |rate| < 0.002:  Steady             (drift below diurnal noise)
      |rate| < 0.008:  Rising/Falling slow (typical fair-weather change)
      |rate| < 0.020:  Rising/Falling      (active synoptic forcing)
      |rate| >= 0.020: Rising/Falling fast (frontal passage / rapid trough)

    For reference: a deep cold front typically drops ~0.10 inHg over 6 hr
    (~0.017/hr → "Falling"); a vigorous occlusion can do 0.20 over 4 hr
    (~0.050/hr → "Falling fast"); a fair-weather Bermuda High ridge
    rebuild runs ~0.04 inHg per 12 hr (~0.003/hr → "Rising slow").
    """
    abs_r = abs(rate_inhg_per_hr)
    if abs_r < 0.002:
        return "Steady"
    sign = "Rising" if rate_inhg_per_hr > 0 else "Falling"
    if abs_r < 0.008:
        return f"{sign} slow"
    if abs_r < 0.020:
        return sign
    return f"{sign} fast"


def derive_pressure_trends(legs, assignments, plan_id=""):
    """Post-process: derive pressure_trend from the sequence of pressure values.

    Algorithm:
      1. For each leg with a pressure value, compute backward-difference rate
         (inHg/hr) vs the prior leg with pressure. WP1 uses forward diff to WP2.
      2. Classify the rate via _classify_pressure_rate().
      3. Detect Bottoming: if the prior leg was Falling/Falling slow/Falling fast
         AND the current is Steady or Rising slow, override to "Bottoming".
      4. Compare to YAML-specified pressure_trend (from assignments[wp_id]):
         - YAML absent → auto-fill leg.pressure_trend with derived
         - YAML matches derived → silent
         - YAML differs → print drift warning, YAML value retained on leg
         - YAML has pressure_trend_auto: true → force derived (treat as absent)

    Args:
        legs: list[Leg] for one plan, in waypoint order
        assignments: dict from forecast yaml: waypoint_assignments[plan_id]
        plan_id: for warning messages

    Returns: number of drift warnings emitted (for testing/logging).
    """
    # Build (idx, pressure, t_hr) tuples for legs that have pressure data
    samples = []
    for i, leg in enumerate(legs):
        if leg.pressure_inhg is not None:
            samples.append((i, leg.pressure_inhg, leg.cum_time_hr))

    if len(samples) < 2:
        return 0  # Need at least two points to compute any rate

    # Compute derived trend for each sample using backward-difference,
    # with the first sample using forward-difference to the second.
    derived = {}  # leg_idx -> derived_trend_string
    prior_classification = None
    for k, (i, p, t) in enumerate(samples):
        if k == 0:
            # Forward diff to next
            _, p_next, t_next = samples[1]
            dt = t_next - t
            rate = (p_next - p) / dt if dt > 0 else 0.0
        else:
            _, p_prev, t_prev = samples[k - 1]
            dt = t - t_prev
            rate = (p - p_prev) / dt if dt > 0 else 0.0

        trend = _classify_pressure_rate(rate)

        # Bottoming detection: prior was Falling*, current Steady or Rising slow
        if prior_classification and prior_classification.startswith("Falling"):
            if trend == "Steady" or trend == "Rising slow":
                trend = "Bottoming"

        derived[i] = trend
        prior_classification = trend

    # Apply derived values, with YAML precedence and drift warnings
    warnings = 0
    for i, leg in enumerate(legs):
        if i not in derived:
            continue
        yaml_val = assignments.get(leg.wp_id, {}).get("pressure_trend")
        force_auto = assignments.get(leg.wp_id, {}).get("pressure_trend_auto", False)
        derived_val = derived[i]

        if yaml_val is None or force_auto:
            # Auto-fill
            leg.pressure_trend = derived_val
        elif yaml_val == derived_val:
            # Match — silent
            leg.pressure_trend = yaml_val
        else:
            # Drift — warn, YAML wins
            tag = f"[{plan_id}] " if plan_id else ""
            print(f"  ⚠ {tag}pressure_trend drift at {leg.wp_id}: "
                  f"YAML='{yaml_val}' but data implies '{derived_val}'")
            warnings += 1
            leg.pressure_trend = yaml_val

    return warnings


def _sail_mode_for_twa(twa: float) -> str:
    if twa < 60:
        return "Close reach"
    if twa < 80:
        return "Close reach"
    if twa < 100:
        return "Beam reach"
    if twa < 150:
        return "Broad reach"
    return "Running"


def _format_notes(*, zone, period, cycle_label, wind_text, tws_lo, tws_hi,
                  sea_lo, sea_hi, wave_primary, wave_secondary,
                  twa, polar, sf, boat, is_motor, addendum, vessel_label="vessel"):
    """Generate the auto-built Notes string for a WP row."""
    period_human = period.replace("_", " ")
    wave_str = f"{wave_primary['direction']} {wave_primary['height_ft']} ft @ {wave_primary['period_s']}s"
    if wave_secondary:
        wave_str += f" + {wave_secondary['direction']} {wave_secondary['height_ft']} ft @ {wave_secondary['period_s']}s"
    speed_str = (
        f"MOTOR {boat:.1f} kt (wind below crossover or polar < motor)"
        if is_motor
        else f"{vessel_label} polar({tws_lo+(tws_hi-tws_lo)/2:.1f}, TWA {twa}°) {polar:.2f} × SF {sf} = {boat:.2f} kt"
    )
    base = (
        f"FORECAST {zone} {period_human} ({cycle_label} cycle): "
        f"{wind_text} {tws_lo}-{tws_hi} kt, seas {sea_lo}-{sea_hi} ft. "
        f"Wave Detail {wave_str}. "
        f"{speed_str}."
    )
    if addendum:
        addendum_clean = " ".join(addendum.split())
        base += f" {addendum_clean}"
    return base


def _format_arrival_notes(wp_id, eta_str, arrival_timing, cond, addendum):
    """Generate the WP6 (arrival) Notes string."""
    civil_dawn_h = arrival_timing["civil_dawn"]
    sunrise_h = arrival_timing["sunrise"]
    cd_str = _decimal_to_hm(civil_dawn_h)
    sr_str = _decimal_to_hm(sunrise_h)
    base = (
        f"ARRIVAL {eta_str} EDT — civil dawn at {arrival_timing['destination_label']} ~{cd_str}, "
        f"sunrise ~{sr_str}."
    )
    if cond:
        wp_str = (
            f" Arrival zone {cond.get('zone', '?')} {cond.get('period', '?').replace('_', ' ')}: "
            f"{cond.get('wind_dir_text', '?')}, seas {cond.get('sea_ft', ['?', '?'])[0]}-{cond.get('sea_ft', ['?', '?'])[1]} ft."
        )
        base += wp_str
    if addendum:
        addendum_clean = " ".join(addendum.split())
        base += f" {addendum_clean}"
    return base


def _decimal_to_hm(decimal_h: float) -> str:
    h = int(decimal_h)
    m = int(round((decimal_h - h) * 60))
    if m == 60:
        h += 1
        m = 0
    if h == 0:
        h12, ampm = 12, "AM"
    elif h < 12:
        h12, ampm = h, "AM"
    elif h == 12:
        h12, ampm = 12, "PM"
    else:
        h12, ampm = h - 12, "PM"
    return f"{h12}:{m:02d} {ampm}"


# ======================================================================
# Watch Brief — 12-hour tactical dashboard segment derivation
# ======================================================================
@dataclass
class WatchSegment:
    """One 3-hour segment of the 12-hour watch brief.

    Built by walking the polar-driven Leg list and projecting the boat's
    expected conditions, performance, and position at the segment midpoint.
    Each segment maps to one Leg (the leg the boat is sailing TO during the
    segment's time span), so the forecast and polar math come straight from
    the existing build pipeline — no second source of truth.
    """
    idx: int
    start_hr: float             # hours from plan depart
    end_hr: float
    start_clock: str            # "Mon 9:30 AM"
    end_clock: str              # "Mon 12:30 PM"
    label: str                  # "Mon 9:30 AM - 12:30 PM"

    # Day/night context
    is_day: bool                # True if segment is mostly daylight
    day_night_text: str         # "DAY" / "NIGHT" / "DAY → DUSK" / "DAWN → DAY"
    contains_sunrise: bool
    contains_sunset: bool

    # Position context (from active leg)
    active_wp_id: str           # WP the boat is approaching during this segment
    active_wp_name: str
    position_label: str         # "Between WP1 (St Aug) and WP2 (Jacksonville)"

    # Forecast conditions (from active leg)
    course: int
    twa: int
    tws: float
    wind_dir_text: str
    wind_dir_deg: int
    wind_low: float
    wind_high: float
    gust: Optional[float]
    seas_low: float
    seas_high: float
    sea_period: float
    sea_from_deg: int
    sea_from_text: str
    pressure: Optional[float]
    pressure_trend: Optional[str]

    # Performance
    boat_speed_polar: float       # Pure polar Vs at this TWS/TWA
    boat_speed_calibrated: float  # Polar × sea_factor (what the boat will actually do)
    sail_mode: str
    aws: float
    awa: int
    distance_segment_nm: float    # NM covered in segment_hours at calibrated speed

    # Risk + tactical
    risk_color: str             # hex (C6EFCE/FFEB9C/FFC7CE)
    risk_level: str             # green/yellow/red
    action: str                 # 1-2 line tactical advice

    # Convenience flags for action derivation
    is_arrival: bool = False
    notes: str = ""


def _t_to_clock_str(plan_depart_hour: float, depart_day: str, t_hr: float) -> str:
    """Convert hours-from-depart to a 12-hour clock string with day prefix.

    Day cycles forward: Sun → Mon → Tue → Wed (etc). Handles past-midnight
    segments correctly.
    """
    DAY_ORDER = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    try:
        di = DAY_ORDER.index(depart_day)
    except ValueError:
        di = 0
    total_h = plan_depart_hour + t_hr
    day_offset = int(total_h // 24)
    hour_of_day = total_h % 24
    day_label = DAY_ORDER[(di + day_offset) % 7]

    h_int = int(hour_of_day)
    m = int(round((hour_of_day - h_int) * 60))
    if m == 60:
        h_int += 1
        m = 0
    if h_int >= 24:
        h_int -= 24
    if h_int == 0:
        h12, ampm = 12, "AM"
    elif h_int < 12:
        h12, ampm = h_int, "AM"
    elif h_int == 12:
        h12, ampm = 12, "PM"
    else:
        h12, ampm = h_int - 12, "PM"
    return f"{day_label} {h12}:{m:02d} {ampm}"


def _find_active_leg(legs, t_hr):
    """Find the leg the boat is sailing AT time t_hr (hours from depart).

    Each leg's cum_time_hr is the time the boat REACHES that WP. Conditions
    on the leg describe the transit INTO that WP from the previous one.
    So for a segment whose midpoint is at t_hr, the active leg is the first
    leg with cum_time_hr >= t_hr (the one the boat is currently approaching).
    """
    for i, leg in enumerate(legs):
        if leg.cum_time_hr >= t_hr:
            return i, leg
    return len(legs) - 1, legs[-1]


def _classify_day_night(start_hour: float, end_hour: float, day_window):
    """Classify a segment's time span vs daylight window.

    day_window is [sunrise_h, sunset_h] in 24-hour decimal at the destination.
    Returns (is_day, label, contains_sunrise, contains_sunset).

    For midnight-spanning segments, normalizes both ends to [0, 24).
    """
    if day_window is None or len(day_window) != 2:
        return True, "DAY", False, False
    sr, ss = day_window  # sunrise hour, sunset hour
    s = start_hour % 24
    e = end_hour % 24
    contains_sr = False
    contains_ss = False
    if s < e:
        contains_sr = s <= sr <= e
        contains_ss = s <= ss <= e
    else:
        # wraps midnight
        contains_sr = sr >= s or sr <= e
        contains_ss = ss >= s or ss <= e

    def hr_is_day(h):
        return sr <= h <= ss

    s_is_day = hr_is_day(s)
    mid = (start_hour + end_hour) / 2 % 24
    mid_is_day = hr_is_day(mid)
    e_is_day = hr_is_day(e)

    if s_is_day and e_is_day and not contains_sr and not contains_ss:
        return True, "DAY", False, False
    if not s_is_day and not e_is_day and not contains_sr and not contains_ss:
        return False, "NIGHT", False, False
    if contains_sunrise := contains_sr:
        return mid_is_day, "DAWN → DAY", True, False
    if contains_ss:
        return mid_is_day, "DAY → DUSK", False, True
    return mid_is_day, "DAY" if mid_is_day else "NIGHT", contains_sr, contains_ss


def _position_label(active_leg, prev_leg, t_hr):
    """Build a compact position description for the watch card.

    'Approaching St Augustine' if the segment lands close to the active WP.
    'St Augustine → Jacksonville' for mid-leg segments (no WP IDs, just
    short city labels — keeps text short enough for the card width).
    """
    def short(name):
        from .charts import _short_location_from_wp_name
        sl = _short_location_from_wp_name(name)
        return sl or name[:18]

    wp_short = short(active_leg.wp_name)
    if prev_leg is None:
        return f"Approaching {wp_short}"
    span = active_leg.cum_time_hr - prev_leg.cum_time_hr
    if span <= 0:
        frac = 1.0
    else:
        frac = (t_hr - prev_leg.cum_time_hr) / span
    if frac > 0.80:
        return f"Approaching {wp_short}"
    prev_short = short(prev_leg.wp_name)
    return f"{prev_short} → {wp_short}"


def _derive_action(seg: WatchSegment) -> str:
    """Generate a tactical action recommendation for the segment.

    Rule-based heuristic combining wind/gust/sea levels, day/night
    transitions, convective probability, and risk classification.
    """
    actions = []

    # Sun events take precedence — schedule-relevant for watch staffing
    if seg.contains_sunrise:
        actions.append("Sunrise this segment — switch off nav lights, day-watch handover")
    if seg.contains_sunset:
        actions.append("Sunset this segment — switch to nav lights, reef before dark if doubtful")

    # Reefing triggers from gust forecast
    if seg.gust is not None and seg.gust >= 25:
        actions.append(f"Reef BEFORE this segment — gusts to {int(seg.gust)} kt forecast")
    elif seg.gust is not None and seg.gust >= 20:
        actions.append(f"Stand-by reef — gusts to {int(seg.gust)} kt (SCA-threshold territory)")
    elif seg.wind_high >= 18:
        actions.append("Watch for SCA-threshold winds; reef if sustained 18+ kt")

    # Night transit in elevated conditions
    if not seg.is_day and seg.risk_level == "yellow" and seg.gust is None:
        actions.append("Night watch in elevated conditions — keep reefed sail plan")

    # Convective watch from notes
    note_low = (seg.notes or "").lower()
    if any(k in note_low for k in ["tstm", "thunderstorm", "convective", "storm"]):
        actions.append("Watch radar for convective cells")

    # Arrival callout
    if seg.is_arrival:
        actions.append("FINAL APPROACH — verify harbor traffic on VHF, prep dock lines/fenders")

    # Sail-mode hint
    sm_low = (seg.sail_mode or "").lower()
    if "motor" in sm_low and seg.tws < 9:
        actions.append("Light wind — motor-sailing regime; monitor fuel burn")

    if not actions:
        actions.append("Maintain pace; conditions stable for this segment")
    return " · ".join(actions)


def elapsed_hr_at_cum(legs, cum_nm):
    """Interpolate elapsed passage time (hr) at a given cumulative distance (NM).

    Used to anchor the Watch Brief window to the boat's current position when
    underway. Each leg carries cum_nm (distance to that WP) and cum_time_hr
    (time to reach it). Linear interpolation between the bracketing legs.
    Clamps to [0, last leg time].
    """
    if not legs or cum_nm is None:
        return 0.0
    cum_nm = float(cum_nm)
    if cum_nm <= legs[0].cum_nm:
        # Before/at first WP — scale from origin (0,0) to first leg.
        if legs[0].cum_nm <= 0:
            return 0.0
        return max(0.0, cum_nm / legs[0].cum_nm * legs[0].cum_time_hr)
    for i in range(1, len(legs)):
        a, b = legs[i - 1], legs[i]
        if cum_nm <= b.cum_nm:
            span_nm = b.cum_nm - a.cum_nm
            if span_nm <= 0:
                return a.cum_time_hr
            frac = (cum_nm - a.cum_nm) / span_nm
            return a.cum_time_hr + frac * (b.cum_time_hr - a.cum_time_hr)
    return legs[-1].cum_time_hr


def build_watch_segments(plan, legs, arrival_timing, total_nm, segment_hours=3.0, n_segments=4, start_offset_hr=0.0):
    """Derive 4 x 3-hour watch segments starting from the plan's depart_hour.

    Each segment is mapped to the active leg the boat is sailing during the
    segment's time window. Forecast, polar, and tactical fields are pulled
    from the leg with no second source of truth.

    Args:
      plan: passage YAML plan block (must have depart_hour, depart_day)
      legs: list of Leg objects from build_legs_for_plan
      arrival_timing: passage YAML arrival_timing block (for day_window)
      total_nm: total passage distance (for arrival-callout detection)
      start_offset_hr: elapsed hours into the passage at which the 12-hour
        watch window begins. 0 for a pre-departure plan; for an underway
        boat, set to the elapsed time at the current fix so the Watch Brief
        tracks the NEXT 12 hours from the boat's position, not from the
        (possibly back-anchored) departure.

    Returns:
      list of WatchSegment, length n_segments.
    """
    depart_hour = float(plan["depart_hour"])
    depart_day = plan.get("depart_day", "Mon")
    day_window = (arrival_timing or {}).get("day_window")
    last_leg_cum = legs[-1].cum_time_hr if legs else 0.0

    segments = []
    for i in range(n_segments):
        # chart_* are 0-based (the strip-chart x-axis is fixed to 0..12 hr).
        # elapsed_* (start_hr/end_hr/mid_hr) include start_offset_hr so the
        # window, clocks, active-leg lookup and day/night track the boat's
        # current position when underway.
        chart_start = i * segment_hours
        chart_end = (i + 1) * segment_hours
        start_hr = start_offset_hr + chart_start
        end_hr = start_offset_hr + chart_end
        mid_hr = start_hr + segment_hours / 2

        leg_idx, active_leg = _find_active_leg(legs, mid_hr)
        prev_leg = legs[leg_idx - 1] if leg_idx > 0 else None
        # In current absolute clock terms, what's the start/end?
        absolute_start = depart_hour + start_hr
        absolute_end = depart_hour + end_hr

        is_day, dn_text, has_sr, has_ss = _classify_day_night(
            absolute_start, absolute_end, day_window
        )

        # Position fraction along passage (for arrival check)
        is_arrival = mid_hr >= last_leg_cum * 0.92 if last_leg_cum > 0 else False

        # Boat speed: polar potential vs calibrated.
        # active_leg.polar_speed is pure polar Vs; active_leg.boat_speed includes
        # sea_factor calibration. For motoring legs, both equal motor_speed.
        polar_bs = active_leg.polar_speed if active_leg.polar_speed > 0 else active_leg.boat_speed
        calib_bs = active_leg.boat_speed

        # Distance covered in segment_hours at calibrated speed
        distance = calib_bs * segment_hours

        seg = WatchSegment(
            idx=i,
            start_hr=chart_start,
            end_hr=chart_end,
            start_clock=_t_to_clock_str(depart_hour, depart_day, start_hr),
            end_clock=_t_to_clock_str(depart_hour, depart_day, end_hr),
            label="",  # set below
            is_day=is_day,
            day_night_text=dn_text,
            contains_sunrise=has_sr,
            contains_sunset=has_ss,
            active_wp_id=active_leg.wp_id,
            active_wp_name=active_leg.wp_name,
            position_label=_position_label(active_leg, prev_leg, mid_hr),
            course=int(active_leg.course_out) if active_leg.course_out is not None else 0,
            twa=active_leg.twa,
            tws=active_leg.wind_kt_avg,
            wind_dir_text=active_leg.wind_dir_text,
            wind_dir_deg=active_leg.wind_dir_deg,
            wind_low=active_leg.wind_kt_low,
            wind_high=active_leg.wind_kt_high,
            gust=active_leg.gust_kt,
            seas_low=active_leg.sea_ft_low,
            seas_high=active_leg.sea_ft_high,
            sea_period=active_leg.sea_period_s,
            sea_from_deg=active_leg.sea_from_deg,
            sea_from_text=active_leg.sea_from_label,
            pressure=active_leg.pressure_inhg,
            pressure_trend=active_leg.pressure_trend,
            boat_speed_polar=round(polar_bs, 1),
            boat_speed_calibrated=round(calib_bs, 1),
            sail_mode=active_leg.sail_mode,
            aws=active_leg.aws,
            awa=active_leg.awa,
            distance_segment_nm=round(distance, 1),
            risk_color=active_leg.risk_color,
            risk_level={"C6EFCE": "green", "FFEB9C": "yellow", "FFC7CE": "red"}.get(active_leg.risk_color, "green"),
            action="",
            is_arrival=is_arrival,
            notes=active_leg.notes or "",
        )
        seg.label = f"{seg.start_clock} – {seg.end_clock}"
        seg.action = _derive_action(seg)
        segments.append(seg)

    return segments
