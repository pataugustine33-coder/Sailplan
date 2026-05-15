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
from .polar import polar_speed, select_sea_factor, apparent_wind


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
