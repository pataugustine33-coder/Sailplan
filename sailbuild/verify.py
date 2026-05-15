"""
verify.py — systematic end-of-run cell-by-cell verification for sailbuild workbooks.

Runs after every build to catch the kinds of bugs that real navigators notice
but a passing test suite wouldn't flag: hard-coded vessel references that leak
into the wrong workbook, polar grid mismatches between Vessel Particulars and
compute.py, buoy coordinate mismatches against NDBC, wrong-vessel metrics in
the Risk Bowtie, and arrival timing that's slipped into the night window.

Each check returns a Finding(severity, location, message).
severity: 'error' (build is wrong, fix before delivery)
          'warn'  (likely correct but worth a manual look)
          'info'  (note for the skipper)

Usage from build.py:
    from sailbuild.verify import verify_workbook
    findings = verify_workbook(output_path, passage, forecast, buoys)
    for f in findings:
        print(f.format())
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Optional
from openpyxl import load_workbook


@dataclass
class Finding:
    severity: str  # 'error' | 'warn' | 'info'
    location: str  # 'tab:cell' or 'YAML:path'
    message: str

    def format(self) -> str:
        sev_icon = {"error": "❌", "warn": "⚠️ ", "info": "ℹ️ "}.get(self.severity, "  ")
        return f"  {sev_icon} [{self.severity.upper()}] {self.location}: {self.message}"


def verify_workbook(output_path: str, passage: dict, forecast: dict, buoys: dict) -> list[Finding]:
    """Run all verification checks against a built workbook.

    Returns a list of findings ordered by severity (errors first).
    Empty list means clean build.
    """
    wb = load_workbook(output_path)
    findings: list[Finding] = []

    findings.extend(_check_vessel_consistency(wb, passage))
    findings.extend(_check_polar_consistency(wb, passage))
    findings.extend(_check_buoy_coords_against_yaml(wb, buoys))
    findings.extend(_check_arrival_timing(wb, passage))
    findings.extend(_check_no_hr48_leakage(wb, passage))
    findings.extend(_check_tab_inventory(wb, passage))
    findings.extend(_check_forecast_freshness(wb, forecast))
    findings.extend(_check_gust_data_populated(wb, passage, forecast))

    # Sort: errors first, then warnings, then info
    severity_order = {"error": 0, "warn": 1, "info": 2}
    findings.sort(key=lambda f: severity_order.get(f.severity, 99))
    return findings


# ==============================================================
# Individual checks
# ==============================================================

def _check_vessel_consistency(wb, passage: dict) -> list[Finding]:
    """The vessel designation in passage YAML should appear in tab titles
    that reference the vessel, and the opposite designation should NOT.

    Catches: HR 48 references that leak into HR 54 workbooks (and vice versa).
    """
    findings = []
    vessel_label = passage.get("vessel", {}).get("designation", "")
    if not vessel_label:
        return findings

    # Determine the "other" common designation that should NOT appear in this
    # workbook unless the Vessel Comparison tab is intentionally included.
    other_designations = []
    if "48" in vessel_label:
        other_designations = ["HR 54", "HR54"]
    elif "54" in vessel_label:
        other_designations = ["HR 48", "HR48"]

    # Skip the Vessel Comparison tab — it's allowed to mention both.
    skip_tabs = {"Vessel Comparison"}

    for tab in wb.sheetnames:
        if tab in skip_tabs:
            continue
        ws = wb[tab]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None or not isinstance(cell.value, str):
                    continue
                for other in other_designations:
                    if other in cell.value:
                        findings.append(Finding(
                            "error",
                            f"{tab}:{cell.coordinate}",
                            f"Cell mentions '{other}' but this passage is '{vessel_label}'. "
                            f"Cell content: {cell.value[:80]!r}",
                        ))
    return findings


def _check_polar_consistency(wb, passage: dict) -> list[Finding]:
    """Vessel Particulars polar grid header should match the passage's design_number.

    Catches: VS_GRID_D1170 alias bug where Vessel Particulars showed HR48
    values even when the passage was HR54.
    """
    findings = []
    design = passage.get("vessel", {}).get("design_number", "")
    if not design or "Vessel Particulars" not in wb.sheetnames:
        return findings

    ws = wb["Vessel Particulars"]
    # Polar subtitle row varies by renderer version — search rows 3-7 for
    # any cell that mentions "Polar" or "TWA × TWS".
    subtitle_found = False
    polar_subtitle_row = None
    for r in range(3, 8):
        v = str(ws.cell(r, 1).value or "")
        if "Polar" in v and ("TWA" in v or "TWS" in v):
            subtitle_found = True
            polar_subtitle_row = r
            if design not in v:
                findings.append(Finding(
                    "error",
                    f"Vessel Particulars:A{r}",
                    f"Polar grid subtitle does not include design_number '{design}'. "
                    f"Got: {v[:80]!r}",
                ))
            break
    if not subtitle_found:
        findings.append(Finding(
            "warn",
            "Vessel Particulars:A3-A7",
            "Could not locate polar grid subtitle in rows 3-7.",
        ))

    # Spot-check polar values — find the TWS header row first (the row that
    # starts with "TWA" in column A and has "TWS" in adjacent columns).
    tws_header_row = None
    for r in range(3, 12):
        v = str(ws.cell(r, 1).value or "")
        if v.startswith("TWA"):
            # Check that col 2 has "TWS"
            if "TWS" in str(ws.cell(r, 2).value or ""):
                tws_header_row = r
                break
    if tws_header_row is None:
        return findings

    # TWA grid is [45, 52, 60, 70, 80, 90, ...] — find the TWA=90 row
    # and column TWS=10 (4th TWS column).
    expected = {"D1170": 7.93, "D1206": 8.23}.get(design)
    if expected is None:
        return findings
    for offset in range(1, 12):
        twa_val = ws.cell(tws_header_row + offset, 1).value
        if twa_val == 90 or twa_val == 90.0:
            # TWS columns start at col 2 with TWS=4,6,8,10 → TWS=10 is col 5
            val = ws.cell(tws_header_row + offset, 5).value
            if val is None:
                break
            if abs(float(val) - expected) > 0.05:
                findings.append(Finding(
                    "error",
                    f"Vessel Particulars:E{tws_header_row + offset}",
                    f"Polar value at TWA 90/TWS 10 is {val}, expected ~{expected} for {design}. "
                    f"Wrong polar grid likely embedded.",
                ))
            break

    return findings


def _check_buoy_coords_against_yaml(wb, buoys: dict) -> list[Finding]:
    """Buoy lat/lon shown in Live Buoy Data should match what's in the YAML.

    Note: This is a self-consistency check (YAML→workbook). For NDBC-truth
    verification, see verify_buoys_against_ndbc() which is a separate
    Chrome-fetched check (called separately when desired).
    """
    findings = []
    if "Live Buoy Data" not in wb.sheetnames or "stations" not in buoys:
        return findings

    # Build a quick id→(lat,lon) map from YAML
    yaml_coords = {
        str(st["id"]): (st.get("lat"), st.get("lon"))
        for st in buoys["stations"]
    }

    # Live Buoy Data renders station id in col 1 area; lat/lon are not always shown
    # in this tab — but if they are, they should match. This check is best-effort.
    # The more reliable check is the Buoys by WP tab.
    if "Buoys by WP" in wb.sheetnames:
        ws = wb["Buoys by WP"]
        for r in range(4, ws.max_row + 1):
            buoy_id = ws.cell(r, 2).value
            lat_str = ws.cell(r, 4).value
            lon_str = ws.cell(r, 5).value
            if buoy_id is None or lat_str is None or lon_str is None:
                continue
            buoy_id = str(buoy_id)
            try:
                lat_val = float(str(lat_str).rstrip("°N").rstrip("°S").rstrip())
                lon_val = float(str(lon_str).rstrip("°W").rstrip("°E").rstrip())
            except (ValueError, AttributeError):
                continue
            yaml_lat, yaml_lon = yaml_coords.get(buoy_id, (None, None))
            if yaml_lat is None:
                # Buoy in workbook but not in buoys YAML — note for cross-check
                continue
            if abs(lat_val - yaml_lat) > 0.01:
                findings.append(Finding(
                    "warn",
                    f"Buoys by WP:D{r}",
                    f"Buoy {buoy_id} lat {lat_val} differs from YAML {yaml_lat} by "
                    f">0.01° (~0.6 NM). Verify against NDBC station page.",
                ))
            if abs(abs(lon_val) - abs(yaml_lon)) > 0.01:
                findings.append(Finding(
                    "warn",
                    f"Buoys by WP:E{r}",
                    f"Buoy {buoy_id} lon {lon_val} differs from YAML {yaml_lon} by "
                    f">0.01° (~0.6 NM). Verify against NDBC station page.",
                ))
    return findings


def _check_arrival_timing(wb, passage: dict) -> list[Finding]:
    """Arrival ETA color should match daylight window in passage.arrival_timing.

    Catches: yellow/red arrival not flagged in the briefing.

    Plan tabs have an ETA in column D for each WP row, but ALSO have a
    position-legend section further down that uses column D for descriptive
    text ("Starboard Bow — ..."). We must stop at the "PASSAGE TOTALS" row
    which marks the end of the WP data block.
    """
    findings = []
    arrival_timing = passage.get("arrival_timing", {})
    if not arrival_timing:
        return findings

    for tab_name in wb.sheetnames:
        if not tab_name.startswith("Plan ") or tab_name.endswith("Bowtie"):
            continue
        ws = wb[tab_name]
        # Find the totals marker (col B says "PASSAGE TOTALS") to bound the WP block.
        totals_row = None
        for r in range(4, ws.max_row + 1):
            b_val = ws.cell(r, 2).value
            if b_val and "PASSAGE TOTAL" in str(b_val).upper():
                totals_row = r
                break

        # Final WP arrival is the LAST row before totals where col D has a value
        # AND the value looks like a time (contains "AM" or "PM").
        upper_bound = totals_row if totals_row else ws.max_row + 1
        last_eta_row = None
        for r in range(4, upper_bound):
            d_val = ws.cell(r, 4).value
            if d_val and ("AM" in str(d_val) or "PM" in str(d_val)):
                last_eta_row = r

        if last_eta_row is None:
            continue
        eta_cell = ws.cell(last_eta_row, 4)
        try:
            fill_color = eta_cell.fill.start_color.value or ""
        except AttributeError:
            fill_color = ""
        eta_text = ws.cell(last_eta_row, 4).value
        if "FFC7CE" in str(fill_color).upper():
            findings.append(Finding(
                "error",
                f"{tab_name}:D{last_eta_row}",
                f"Arrival ETA {eta_text!r} is in NIGHT window (red). "
                f"Consider departure shift using reverse calculator on Format Reference tab.",
            ))
        elif "FFEB9C" in str(fill_color).upper():
            findings.append(Finding(
                "warn",
                f"{tab_name}:D{last_eta_row}",
                f"Arrival ETA {eta_text!r} is in TWILIGHT window (yellow). Marginal — "
                f"verify daylight margin and channel-marker lighting at destination.",
            ))
    return findings


def _check_no_hr48_leakage(wb, passage: dict) -> list[Finding]:
    """Scan every cell for known hard-coded HR48 phrases that should have been
    parameterized. This is the canary check that flags any regression.
    """
    findings = []
    vessel_label = passage.get("vessel", {}).get("designation", "")
    if "48" in vessel_label:
        return []  # this IS HR 48, so HR 48 mentions are correct

    # Known leak phrases that historically appeared hard-coded in renderers.
    # These should ALL be parameterized via vessel.designation now.
    canary_phrases = [
        "manageable for HR 48",
        "HR 48 Comfort & Safety Radar",
        "HR 48 Mk II against this plan",
        "HR 48 polar(",
    ]

    skip_tabs = {"Vessel Comparison"}
    for tab in wb.sheetnames:
        if tab in skip_tabs:
            continue
        ws = wb[tab]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None or not isinstance(cell.value, str):
                    continue
                for phrase in canary_phrases:
                    if phrase in cell.value:
                        findings.append(Finding(
                            "error",
                            f"{tab}:{cell.coordinate}",
                            f"Hard-coded HR 48 phrase leaked into a {vessel_label} workbook: "
                            f"{phrase!r}. Renderer is not parameterized.",
                        ))
    return findings


def _check_tab_inventory(wb, passage: dict) -> list[Finding]:
    """Tab count matches expected based on number of plans and Vessel Comparison flag."""
    findings = []
    n_plans = len(passage.get("plans", []))
    include_comparison = passage.get("include_vessel_comparison", False)

    # Base tabs: Briefing, Format Reference, Waypoints, Vessel Particulars,
    # Live Buoy Data, Refresh Cadence, Forecast Sources by WP, Buoys by WP,
    # Forecast Products, URL Quick Reference, Glossary, Verification Scorecard
    # = 12 fixed
    expected = 12
    expected += 2 * n_plans  # Plan + Bowtie per plan
    if include_comparison:
        expected += 1

    actual = len(wb.sheetnames)
    if actual != expected:
        findings.append(Finding(
            "warn",
            "workbook:tabs",
            f"Tab count is {actual}, expected {expected} "
            f"(12 base + {2*n_plans} plan/bowtie pairs"
            f"{' + 1 Vessel Comparison' if include_comparison else ''}).",
        ))
    return findings


def _check_gust_data_populated(wb, passage: dict, forecast: dict) -> list[Finding]:
    """Plan tabs: column I (Gust kt) must not be "—" for WPs whose forecast
    zone-period has `wind_gust_kt` defined in the YAML.

    Catches: regression where the column is rendered but the value isn't
    written, leaving the navigator looking at a forecast without gust data
    even though the YAML had it. Gusts are decision-relevant for reefing.
    """
    findings = []
    zones = forecast.get("zones", {})
    assignments_by_plan = forecast.get("waypoint_assignments", {})

    # Build tab_label → plan_id map so we check each Plan tab against ITS
    # plan's assignments, not all plans'.
    tab_to_plan_id = {}
    for plan in passage.get("plans", []):
        label = plan.get("tab_label", "")
        # Tab name may be truncated to 31 chars (openpyxl limit); match by prefix
        tab_to_plan_id[label[:31]] = plan["id"]

    for tab_name in wb.sheetnames:
        if not tab_name.startswith("Plan ") or tab_name.endswith("Bowtie"):
            continue
        plan_id = tab_to_plan_id.get(tab_name)
        if plan_id is None:
            continue  # not a plan tab we can map (defensive)
        plan_assignments = assignments_by_plan.get(plan_id, {})
        ws = wb[tab_name]

        # Walk WP rows (start at row 4, stop at PASSAGE TOTALS marker).
        totals_row = None
        for r in range(4, ws.max_row + 1):
            b_val = ws.cell(r, 2).value
            if b_val and "PASSAGE TOTAL" in str(b_val).upper():
                totals_row = r
                break
        upper = totals_row if totals_row else ws.max_row + 1

        for r in range(4, upper):
            wp_id = ws.cell(r, 1).value
            gust_cell = ws.cell(r, 9).value
            if not wp_id or wp_id == "WP":
                continue
            if str(wp_id) not in plan_assignments:
                continue
            assignment = plan_assignments[str(wp_id)]
            zone = assignment.get("zone")
            period = assignment.get("period")
            if not zone or not period:
                continue
            period_data = zones.get(zone, {}).get("periods", {}).get(period, {})
            yaml_gust = period_data.get("wind_gust_kt")
            if yaml_gust and (gust_cell == "—" or gust_cell is None):
                findings.append(Finding(
                    "error",
                    f"{tab_name}:I{r}",
                    f"Gust kt cell is empty/dash but YAML "
                    f"({zone} {period} for {plan_id}) specifies wind_gust_kt={yaml_gust}. "
                    f"Renderer is dropping gust data on the floor.",
                ))

        # Second check: if NONE of the WPs in this plan have gust data populated,
        # warn — the forecast YAML likely doesn't have wind_gust_kt for any of
        # the zone-periods used. NWS often omits gusts when conditions are benign,
        # but for operational use we usually want at least an estimated gust ceiling.
        gust_values_found = 0
        for r in range(4, upper):
            wp_id = ws.cell(r, 1).value
            gust_cell = ws.cell(r, 9).value
            if not wp_id or wp_id == "WP":
                continue
            if gust_cell and gust_cell != "—" and gust_cell is not None:
                gust_values_found += 1
        if gust_values_found == 0:
            findings.append(Finding(
                "warn",
                f"{tab_name}:I (column)",
                f"No gust data populated for any WP in this plan. Forecast YAML "
                f"may be missing wind_gust_kt for the zone-periods used. "
                f"Consider adding gust estimates from buoy observations.",
            ))
    return findings


def _check_forecast_freshness(wb, forecast: dict) -> list[Finding]:
    """Format Reference freshness panel should not show any sources as STALE."""
    findings = []
    if "Format Reference" not in wb.sheetnames:
        return findings
    ws = wb["Format Reference"]
    # Status column is column 5 (header row 4). Data rows start at 5.
    for r in range(5, 15):
        status = ws.cell(r, 5).value
        if status is None:
            break
        status_str = str(status).upper()
        if "STALE" in status_str or "EXPIRED" in status_str:
            source = ws.cell(r, 1).value
            findings.append(Finding(
                "warn",
                f"Format Reference:E{r}",
                f"Source {source!r} is {status_str}. Pull fresh cycle before departure.",
            ))
        elif "OFFLINE" in status_str:
            source = ws.cell(r, 1).value
            findings.append(Finding(
                "info",
                f"Format Reference:E{r}",
                f"Source {source!r} is OFFLINE. Cross-check with neighboring buoys.",
            ))
    return findings


# ==============================================================
# Convenience runner with formatted console output
# ==============================================================

def print_verification_report(output_path: str, passage: dict, forecast: dict, buoys: dict) -> int:
    """Print the verification report to stdout. Returns 0 if clean, N if findings."""
    print("\n" + "=" * 70)
    print("END-OF-RUN VERIFICATION")
    print("=" * 70)
    findings = verify_workbook(output_path, passage, forecast, buoys)
    if not findings:
        print("  ✓ All checks passed. Workbook is internally consistent.")
        return 0
    errors = sum(1 for f in findings if f.severity == "error")
    warns = sum(1 for f in findings if f.severity == "warn")
    infos = sum(1 for f in findings if f.severity == "info")
    print(f"  Findings: {errors} error(s), {warns} warning(s), {infos} info")
    for f in findings:
        print(f.format())
    return errors
