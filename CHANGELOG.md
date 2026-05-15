# Sailplan Changelog

Major design decisions and feature additions, in reverse chronological order.

---

## 2026-05-15 — Gust column wiring fix + verifier strengthening

**Problem:** Plan tab Gust column kept showing "—" for all waypoints even when
buoy observations clearly showed gusts (e.g., 41033 sustained 8 kt / gust 12 kt).

**Root cause:** Two layers needed updating:
1. The renderer was correct, but the `Leg` dataclass had no `gust_kt` field.
2. Even with the field added, the forecast YAML schema didn't always include
   `wind_gust_kt` per zone-period — NWS often omits gusts in CWFCHS bulletins
   when conditions are benign.

**Fix:**
- Added `gust_kt: Optional[float]` to `Leg` dataclass.
- `resolve_leg_conditions` now pulls `wind_gust_kt` from forecast zone-period.
- Plan tab renderer writes `g{gust_kt}` or `—`.
- Forecast YAMLs now include `wind_gust_kt` derived from buoy observations
  when NWS omits the field.
- Verifier upgraded with `_check_gust_data_populated` — flags ERROR when YAML
  has gust but cell is "—", and WARN when entire Plan tab gust column is "—"
  (catches "forecaster forgot gusts entirely" pattern).

## 2026-05-15 — Presentation polish

Major styling refactor across every tab. New `styles.py` (271 lines) with
type-aware helpers: `style_page_title`, `style_section_header`,
`style_table_header`, `style_number_cell`, `style_centered_cell`,
`style_text_cell`. Calibri 11pt body / 16pt navy title / 12pt section
headers. Dark blue + white column headers (`305496`). Alternating row bands
on dense tables. Freeze panes on data-heavy tabs. Right-aligned numbers
with appropriate decimal precision. Center-aligned codes (TWA, course,
sea angle). Left-aligned wrapping text for descriptions/notes/risk.

**Tabs polished:** Plan tabs, Briefing, Waypoints, Vessel Particulars,
Live Buoy Data, Forecast Sources by WP, Buoys by WP, Refresh Cadence,
Verification Scorecard, Glossary, URL Quick Reference, Forecast Products.

## 2026-05-15 — Tab reorganization

Reorganized the 16-tab workbook into five logical reading-order groups:
1. The Answer (Pre-Departure Briefing)
2. The Plans (Plan + Bowtie per scenario)
3. The Weather (Live Buoy Data, Forecast Sources, Buoys by WP, Refresh Cadence)
4. Route & Vessel (Waypoints, Vessel Particulars)
5. Static Reference (Format Reference, Forecast Products, URL Quick Reference,
   Glossary, Verification Scorecard)

The skipper reads top-to-bottom: "what should I do?" → "here's the plan" →
"here's the weather behind it" → "here's the route" → "reference if needed."

## 2026-05-15 — Vessel Comparison made opt-in

Most passages use a single vessel; the Vessel Comparison tab only adds value
when weighing one design against another. Default is now OFF. Set
`include_vessel_comparison: true` in passage YAML to enable.

## 2026-05-14 — Systematic verifier (`verify.py`)

New end-of-run verification runs automatically after every build. Seven checks:

1. **Vessel consistency** — flags HR48 references in HR54 workbook (and vice versa)
2. **Polar consistency** — Vessel Particulars subtitle + spot-check polar values match passage design
3. **Buoy coordinates** — workbook lat/lon vs YAML (>0.01° = warn)
4. **Arrival timing** — flags yellow/red ETA cell colors
5. **No HR48 leakage** — canary check for hard-coded HR48 strings
6. **Tab inventory** — count matches expected (12 base + 2×plans + comparison)
7. **Forecast freshness** — Format Reference panel showing STALE/EXPIRED
8. **Gust data populated** (added 2026-05-15) — Plan tab column I

This is the standing procedure that catches regressions before workbook delivery.

## 2026-05-14 — HR 54 polar selection

`compute.py` now reads `passage.vessel.design_number` and threads it to
`polar_speed()`. POLARS dict in `polar.py` keyed by design number (D1170,
D1206) with grids extracted from manufacturer speed tables. Eliminates the
"VS_GRID_D1170 alias" bug that silently used HR48 polar even for HR54
passages.

## 2026-05-14 — Six hard-coded HR48 leaks fixed

All hard-coded "HR 48" strings in renderers parameterized via
`passage.vessel.designation`:
- `briefing.py` exposure gate text
- `compute.py` Notes field "HR 48 polar(...)"
- `risk_bowtie.py` title + subtitle
- `support.py` glossary entries (Motor crossover, Code D)
- `plan.py` footer legend polar reference
- Bowtie hardcoded Brewer CR (42.9) and CSF (1.66) now read from vessel YAML

## 2026-05-14 — Course-up wind/sea rose embedding

`rose.py` generates SVG rose with boat pointing top; N/E/S/W and wind/sea
arrows rotate by -course_deg. Cairosvg required. One rose per waypoint in
Plan tab column U.

## 2026-05-14 — KML/GPX route export

`export.py` produces:
- `{Route name}_Route.kml` — for Google Earth, Google My Maps, OpenCPN
- `{Route name}_Route.gpx` — for chartplotters (waypoints + connected route)

## 2026-05-14 — Multi-design polar support

`polar.py`: `POLARS = {"D1170": HR48_grid, "D1206": HR54_grid}`. HR48 grid
from `/mnt/project/HR48Boatspeeds.xls`; HR54 grid from
`/mnt/project/HR54speedtable.xls` (extracted via xlrd).

## Earlier — Initial methodology

The standing 17-column Plan tab format, color coding (green/yellow/red),
sea-factor calibration, arrival-timing discipline, and verification cadence
were established earlier under the Beast Mode Boating project. The code in
this repository codifies that methodology.
