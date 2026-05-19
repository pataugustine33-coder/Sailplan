# Sailplan Changelog

Major design decisions and feature additions, in reverse chronological order.

---
## 2026-05-19 — Methodology compliance verifier

Added `_check_methodology_compliance()` to `sailbuild/verify.py`. This is
the catch-all "does the workbook deliver what the project methodology
says it should?" check.

The other verify functions cover specific failure modes (vessel
consistency, polar grid match, buoy coord drift, image presence, cell
text scan, geographic zone validation). This one covers the structural
contract:

  A. Mandatory tabs present (12 fixed + 2 per plan + optional Vessel
     Comparison). Each missing tab produces an error naming the tab
     and its purpose, e.g.:
       ❌ Missing mandatory tab 'Verification Scorecard'. Purpose:
       Pre/under/post scorecard for verification methodology.

  B. Plan tabs carry all 17 mandatory columns (A-Q) plus the Sea
     Source standing-instruction column. Accepts a list of header
     alternatives for graceful evolution (e.g. "ETA (EDT)" or "ETA").
     Each missing column produces an error citing the column letter,
     accepted header labels, and the methodology requirement.

  C. Plan tabs have one row per WP defined in the passage YAML.
     Catches truncated builds.

  D. Pressure Trend values use the controlled vocabulary from the
     system prompt: Rising fast / Rising / Rising slow / Steady /
     Falling slow / Falling / Falling fast / Bottoming.

  E. ETAs match the project format: 12-hour clock with day prefix,
     e.g. "Tue 3:00 PM". Catches "2026-05-19 15:00" and "3:00 PM"
     (missing day prefix).

Rule set is declarative — represented as the `_PLAN_TAB_REQUIRED_COLUMNS`
constant and the `_required_tabs(passage)` helper. When the standing
instructions evolve, you add a rule entry rather than write a new
function.

Why this check matters: the project's failure modes that hurt the most
are structural (wrong tabs, wrong columns, wrong format conventions),
not arithmetic. A hand-rolled workbook that bypasses build.py can be
syntactically valid (no Excel errors, no empty cells, formulas all
evaluated) and still violate the methodology because it has 10 tabs
instead of 16, or 17 plan columns instead of 28. Without a structural
check the existing verify functions all pass, build.py exits 0, and
the operator delivers a non-conforming workbook.

Test results on first run:
  - Tue 5/19 Charleston-Beaufort build: clean (the pipeline produces
    a methodology-compliant workbook by construction).
  - Synthetic broken workbook (2 tabs, 3 plan columns): 27 methodology
    errors emitted, each naming exactly what's missing and where.
  - Existing chs-savannah build: no methodology errors (only the
    pre-existing zone/arrival errors).

---
## 2026-05-19 — Plan-tab image-presence verifier

Added `_check_plan_tab_images()` to `sailbuild/verify.py`. For each
plan tab, counts embedded images by anchor column and verifies:
  - Wind/Sea Rose images at column U: must equal number of WPs with
    non-null course_out (i.e., every non-arrival waypoint)
  - Polar @ TWS images at column V: same expected count
  - At least one timeline-strip image anchored elsewhere

Catches the silent-failure class where a critical rendering library
(cairosvg for the rose, matplotlib for the polar) is missing in the
build environment and the renderer falls back to writing "—" into
the cell value. The cell-scan verifier sees "—" as a non-empty
string and moves on; only an image-collection-aware check catches
the gap.

Real example: cairosvg was not installed in the environment used for
the Tue 5/19 Charleston-Beaufort build cycle. 0/8 wind/sea roses
rendered on each plan tab. Cell-scan verifier missed it. Skipper
caught it visually. Now the verifier emits:

  ❌ [ERROR] Plan A - Tue 3 PM Depart:U (Wind/Sea Rose column):
  Only 0/8 wind/sea rose images embedded. Likely cause: cairosvg
  not installed... Run: pip install cairosvg

The error message names the most-likely cause and the fix, so a
future build that hits this in CI gets the right next step inline.

---
## 2026-05-19 — Geographic WP-to-zone validator + workbook cell scan

Added two new checks to `sailbuild/verify.py` that run on every build:

**1. `_check_wp_zone_geography(passage, forecast)`** — for every WP in
every plan, verifies that the WP's lat/lon actually falls inside the
forecast zone it's been assigned to. Catches the silent-corruption
class where a coastal-zone bulletin drives the polar model at a
waypoint that's actually 20-60 NM offshore (different forecast, wrong
data, plan model is materially wrong but the pipeline runs clean).

Implementation: `ZONE_REGISTRY` table mapping each NWS marine zone ID
to its lat band and (min_nm, max_nm) distance band offshore, plus a
list of coastline reference points along the zone's coverage. For
each WP, Haversine to nearest reference, then check both bands. ±2
NM tolerance at distance-band edges, ±0.1° tolerance at lat-band
edges. Unknown zones produce WARN (not ERROR) so unrecognized zones
don't block a build — they signal that the registry needs an entry.

Runtime check uncovered three pre-existing real bugs:
  - chs-savannah: WP1 32.75°N assigned to AMZ362 (covers 32.0-32.6°N)
  - chs-staug: WP1/WP2/WP3 in same lat-band error + ≥25 NM offshore
    while assigned to 0-20 NM zones

Initial ZONE_REGISTRY covers KCHS (AMZ340/360/362/364/380/382/384),
KILM (AMZ250/252/254/256/280/284), and KMHX (AMZ150/152/154/156/158/188).
JAX/MFL/MLB/KEY zones to be added when those routes get their next
build cycle.

**2. `_check_workbook_cells(wb, forecast)`** — walks every populated
cell on every tab and flags:
  - Excel error literals (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?,
    #NUM!, #NULL!, #GETTING_DATA) — ERROR severity.
  - Formulas that LibreOffice didn't evaluate (value starts with "=")
    — ERROR severity.
  - Weather-Risk cells whose text starts with "GREEN"/"YELLOW"/"RED"
    color words — violates project color-only standard — ERROR.
  - Placeholder markers (TBC/TODO/FIXME/XXX/TBD/PLACEHOLDER) — WARN.
  - Date references (M/D format with day-of-week prefix) that don't
    match the build's cycle dates — WARN. Skips narrative/lessons-
    learned cells that legitimately cite past cycles as examples.

Why it matters: the existing verifier catches structural problems
(vessel consistency, polar grid mismatch, buoy coord drift, arrival
timing). It did NOT catch syntactic problems sitting inside cells.
The cell scan covers that gap so the project's "color-only risk",
"no stale dates", and "no placeholder text" standards are enforced
at build time rather than relying on the operator to remember.

Both checks wired into `verify_workbook()`; both contribute to the
errors-count return code so a CI/--require-style block-on-error
gate would catch them.

---
## 2026-05-19 — Chrome-first weather pull becomes stop-and-ask SOP

Added explicit standing rule: Claude must attempt the live Chrome fetch
FIRST for every weather source on every passage build. If Chrome is
unavailable, or any source fails to fetch live, Claude STOPS and asks
permission before substituting older / search-snippet / cached data.
Silent fallback to stale data is no longer permitted.

Rationale: the project's whole verifier discipline (issuance times in
the data-freshness panel, "is this forecast or modeled by you" audits,
verification scorecard) only works if the analyst can trust at a glance
that bulletins named in the workbook header reflect the bulletins
actually used to drive the plan. A silent fallback to a Thursday
bulletin pulled via search snippet — when the build runs Tuesday and
ought to use Tuesday's 5 AM cycle — corrupts that trust. Better to
pause and let the skipper decide whether to (a) paste fresh text, (b)
fix Chrome, or (c) explicitly authorize an older-data build.

Documented in HOW_TO_USE.md §3 (Day-of-passage workflow) and added to
the standing rules list as rule #8.

---
## 2026-05-15 — Mini polar charts per WP

Per-waypoint boat polar curves now embedded in a new Plan tab column V
("Polar @ TWS"), positioned right next to the wind/sea rose. Each polar
shows the vessel's speed curve at the WP's mid-range TWS with a red dot
marking the actual TWA at that leg.

Visual answer to "are we at the speed peak or speed valley for this leg?"
At a glance you see:
- Where you sit on the polar (red dot position)
- The curve's shape (where peak speed lives at this TWS)
- Whether nudging the angle 10-20° would gain or lose speed

Uses the same `polar_speed(tws, twa, design)` function that drives the
ETA calculation, so the chart matches the workbook's computed values.
Design number (D1170 / D1206) flows through from passage YAML so HR48
and HR54 passages each render their correct polar shape.

Plan tab now has 28 columns (A through AB). Notes/Risk/Sea Source/Cum
columns shifted right by one to accommodate the new visual column.
Frozen panes remain at D4.

---

## 2026-05-15 — Visualization layer (matplotlib charts)

Picture-is-worth-a-thousand-words pass. Added three chart types via a new
`sailbuild/charts.py` module using matplotlib + PNG embedding (same pattern
as the wind/sea rose):

### Risk Bowtie Radar Chart
The six-axis Risk Bowtie scores (Stability, Motion Comfort, Capsize
Resistance, Speed Margin, Heel Margin, Range) now render as a polar/spider
chart at the top of each Bowtie tab. Background zones color-code the 1-10
scale (red <4, yellow 4-7, green 7+). Score annotations at each vertex.
The polygon shape tells you the vessel's profile at a glance: balanced =
good all-rounder, spiky = strengths/weaknesses, small = concerning.

### Wind/Sea Timeline Strip
Horizontal chart spanning the passage width, embedded between the totals
row and legend block on each Plan tab. Wind speed as vertical bars (sustained)
with gust marks where present, sea height as a line on a secondary y-axis.
SCA threshold (18 kt) and reef threshold (25 kt) drawn as dashed reference
lines. Shows the shape of the passage at a glance — where the rough patches
are without reading individual WP rows.

### Daylight-Banded KML Route
The exported KML route is now split into per-leg LineString segments, each
colored by the destination WP's ETA window:
- Green = day arrival at that WP
- Yellow/amber = twilight arrival
- Red = night arrival

Open in Google Earth and you see the daylight bands as a visual property
of the route itself. The primary plan (first defined in passage YAML) drives
the banding; the GPX file remains plain for chartplotter compatibility.

### Chart Module
New `sailbuild/charts.py` exposes:
- `radar_chart_png_bytes(axes_data)` — six-axis radar
- `radar_overlay_png_bytes(plans_data)` — multi-plan radar overlay (for
  future Plan A vs Plan B comparison on the Pre-Departure Briefing)
- `timeline_strip_png_bytes(legs)` — wind/sea timeline
- `mini_polar_png_bytes(tws, twa, design)` — small polar with TWA marker
  (for future embedding next to per-WP rose)

All chart generators return BytesIO containing PNG bytes for embedding via
openpyxl.drawing.image.Image.

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
