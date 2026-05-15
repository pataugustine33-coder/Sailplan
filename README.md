# Sailplan — Beast Mode Boating Passage Planning System

A complete sailing passage planning toolkit for offshore cruising on the
Hallberg-Rassy 48 Mk II and HR 54. Produces a 16-tab Excel workbook with
weather forecast, plan tabs, course-up wind/sea roses, and verification
scorecard — plus KML and GPX files for Google Earth and chartplotters.

This is the codified version of the methodology developed under the
**Beast Mode Boating** media project. It is intended to be used through
Claude (an AI assistant) — you don't need to know how to program to run it.

---

## What this does

Given:
- A route (origin, destination, waypoints)
- A vessel (HR 48 or HR 54)
- A weather forecast pulled from NWS (CWFCHS, AFDCHS, NDBC buoys)

It produces:
- **Sail Plan and Weather Risk - {Origin}-{Destination}.xlsx** — 16-tab workbook
  with everything needed for a GO/NO-GO departure decision and ongoing
  underway verification.
- **Route KML** — for Google Earth, Google My Maps, OpenCPN
- **Route GPX** — for chartplotters (Garmin, B&G, Raymarine, Navionics)

The standard methodology (column format, color coding, arrival-timing
discipline, verification cadence) is documented in the
**Beast Mode Boating** Claude project system prompt.

---

## How to use it (no programming required)

1. Start a new Claude conversation in the Beast Mode Boating project.
2. Tell Claude the passage you want to plan (e.g., "Charleston to Beaufort
   NC, departing Sunday 2 PM").
3. Claude will:
   - Pull fresh weather from NWS via Chrome
   - Update the input YAML files in this codebase
   - Build the workbook by running `build.py`
   - Run the verification check
   - Hand you the .xlsx, .kml, and .gpx files

You never touch the code yourself. Claude does the work; this repository
is the toolkit it uses.

---

## What's in this repo

### Top-level files
- `build.py` — The main script that builds a workbook from inputs.
- `weather_pull.py` — Fetches weather data from NWS/NDBC (sometimes used
  for batch fresh pulls).
- `parse.py` — Parses raw NWS bulletin text into structured YAML.

### `inputs/` — passage definitions and weather data
- `passages/` — One YAML file per route (e.g., `chs-savannah.yaml`).
  Defines the route, vessel, waypoints, calibration, departure scenarios.
- `forecasts/` — One YAML file per forecast cycle pulled from NWS.
- `buoys/` — One YAML file per buoy data pull from NDBC.
- `lessons.yaml` — Methodology lessons captured over time (shown in the
  Verification Scorecard tab).

### `sailbuild/` — the workbook builder code
- `build.py` — main builder
- `compute.py` — turns weather + vessel into leg-by-leg sail plan
- `polar.py` — HR 48 (D1170) and HR 54 (D1206) Frers VPP speed polars
- `rose.py` — generates course-up wind/sea compass roses
- `export.py` — writes KML and GPX route files
- `verify.py` — systematic end-of-run cell-by-cell verification
- `styles.py` — shared cell styling (colors, fonts, alignment)
- `tabs/` — one file per tab type (Plan, Risk Bowtie, Briefing, etc.)
- `parsers/` — read raw NWS bulletin text and NDBC buoy feeds

### `scripts/` — utility scripts (rarely needed)

---

## The 16 tabs in the workbook

Organized into five logical groups so you read top to bottom:

**1. The Answer**
1. Pre-Departure Briefing — single-screen GO/NO-GO dashboard

**2. The Plans**
2. Plan A - {departure scenario} — primary deliverable with WP-by-WP details
3. Plan A Bowtie — six-axis comfort/safety scorecard for Plan A
4. Plan B (optional) — alternative departure scenario
5. Plan B Bowtie

**3. The Weather**
6. Live Buoy Data — current NDBC observations
7. Forecast Sources by WP — which NWS zone covers each waypoint
8. Buoys by WP — which buoy to cross-check at each waypoint
9. Refresh Cadence — when each NWS product issues

**4. Route & Vessel**
10. Waypoints — route geographic detail + paste-ready CSV
11. Vessel Particulars — polar grid + stability metrics
12. Vessel Comparison (optional) — HR 48 vs HR 54 side-by-side

**5. Static Reference**
13. Format Reference — workbook conventions, sea factors, freshness panel
14. Forecast Products — what each NWS product is and when it issues
15. URL Quick Reference — bookmark-ready URLs for this passage
16. Glossary — definitions of TWA, AWS, Wave Detail, etc.
17. Verification Scorecard — methodology lessons + post-passage tracking

(Vessel Comparison is opt-in via `include_vessel_comparison: true` in the
passage YAML. Most workbooks have 16 tabs.)

---

## Standing procedures

These are the rules the system follows on every passage. They live in the
Claude project system prompt and are enforced by the code:

1. **Pull fresh weather via Chrome at session start.** The system never
   uses cached weather. Every passage build pulls the latest CWFCHS, AFDCHS,
   and relevant NDBC buoys.

2. **End-of-run cell-by-cell verification.** After every build, the
   verifier runs seven checks: vessel consistency, polar consistency,
   buoy coordinates, arrival timing, no HR48-leakage-into-HR54-workbook,
   tab inventory, forecast freshness, and gust column populated. Errors
   stop delivery; warnings flag for review.

3. **Arrival timing is a planning constraint, not a result.** The system
   color-codes arrivals: green = full daylight, yellow = civil twilight,
   red = night. Red arrivals trigger a recommendation to shift departure.

4. **Single-vessel format is the default.** Vessel Comparison is opt-in
   per passage.

5. **Color-only weather risk.** No "GREEN/YELLOW/RED" label words — the
   cell color carries the meaning, the cell text gives the reason.

---

## Vessel specifications

### HR 48 Mk II (D1170, half-load)
- Displacement: 40,785 lb
- Motor speed: 6.0 kt
- Motor crossover TWS: 7 kt (below this, motor is faster than sailing)
- Fuel: 110 gal × 1.5 gph at 6 kt = ~440 NM motor range
- Brewer Comfort Ratio: 42.9 (moderate cruiser)
- Capsize Screening Formula: 1.66 (acceptable offshore)

### HR 54 (D1206, half-load)
- Displacement: 49,600 lb
- Motor speed: 6.5 kt
- Motor crossover TWS: 7 kt
- Fuel: 238 gal × 1.8 gph = ~859 NM motor range
- Brewer Comfort Ratio: 52.1 (heavy cruiser)
- Capsize Screening Formula: 1.55 (better than HR 48 offshore)

---

## Glossary of project terms

- **Passage YAML** — the file defining route, vessel, plans for a specific trip
- **Forecast YAML** — the structured weather data from one NWS issuance cycle
- **Buoys YAML** — the structured NDBC buoy observations from one pull
- **Plan tab** — one tab per departure scenario; the heart of the workbook
- **Risk Bowtie** — six-axis comfort/safety scorecard derived from vessel + conditions
- **Verifier** — systematic post-build sanity check that catches regressions
- **Standing procedure** — a rule that applies to every passage automatically

---

## Project history

See **CHANGELOG.md** for the timeline of major design decisions.

---

## License / use

Private use under the Beast Mode Boating project.
