# How to Use This System

You don't run this code yourself. Claude does. Here's the workflow.

---

## Day-of-passage workflow

### 1. Open Claude

Go to claude.ai (or the Claude app) and start a new conversation in the
**Beast Mode Boating** project. The project system prompt loads your
standing procedures automatically.

### 2. Tell Claude what you want

Example openings:

> "I'm planning a trip from Charleston to Beaufort NC departing Sunday at 2 PM. Vessel is the HR 48."

> "Build me a Sail Plan for Charleston to Bahamas departing Friday morning. Use the HR 54."

> "Refresh the Charleston-Savannah plan with the current 4 PM forecast cycle."

### 3. Claude pulls fresh weather via Chrome

You'll need Claude in Chrome connected — Claude will use whichever browser
is available (per your standing rule, no need to pick one).

Claude will fetch:
- **CWFCHS** (Coastal Waters Forecast) from the relevant NWS office
- **AFDCHS** (Area Forecast Discussion)
- **NDBC buoys** along your route (typically 3-5 stations)

### 4. Claude updates the YAML files

The repository has three input file types:
- `inputs/passages/your-route.yaml` — route, vessel, plans
- `inputs/forecasts/YYYY-MM-DD-cycle.yaml` — fresh weather
- `inputs/buoys/YYYY-MM-DD-cycle.yaml` — fresh observations

Claude updates these to match what was just pulled.

### 5. Claude builds the workbook

Running `python build.py` produces:
- `Sail Plan and Weather Risk - Charleston-Beaufort.xlsx`
- `Charleston-Beaufort_Route.kml`
- `Charleston-Beaufort_Route.gpx`

### 6. The verifier runs automatically

Seven systematic checks confirm the workbook is internally consistent.
If errors fire (e.g., red arrival, HR48 leak in HR54 workbook), Claude
flags them before delivery.

### 7. Claude hands you the files

You download from the conversation and you're done. Workbook into iPad,
KML into Google My Maps for visualization, GPX into the chartplotter.

---

## Underway workflow

When you're 12 hours into the passage and want to check the latest forecast:

> "Pull a fresh cycle and update the Charleston-Beaufort plan."

Claude will:
1. Re-pull current CWFCHS, AFDCHS, and buoys
2. Update the forecast and buoy YAMLs
3. Rebuild the workbook
4. Run the verifier
5. Flag any material change from the prior plan

The new workbook will reflect the latest weather. Your historical comparison
lives in the git history (every commit on GitHub timestamps the state of
the plan at that point).

---

## Post-passage workflow

After the trip, capture lessons:

> "Add a methodology lesson: 41033 Fripp was the better on-route buoy than 41004 Edisto for the Charleston-Savannah route."

Claude will append to `inputs/lessons.yaml` and rebuild so the Verification
Scorecard tab reflects the new lesson.

---

## What to do if something breaks

### "I think Claude is working from an old version of the code"

Each new Claude conversation pulls from this GitHub repo. If you've recently
made improvements but Claude is acting like it doesn't know about them:

1. Check the GitHub repo to confirm the latest version is uploaded.
2. Tell Claude: "Pull the latest version from
   github.com/pataugustine33-coder/Sailplan and confirm the version number
   matches the latest CHANGELOG entry."

### "The verifier is flagging errors"

Errors mean the workbook isn't safe to deliver. Read the error message —
it tells you exactly what's wrong:
- **Vessel leak** — wrong vessel mentioned somewhere; Claude will hunt it down
- **Red arrival** — ETA falls in night window; Claude will suggest
  departure shift via reverse calculator
- **Buoy coordinate mismatch** — YAML lat/lon doesn't match NDBC truth
- **Gust column empty** — forecast YAML missing `wind_gust_kt`

### "I want to change the methodology"

The methodology is in two places:
1. The Beast Mode Boating Claude project system prompt (settings → custom instructions)
2. This codebase (specifically the rules in `verify.py` and the column
   definitions in `tabs/plan.py`)

If you change methodology, update both so the system stays consistent.

---

## GitHub workflow for updates

When Claude makes improvements during a session:

### Option 1 — Manual upload (works fine for occasional updates)

1. At the end of the session, Claude produces `sailbuild.zip`.
2. Download it from the conversation.
3. Unzip on your computer.
4. Go to the GitHub repo in a browser.
5. Click **"Add file" → "Upload files"** and drag the unzipped contents in.
6. GitHub will ask for a "commit message" — type a brief description
   (e.g., "Added gust column verifier check").
7. Click "Commit changes."

### Option 2 — Ask Claude to push directly

Future enhancement: Claude can push directly to GitHub via the GitHub API.
This requires setting up a Personal Access Token. Worth setting up if you
want zero-friction updates.

---

## Standing rules reminder

These apply on every passage automatically:

1. Pull fresh weather at session start
2. Run end-of-run verifier before delivering workbook
3. Color-only weather risk (no GREEN/YELLOW/RED label words)
4. Times in 12-hour clock with day prefix ("Mon 1:50 PM")
5. Arrival timing is a planning constraint, not a result — daylight
   arrivals preferred; night arrivals get red and a departure-shift
   recommendation
6. Single-vessel format is default; Vessel Comparison is opt-in
7. When multiple Chrome browsers connected, pick any and proceed
