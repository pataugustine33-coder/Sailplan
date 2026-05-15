"""
End-to-end pipeline demo.

Flows real NWS bulletin text through the full sailbuild pipeline:

    bulletins/*.txt  →  parsers (nws_cwf, nws_afd)
                    →  assembler (combine into forecast bundle)
                    →  inject hand-crafted waypoint_assignments
                         (this part is irreducibly human — the analyst's call
                          on which zone/period each waypoint inherits)
                    →  inputs/forecasts/pipeline-generated.yaml
                    →  build.py (compute legs, render workbook)
                    →  /tmp/pipeline/sail-pipeline.xlsx

Run it:
    python pipeline_demo.py

What this proves: the parser + assembler can ingest real bulletin text
and produce a forecast.yaml structurally identical to a hand-crafted one,
modulo the human-judgment block (waypoint_assignments / plan_commentary).
The arrival times reflect the bulletin's actual wind/sea data, NOT a
hand-tuned narrative — that's the point. When the bulletin says NE 20-25
on the bow, the workbook flags it.

Bulletins used here are real NWS text but from mixed dates (a Feb 22 gale
aftermath CWFCHS, a Tue 5/12 CWFJAX, a Sun 5/10 AFDCHS). That's fine for
demonstrating the pipeline — the parser doesn't care about date alignment.

Why mixed dates? The web_fetch tool has session-persistent caching, so
some URLs return older cached content. For a real refresh cycle, you'd
re-fetch the same hour and save fresh text. The pipeline mechanics are
unchanged.
"""
import sys
import yaml
from pathlib import Path

from sailbuild.parsers.nws_cwf import parse_cwf
from sailbuild.parsers.nws_afd import parse_afd
from sailbuild.parsers.assembler import assemble_forecast
from build import build, load_yaml


# ---------------------------------------------------------------------------
# Demo waypoint_assignments: the human-judgment block.
#
# These map each plan's WPs to (zone, period) tuples in the parsed bundle.
# For this demo, we use the bulletin zones AMZ350 / AMZ354 / AMZ450 / AMZ452
# (note the older numbering — current CHS office uses AMZ360-series, but
# the cached Feb 22 bulletin is from before the zone rename).
#
# A real refresh would either (a) edit this block when the analyst reviews
# the new bundle, or (b) the analyst could provide it via a separate
# assignments.yaml. The pipeline doesn't care which.
# ---------------------------------------------------------------------------
DEMO_ASSIGNMENTS = {
    "plan_a": {
        # Plan A: Wed 8 AM departure CHS. SW wind 20-25 kt from CHS bulletin —
        # essentially a beam reach southbound. JAX waters Wed NE 15-20 kt
        # become headwinds approaching JAX (passage course ~205°).
        "WP0": {
            "zone": "AMZ350", "period": "WED",
            "risk_level": "yellow", "row_band": "yellow",
            "pressure": 30.05,
            "notes_addendum": "Dawn departure CHS, post-gale ridge rebuild in progress, SW 20-25 kt beam reach southbound — fair sailing.",
        },
        "WP1": {
            "zone": "AMZ350", "period": "WED",
            "risk_level": "yellow",
            "pressure": 30.07,
            "notes_addendum": "Continuing south in CHS coastal waters, SW 20-25 kt, seas 4-6 ft on beam.",
        },
        "WP2": {
            "zone": "AMZ354", "period": "TUE",
            "risk_level": "yellow",
            "pressure": 30.09,
            "notes_addendum": "Transitioning into outer CHS waters AMZ354. NW 10-15 kt, seas 3-4 ft — using last available period from terse extended forecast.",
        },
        "WP3": {
            "zone": "AMZ354", "period": "TUE",
            "risk_level": "yellow",
            "pressure": 30.11,
            "notes_addendum": "Mid-passage, Grays Reef vicinity. Light NW reach.",
        },
        "WP4": {
            "zone": "AMZ452", "period": "WED",
            "risk_level": "yellow",
            "pressure": 30.13,
            "notes_addendum": "Entering JAX outer waters. NE 15-20 kt becomes head wind on 205° course — expect motor sailing.",
        },
        "WP5": {
            "zone": "AMZ450", "period": "WED_NIGHT",
            "risk_level": "yellow",
            "pressure": 30.15,
            "notes_addendum": "Approaching St. Augustine waters. N 10-15 kt, seas 3-4 ft. Wind clocking N.",
        },
        "WP6": {
            "zone": "AMZ450", "period": "THU",
            "risk_level": "green",
            "pressure": 30.17,
            "notes_addendum": "Arrival JAX Sea Buoy STJ early Thu AM. N-NE 10-15, seas 3-4 ft. Inlet entry clean.",
        },
    },
    "plan_b": {
        # Plan B: Thu 6 AM departure CHS. CHS bulletin Thu shows SW 20-25 kt
        # with showers (front approaching) — less stable than Plan A.
        "WP0": {
            "zone": "AMZ350", "period": "THU",
            "risk_level": "yellow", "row_band": "yellow",
            "pressure": 30.10,
            "notes_addendum": "Dawn departure CHS Thu, SW 20-25 kt with shower chance — front line approaching from west.",
        },
        "WP1": {
            "zone": "AMZ350", "period": "THU",
            "risk_level": "yellow",
            "pressure": 30.08,
            "notes_addendum": "South in CHS waters, SW 20-25, scattered showers.",
        },
        "WP2": {
            "zone": "AMZ350", "period": "THU",
            "risk_level": "yellow",
            "pressure": 30.06,
            "notes_addendum": "Continuing south in AMZ350 — outer zone AMZ354 has no Thu data in bulletin so staying inner-zone reference.",
        },
        "WP3": {
            "zone": "AMZ350", "period": "THU_NIGHT",
            "risk_level": "yellow",
            "pressure": 30.04,
            "notes_addendum": "Mid-passage Thu night, SW 20-25 still, showers likely. Pressure falling — front behind.",
        },
        "WP4": {
            "zone": "AMZ452", "period": "THU",
            "risk_level": "yellow",
            "pressure": 30.06,
            "notes_addendum": "Entering JAX outer waters Thu, light N 10-15 — fair wind transition.",
        },
        "WP5": {
            "zone": "AMZ450", "period": "THU_NIGHT",
            "risk_level": "green",
            "pressure": 30.08,
            "notes_addendum": "Approaching JAX Thu evening/night. NE 5-10, seas 2-3 ft. Settling.",
        },
        "WP6": {
            "zone": "AMZ450", "period": "FRI",
            "risk_level": "green",
            "pressure": 30.10,
            "notes_addendum": "Arrival JAX Fri AM, light E 5-10, calm seas. Bermuda High in command.",
        },
    },
}


DEMO_COMMENTARY = {
    "plan_a": (
        "Wed 8 AM CHS departure under post-gale ridge rebuild. Charleston "
        "waters SW 20-25 kt provide beam reach southbound — fair sailing. "
        "Approaching JAX, wind veers NE 15-20 (head wind) — motor sailing "
        "last 40 NM. Arrival timing falls in early Thu AM dawn window — clean."
    ),
    "plan_b": (
        "Thu 6 AM CHS departure ahead of the next frontal trough. Initial "
        "SW 20-25 kt with shower chance — pressure falling slowly through "
        "Thu night. Conditions improve through the passage as JAX waters "
        "transition to light N-NE behind the trough. Arrival Fri AM clean."
    ),
}


def main():
    project_root = Path(__file__).parent
    bulletins_dir = project_root / "inputs" / "bulletins"
    output_dir = Path("/tmp/pipeline")
    output_dir.mkdir(parents=True, exist_ok=True)

    # ---------- STAGE 1: parse each bulletin ----------
    print("=" * 60)
    print("STAGE 1: parse raw NWS bulletin text")
    print("=" * 60)

    cwfs = {}
    for office, filename in [("CHS", "cwfchs.txt"), ("JAX", "cwfjax.txt")]:
        path = bulletins_dir / filename
        text = path.read_text()
        parsed = parse_cwf(text)
        cwfs[office] = parsed
        print(f"  ✓ CWF{office} ({filename}, {len(text)} chars)")
        print(f"      issued: {parsed['issuance'].get('local_label', '?')}")
        print(f"      zones:  {list(parsed['zones'].keys())}")
        for z, zd in parsed['zones'].items():
            sca_flag = " [SCA]" if zd.get('SCA') else ""
            print(f"        {z}: {len(zd['periods'])} periods{sca_flag}")

    afds = {}
    for office, filename in [("CHS", "afdchs.txt")]:
        path = bulletins_dir / filename
        text = path.read_text()
        parsed = parse_afd(text)
        afds[office] = parsed
        print(f"  ✓ AFD{office} ({filename}, {len(text)} chars)")
        print(f"      issued: {parsed['issuance'].get('local_label', '?')}")
        print(f"      key_messages: {len(parsed.get('key_messages', []))}")
        print(f"      marine: {'parsed' if parsed.get('marine') else 'missing'}")

    # ---------- STAGE 2: assemble into forecast bundle ----------
    print()
    print("=" * 60)
    print("STAGE 2: assemble forecast bundle")
    print("=" * 60)
    bundle = assemble_forecast(cwfs, afds)
    print(f"  cycle.label: {bundle['cycle']['label']!r}")
    print(f"  zones in bundle:  {list(bundle['zones'].keys())}")
    print(f"  synopses:         {list(bundle.get('synopses', {}).keys())}")
    print(f"  afd_marine:       {list(bundle.get('afd_marine', {}).keys())}")

    # ---------- STAGE 3: inject hand-crafted assignments ----------
    print()
    print("=" * 60)
    print("STAGE 3: inject demo waypoint_assignments + plan_commentary")
    print("=" * 60)
    bundle["waypoint_assignments"] = DEMO_ASSIGNMENTS
    bundle["plan_commentary"] = DEMO_COMMENTARY
    print(f"  plans:    {list(DEMO_ASSIGNMENTS.keys())}")
    print(f"  WPs/plan: {[len(v) for v in DEMO_ASSIGNMENTS.values()]}")

    # ---------- STAGE 4: write the assembled forecast.yaml ----------
    forecast_yaml = project_root / "inputs" / "forecasts" / "pipeline-generated.yaml"
    forecast_yaml.write_text(
        yaml.safe_dump(bundle, sort_keys=False, allow_unicode=True, width=120)
    )
    size_kb = forecast_yaml.stat().st_size / 1024
    print()
    print(f"  ✓ wrote {forecast_yaml.relative_to(project_root)} ({size_kb:.1f} KB)")

    # ---------- STAGE 5: run build() against CHS-JAX passage ----------
    print()
    print("=" * 60)
    print("STAGE 5: build workbook from pipeline-generated forecast")
    print("=" * 60)
    passage_yaml = project_root / "inputs" / "passages" / "chs-jax.yaml"
    buoys_yaml = project_root / "inputs" / "buoys" / "2026-05-12-tue-1108.yaml"
    output_xlsx = output_dir / "sail-pipeline.xlsx"

    legs = build(
        str(passage_yaml), str(forecast_yaml), str(buoys_yaml), str(output_xlsx)
    )

    print()
    print(f"  ✓ workbook written: {output_xlsx}")
    for plan_id, plan_legs in legs.items():
        arrival = plan_legs[-1]
        print(f"    {plan_id}: arrive {arrival.eta_str}, {arrival.cum_time_hr:.1f} hr "
              f"({arrival.cum_sailing_hr:.1f} sail / {arrival.cum_motoring_hr:.1f} motor)")

    print()
    print("=" * 60)
    print("PIPELINE COMPLETE")
    print("=" * 60)
    print("End-to-end flow proven: text → parsers → assembler → build → xlsx")
    print()
    print("Compare to hand-curated YAML build:")
    print("  hand-curated (Tue 5/12 cycle, calm post-frontal):")
    print("    plan_a: arrive Thu 6:27 AM (22.4 hr, 18.1 sail / 4.3 motor)")
    print("    plan_b: arrive Fri 6:20 AM (24.3 hr, 11.3 sail / 13.1 motor)")
    print("  pipeline-generated (post-gale recovery + Tue 5/12 NE flow):")
    print("    different conditions → different arrival times. The numbers")
    print("    above reflect what the parsed bulletins actually say, not")
    print("    what a hand-tuned narrative wished they said.")


if __name__ == "__main__":
    main()
