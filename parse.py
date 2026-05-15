"""
CLI: parse raw NWS bulletin text files into a forecast.yaml.

Usage:
  python parse.py \
      --cwf CHS:cwfchs.txt --cwf JAX:cwfjax.txt \
      --afd CHS:afdchs.txt --afd JAX:afdjax.txt \
      --output forecast.yaml

Each --cwf and --afd argument is OFFICE:filename pair.
"""
import argparse
import sys
import yaml
from pathlib import Path

from sailbuild.parsers.nws_cwf import parse_cwf
from sailbuild.parsers.nws_afd import parse_afd
from sailbuild.parsers.assembler import assemble_forecast


def main():
    p = argparse.ArgumentParser(description="Parse NWS bulletins → forecast.yaml")
    p.add_argument("--cwf", action="append", default=[],
                   help="OFFICE:filename pair, e.g., CHS:cwfchs.txt")
    p.add_argument("--afd", action="append", default=[],
                   help="OFFICE:filename pair, e.g., CHS:afdchs.txt")
    p.add_argument("--output", required=True, help="Output forecast.yaml path")
    args = p.parse_args()

    cwfs = {}
    for spec in args.cwf:
        if ":" not in spec:
            print(f"  ✗ Bad --cwf spec '{spec}'; expected OFFICE:filename", file=sys.stderr)
            sys.exit(1)
        office, path = spec.split(":", 1)
        text = Path(path).read_text()
        parsed = parse_cwf(text)
        cwfs[office.upper()] = parsed
        print(f"  ✓ Parsed CWF {office}: {len(parsed.get('zones', {}))} zones, "
              f"synopsis={'yes' if parsed.get('synopsis') else 'no'}, "
              f"issued {parsed.get('issuance', {}).get('local_label', '?')}")

    afds = {}
    for spec in args.afd:
        if ":" not in spec:
            print(f"  ✗ Bad --afd spec '{spec}'", file=sys.stderr)
            sys.exit(1)
        office, path = spec.split(":", 1)
        text = Path(path).read_text()
        parsed = parse_afd(text)
        afds[office.upper()] = parsed
        print(f"  ✓ Parsed AFD {office}: {len(parsed.get('key_messages', []))} key messages, "
              f"marine={'yes' if parsed.get('marine') else 'no'}, "
              f"issued {parsed.get('issuance', {}).get('local_label', '?')}")

    bundle = assemble_forecast(cwfs, afds)

    # Add empty waypoint_assignments and plan_commentary stubs that the analyst
    # fills in. The skeleton tells the human what they need to edit.
    bundle["waypoint_assignments"] = {
        "_TODO": "For each plan in passage.yaml, list each WP with its zone+period assignment. "
                 "Optionally include override (wind_kt, wind_dir_deg, wind_dir_text) and "
                 "notes_addendum / weather_risk prose."
    }
    bundle["plan_commentary"] = {
        "_TODO": "For each plan in passage.yaml, write a description block. This appears at the top of "
                 "the Plan tab and synthesizes the cycle's implications for that departure scenario."
    }

    Path(args.output).write_text(yaml.safe_dump(bundle, sort_keys=False, allow_unicode=True, width=120))
    print(f"\n✓ Wrote forecast bundle: {args.output}")
    print(f"  Zones: {list(bundle['zones'].keys())}")
    print(f"  Synopses: {list(bundle.get('synopses', {}).keys())}")
    print(f"\n⚠ Next step: open {args.output} and fill in waypoint_assignments and plan_commentary for each plan.")


if __name__ == "__main__":
    main()
