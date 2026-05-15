"""
weather_pull.py — systematic weather source acquisition for a passage.

Usage:
    # Pull all sources defined in a route's source manifest
    python weather_pull.py inputs/sources/chs-jax.yaml

    # Re-read a paste file the user prepared in chat
    python weather_pull.py inputs/sources/chs-jax.yaml --ingest-pastes paste_response.txt

    # Show status without fetching (just read existing cached pulls)
    python weather_pull.py inputs/sources/chs-jax.yaml --status

Workflow:
    1. Read the route source manifest (YAML).
    2. For each source, try URLs in order; accept the first that returns text
       with a parseable, FRESH or STALE issuance timestamp. Reject ARCHIVED.
    3. Save accepted text to inputs/pastes/<source_id>.txt (overwriting any
       prior pull) AND write a status JSON to inputs/pastes/_status.json.
    4. For sources that failed all URLs, emit a paste request to stdout.
    5. Exit codes: 0 = all required sources covered; 2 = some required
       missing (build.py --require should check this).

Paste ingestion:
    The user pastes text into chat. Claude (this assistant) writes the
    content to inputs/pastes/<source_id>.txt and reruns this with --status
    to re-check coverage. The source-id-prefixed format the user pastes
    looks like:
        --- cwf:CHS ---
        FZUS52 KCHS 122130
        CWFCHS
        ...
        --- buoy:41033 ---
        Conditions at 41033 as of...
        ...
"""
import argparse
import json
import sys
import urllib.request
import urllib.error
from datetime import datetime, timezone
from pathlib import Path

import yaml

from sailbuild.freshness import assess


def banner(s):
    bar = "=" * (len(s) + 4)
    print(f"\n{bar}\n  {s}\n{bar}")


def fetch_url(url: str, timeout: float = 15.0) -> str | None:
    """Fetch a URL, return its body as text, or None on any failure."""
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "sailbuild/1.0 (passage-planning)"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
            # Try utf-8 first, fall back to iso-8859-1 (NWS sometimes uses this)
            for enc in ("utf-8", "iso-8859-1"):
                try:
                    return data.decode(enc)
                except UnicodeDecodeError:
                    continue
            return data.decode("utf-8", errors="replace")
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, Exception):
        return None


def try_source(source: dict, now_utc: datetime, pastes_dir: Path) -> dict:
    """Try each URL in a source's fallback chain. Return a status dict."""
    sid = source["id"]
    kind = source["kind"]
    refresh_hr = float(source.get("refresh_hr", 6))
    urls = source["urls"]

    result = {
        "id": sid,
        "kind": kind,
        "required": bool(source.get("required", False)),
        "refresh_hr": refresh_hr,
        "location": source.get("location", ""),
        "attempts": [],
        "accepted_url": None,
        "tier": None,
        "age_str": None,
        "issued_utc": None,
        "saved_to": None,
        "covered": False,
    }

    for url in urls:
        attempt = {"url": url, "fetched": False, "tier": None, "age_str": None}
        text = fetch_url(url)
        if text is None:
            attempt["error"] = "fetch_failed"
            result["attempts"].append(attempt)
            continue
        attempt["fetched"] = True
        attempt["bytes"] = len(text)
        # Assess freshness
        a = assess(text, kind, refresh_hr, now_utc=now_utc)
        attempt["tier"] = a["tier"]
        attempt["age_str"] = a["age_str"]
        result["attempts"].append(attempt)

        # Accept if FRESH or STALE; reject ARCHIVED/UNKNOWN and try next URL
        if a["tier"] in ("FRESH", "STALE"):
            out_path = pastes_dir / f"{sid.replace(':', '_')}.txt"
            out_path.write_text(text)
            # Clear any old paste sentinel — this is now a fetch-sourced copy
            sentinel = pastes_dir / f"{sid.replace(':', '_')}.paste"
            if sentinel.exists():
                sentinel.unlink()
            result["accepted_url"] = url
            result["tier"] = a["tier"]
            result["age_str"] = a["age_str"]
            result["issued_utc"] = a["issued_utc"].isoformat() if a["issued_utc"] else None
            result["saved_to"] = str(out_path)
            result["covered"] = True
            return result

    # All URLs failed (no fresh data found in any)
    # If a paste file already exists from a prior --ingest-pastes, use it
    paste_path = pastes_dir / f"{sid.replace(':', '_')}.txt"
    if paste_path.exists():
        text = paste_path.read_text()
        a = assess(text, kind, refresh_hr, now_utc=now_utc)
        if a["tier"] in ("FRESH", "STALE"):
            result["accepted_url"] = "user_paste"
            result["tier"] = a["tier"]
            result["age_str"] = a["age_str"]
            result["issued_utc"] = a["issued_utc"].isoformat() if a["issued_utc"] else None
            result["saved_to"] = str(paste_path)
            result["covered"] = True
            return result

    return result


def print_status_table(results: list[dict]):
    """Print a compact status table to stdout."""
    print(f"\n{'ID':<18}{'Kind':<10}{'Req':<5}{'Source':<32}{'Tier':<10}{'Age':<12}")
    print("-" * 87)
    for r in results:
        req_mark = "yes" if r["required"] else " "
        if r["covered"]:
            src = r["accepted_url"]
            if src == "user_paste":
                src_short = "(user paste)"
            else:
                src_short = src.replace("https://", "")[:30]
            print(f"{r['id']:<18}{r['kind']:<10}{req_mark:<5}{src_short:<32}"
                  f"{r['tier']:<10}{r['age_str']:<12}")
        else:
            print(f"{r['id']:<18}{r['kind']:<10}{req_mark:<5}{'FAILED ALL URLS':<32}"
                  f"{'—':<10}{'—':<12}")


def emit_paste_requests(failed: list[tuple[dict, dict]]):
    """Print structured paste requests for failed sources."""
    if not failed:
        return
    banner(f"PASTE REQUEST — {len(failed)} sources need browser paste")
    print("Open these URLs in your browser and paste the bulletin body back")
    print("to chat, prefixed with the source ID like '--- cwf:CHS ---'.\n")
    for i, (source, result) in enumerate(failed, 1):
        print(f"[{i}] {source['id']}  ({source['kind']})")
        if result.get("attempts"):
            tried = ", ".join(
                f"{a.get('tier') or 'fail'}({a.get('age_str') or '?'})"
                for a in result["attempts"]
            )
            print(f"    Tried: {tried}")
        print(f"    Open: {source['urls'][0]}")
        if source.get("paste_hint"):
            for line in source["paste_hint"].strip().splitlines():
                print(f"    {line}")
        print()


def ingest_drive_folder(folder: Path, pastes_dir: Path, manifest: dict, now_utc: datetime) -> list[dict]:
    """Ingest a folder of bulletin files written by pull_to_drive.sh (or any
    process that drops one bulletin per file using the convention
    <sid_with_underscore>.txt — e.g. cwf_CHS.txt, buoy_41033.txt).

    This is the architectural alternative to copy/paste: a user-side script
    fetches NOAA endpoints from a network that's allowed to reach them, dumps
    the responses to a Google Drive-synced folder, and this function reads
    that folder. Provenance is marked as "drive_sync" so the freshness panel
    can distinguish it from raw user paste."""
    if not folder.exists() or not folder.is_dir():
        print(f"  ⚠ ingest-drive folder not found: {folder}")
        return []
    sources_by_id = {s["id"]: s for s in manifest["sources"]}
    saved = []
    for sid, src in sources_by_id.items():
        fname = sid.replace(":", "_") + ".txt"
        src_path = folder / fname
        if not src_path.exists():
            continue
        body = src_path.read_text()
        if not body.strip():
            print(f"  ⚠ {sid} file in drive folder is empty, skipping")
            continue
        out = pastes_dir / fname
        out.write_text(body)
        # Sentinel: distinguishes drive-sync source from raw user paste so
        # the freshness panel can show "drive sync" instead of "user paste".
        (pastes_dir / f"{sid.replace(':', '_')}.paste").write_text("drive_sync\n")
        a = assess(body, src["kind"], float(src.get("refresh_hr", 6)), now_utc=now_utc)
        print(f"  ✓ ingested {sid} ({a['tier']}, {a['age_str']}) from drive folder")
        saved.append({"id": sid, "tier": a["tier"]})
    return saved


def ingest_paste_file(path: Path, pastes_dir: Path, manifest: dict, now_utc: datetime) -> list[dict]:
    """Split a paste response file by '--- source:id ---' markers, save each
    section to its own file, and re-assess freshness."""
    text = path.read_text()
    # Split on lines like "--- cwf:CHS ---" (allow surrounding whitespace)
    import re
    splits = re.split(r"^\s*---\s*([\w:]+)\s*---\s*$", text, flags=re.MULTILINE)
    # splits[0] is anything before the first marker; subsequent items pair (id, body)
    saved = []
    sources_by_id = {s["id"]: s for s in manifest["sources"]}
    for i in range(1, len(splits), 2):
        sid = splits[i].strip()
        body = splits[i + 1].strip() if i + 1 < len(splits) else ""
        if sid not in sources_by_id:
            print(f"  ⚠ paste section '{sid}' not in manifest, skipping")
            continue
        if not body:
            print(f"  ⚠ paste section '{sid}' is empty, skipping")
            continue
        out = pastes_dir / f"{sid.replace(':', '_')}.txt"
        out.write_text(body)
        # Sentinel: mark this file as user-pasted so future --status reads
        # preserve the provenance instead of showing "(cached)".
        (pastes_dir / f"{sid.replace(':', '_')}.paste").write_text("user_paste\n")
        src = sources_by_id[sid]
        a = assess(body, src["kind"], float(src.get("refresh_hr", 6)), now_utc=now_utc)
        saved.append({
            "id": sid,
            "saved_to": str(out),
            "bytes": len(body),
            "tier": a["tier"],
            "age_str": a["age_str"],
        })
        print(f"  ✓ ingested {sid}: {len(body)} bytes, {a['tier']} ({a['age_str']})")
    return saved


def main():
    ap = argparse.ArgumentParser(description="Systematic weather pull for a passage")
    ap.add_argument("manifest", help="Path to inputs/sources/<route>.yaml")
    ap.add_argument("--status", action="store_true",
                    help="Don't fetch — just re-read existing pastes and report")
    ap.add_argument("--ingest-pastes",
                    help="Path to a paste-response file to split and ingest first")
    ap.add_argument("--ingest-drive",
                    help="Path to a folder of bulletin files (e.g. a Google Drive "
                         "-synced folder populated by scripts/pull_to_drive.sh). "
                         "Files named <sid_with_underscore>.txt are ingested as if "
                         "pasted, e.g. cwf_CHS.txt -> source id 'cwf:CHS'.")
    ap.add_argument("--pastes-dir", default=None,
                    help="Directory to store fetched/pasted bulletins (default inputs/pastes/)")
    args = ap.parse_args()

    manifest_path = Path(args.manifest)
    if not manifest_path.exists():
        print(f"✗ Manifest not found: {manifest_path}")
        sys.exit(1)
    manifest = yaml.safe_load(manifest_path.read_text())

    pastes_dir = Path(args.pastes_dir) if args.pastes_dir else manifest_path.parent.parent / "pastes"
    pastes_dir.mkdir(parents=True, exist_ok=True)

    now_utc = datetime.now(timezone.utc)
    print(f"Route: {manifest['route']}    Pull time: {now_utc.isoformat()}")

    # Phase 1: ingest from paste file OR drive folder if provided
    if args.ingest_pastes:
        banner("Ingest pastes")
        ingest_paste_file(Path(args.ingest_pastes), pastes_dir, manifest, now_utc)
    if args.ingest_drive:
        banner("Ingest from Drive folder")
        ingest_drive_folder(Path(args.ingest_drive), pastes_dir, manifest, now_utc)

    # Phase 2: fetch each source (or read existing paste if --status)
    if args.status:
        banner("Status (no fetch)")
    else:
        banner("Fetching sources")

    results = []
    for source in manifest["sources"]:
        if args.status:
            # Just check whether a paste file exists for this source
            sid = source["id"]
            paste_path = pastes_dir / f"{sid.replace(':', '_')}.txt"
            r = {"id": sid, "kind": source["kind"], "required": bool(source.get("required", False)),
                 "refresh_hr": float(source.get("refresh_hr", 6)),
                 "location": source.get("location", ""),
                 "attempts": [], "accepted_url": None, "covered": False}
            if paste_path.exists():
                text = paste_path.read_text()
                a = assess(text, source["kind"], r["refresh_hr"], now_utc=now_utc)
                if a["tier"] in ("FRESH", "STALE"):
                    # Check sentinel — provenance varies by ingest origin
                    sentinel = pastes_dir / f"{sid.replace(':', '_')}.paste"
                    if sentinel.exists():
                        marker = sentinel.read_text().strip()
                        accepted = marker if marker else "user_paste"
                    else:
                        accepted = "(cached)"
                    r.update(accepted_url=accepted, tier=a["tier"], age_str=a["age_str"],
                             issued_utc=a["issued_utc"].isoformat() if a["issued_utc"] else None,
                             saved_to=str(paste_path), covered=True)
                else:
                    r.update(tier=a["tier"], age_str=a["age_str"])
        else:
            print(f"  → {source['id']} ...")
            r = try_source(source, now_utc, pastes_dir)
            if r["covered"]:
                print(f"    ✓ {r['tier']} ({r['age_str']}) via {r['accepted_url'][:60]}")
            else:
                print(f"    ✗ no fresh source found (tried {len(r['attempts'])} URLs)")
        results.append(r)

    # Phase 3: Status table + paste requests for failures
    print_status_table(results)

    # Identify required sources still missing
    failed = []
    for source, result in zip(manifest["sources"], results):
        if not result["covered"] and source.get("required", False):
            failed.append((source, result))

    emit_paste_requests(failed)

    # Phase 4: Write status JSON for build.py to consume
    status_path = pastes_dir / "_status.json"
    status_path.write_text(json.dumps({
        "route": manifest["route"],
        "pulled_at": now_utc.isoformat(),
        "sources": results,
    }, indent=2, default=str))
    print(f"\nStatus written to {status_path}")

    # Exit code: 2 if any required source is missing
    sys.exit(2 if failed else 0)


if __name__ == "__main__":
    main()
