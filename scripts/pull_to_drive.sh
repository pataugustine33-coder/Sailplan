#!/usr/bin/env bash
# pull_to_drive.sh — fetch NWS/NDBC bulletins to a Google Drive-synced folder
#
# Designed to run on a machine you control (Mac/Linux/Pi/etc) where outbound
# HTTPS to NOAA is allowed. The output directory should be inside a folder
# that the Google Drive desktop app syncs to your Drive — that way the files
# land in Drive automatically, and Claude can read them via the Drive MCP at
# build time. No paste step.
#
# Configure DRIVE_DIR below to a Drive-synced path on your machine, e.g.:
#   macOS: ~/Library/CloudStorage/GoogleDrive-<you>@gmail.com/My Drive/sailbuild/pastes
#   Linux: ~/GoogleDrive/sailbuild/pastes  (with google-drive-ocamlfuse or similar)
#   Or just any folder you periodically upload manually.
#
# Schedule with cron / launchd / Task Scheduler — NWS issues CWF/AFD 4x/day,
# so every 3 hours is plenty:
#   crontab line:  0 */3 * * *  /path/to/pull_to_drive.sh chs-jax
#
# Usage:  pull_to_drive.sh <route>
#   route ∈ {chs-jax, chs-beaufort, chs-bahamas, chs-nantucket}

set -euo pipefail

DRIVE_DIR="${DRIVE_DIR:-$HOME/GoogleDrive/sailbuild/pastes}"
ROUTE="${1:-chs-jax}"

# Source manifest — mirrors the YAML manifests under inputs/sources/.
# Format: source_id|url   (the first URL that returns content with a recent
# issuance timestamp wins; if all stale, we still save the freshest one
# and let weather_pull's classifier tier it as STALE.)
declare -A SOURCES

case "$ROUTE" in
  chs-jax)
    SOURCES[cwf:CHS]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kchs.cwf.chs.txt"
    SOURCES[cwf:JAX]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kjax.cwf.jax.txt"
    SOURCES[afd:CHS]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus62.kchs.afd.chs.txt"
    SOURCES[afd:JAX]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus62.kjax.afd.jax.txt"
    SOURCES[buoy:41033]="https://www.ndbc.noaa.gov/data/realtime2/41033.txt"
    SOURCES[buoy:41008]="https://www.ndbc.noaa.gov/data/realtime2/41008.txt"
    SOURCES[buoy:41112]="https://www.ndbc.noaa.gov/data/realtime2/41112.txt"
    ;;
  chs-beaufort)
    SOURCES[cwf:CHS]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kchs.cwf.chs.txt"
    SOURCES[cwf:ILM]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kilm.cwf.ilm.txt"
    SOURCES[cwf:MHX]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kmhx.cwf.mhx.txt"
    SOURCES[afd:CHS]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus62.kchs.afd.chs.txt"
    SOURCES[afd:ILM]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus62.kilm.afd.ilm.txt"
    SOURCES[afd:MHX]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus62.kmhx.afd.mhx.txt"
    SOURCES[buoy:41033]="https://www.ndbc.noaa.gov/data/realtime2/41033.txt"
    SOURCES[buoy:41013]="https://www.ndbc.noaa.gov/data/realtime2/41013.txt"
    SOURCES[buoy:41025]="https://www.ndbc.noaa.gov/data/realtime2/41025.txt"
    ;;
  chs-bahamas)
    SOURCES[cwf:CHS]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kchs.cwf.chs.txt"
    SOURCES[cwf:JAX]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kjax.cwf.jax.txt"
    SOURCES[cwf:MFL]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kmfl.cwf.mfl.txt"
    SOURCES[cwf:KEY]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.keyw.cwf.eyw.txt"
    SOURCES[afd:CHS]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus62.kchs.afd.chs.txt"
    SOURCES[afd:MFL]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus62.kmfl.afd.mfl.txt"
    SOURCES[buoy:41033]="https://www.ndbc.noaa.gov/data/realtime2/41033.txt"
    SOURCES[buoy:41112]="https://www.ndbc.noaa.gov/data/realtime2/41112.txt"
    SOURCES[buoy:41114]="https://www.ndbc.noaa.gov/data/realtime2/41114.txt"
    ;;
  chs-nantucket)
    SOURCES[cwf:CHS]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kchs.cwf.chs.txt"
    SOURCES[cwf:MHX]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus52.kmhx.cwf.mhx.txt"
    SOURCES[cwf:AKQ]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus51.kakq.cwf.akq.txt"
    SOURCES[cwf:PHI]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus51.kphi.cwf.phi.txt"
    SOURCES[cwf:OKX]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus51.kokx.cwf.okx.txt"
    SOURCES[cwf:BOX]="https://tgftp.nws.noaa.gov/data/raw/fz/fzus51.kbox.cwf.box.txt"
    SOURCES[afd:MHX]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus62.kmhx.afd.mhx.txt"
    SOURCES[afd:OKX]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus61.kokx.afd.okx.txt"
    SOURCES[afd:BOX]="https://tgftp.nws.noaa.gov/data/raw/fx/fxus61.kbox.afd.box.txt"
    SOURCES[buoy:41001]="https://www.ndbc.noaa.gov/data/realtime2/41001.txt"
    SOURCES[buoy:44025]="https://www.ndbc.noaa.gov/data/realtime2/44025.txt"
    SOURCES[buoy:44008]="https://www.ndbc.noaa.gov/data/realtime2/44008.txt"
    ;;
  *)
    echo "Unknown route: $ROUTE" >&2
    echo "Valid routes: chs-jax, chs-beaufort, chs-bahamas, chs-nantucket" >&2
    exit 1
    ;;
esac

mkdir -p "$DRIVE_DIR"
INDEX="$DRIVE_DIR/_index.txt"
echo "# pull_to_drive.sh fetch — $(date -u +%Y-%m-%dT%H:%M:%SZ) — route=$ROUTE" > "$INDEX"

OK=0
FAIL=0
for sid in "${!SOURCES[@]}"; do
  url="${SOURCES[$sid]}"
  fname="${sid//:/_}.txt"
  if curl -sS --max-time 20 -A "sailbuild-pull/1.0" "$url" -o "$DRIVE_DIR/$fname.tmp"; then
    # Sanity check: file must be non-empty and look like an NWS or NDBC product
    if [[ -s "$DRIVE_DIR/$fname.tmp" ]]; then
      mv "$DRIVE_DIR/$fname.tmp" "$DRIVE_DIR/$fname"
      echo "OK  $sid  $url" | tee -a "$INDEX"
      OK=$((OK+1))
    else
      rm -f "$DRIVE_DIR/$fname.tmp"
      echo "EMPTY  $sid  $url" | tee -a "$INDEX"
      FAIL=$((FAIL+1))
    fi
  else
    echo "FAIL  $sid  $url" | tee -a "$INDEX"
    FAIL=$((FAIL+1))
  fi
done

echo "" | tee -a "$INDEX"
echo "Done — $OK ok, $FAIL failed, output in $DRIVE_DIR" | tee -a "$INDEX"
