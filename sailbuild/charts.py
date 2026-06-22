"""
Chart and visualization generators.

Produces PNG bytes suitable for embedding in workbook tabs the same way
`rose.py` does. Uses matplotlib with the Agg backend (no display server).

Available chart types:
  - radar_chart_png_bytes(axes_data) — six-axis radar for Risk Bowtie
  - radar_overlay_png_bytes(plans_data) — multi-plan radar overlay
  - timeline_strip_png_bytes(legs) — wind/sea strip across passage
  - mini_polar_png_bytes(tws, twa, design) — small polar with TWA marker

All return raw PNG bytes; embed via openpyxl.drawing.image.Image with io.BytesIO.
"""
import io
import math

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

# Color palette aligned with workbook styles
COLOR_TITLE = "#1F3864"
COLOR_PLAN_A = "#305496"   # Dark blue — primary plan
COLOR_PLAN_B = "#C0504D"   # Muted red — alternative
COLOR_GOOD = "#70AD47"
COLOR_NEUTRAL = "#FFC000"
COLOR_BAD = "#C0504D"
COLOR_GRID = "#BFBFBF"
COLOR_BAND = "#F2F2F2"

# Wind/sea timeline palette — kept distinct from plan/risk colors so the
# chart "reads right" intuitively: water = blue, wind = warm tones.
COLOR_WIND_BAR = "#F4B183"        # Warm peach — sustained wind bar fill
COLOR_WIND_EDGE = "#C65911"       # Burnt orange — bar edge + label
COLOR_WIND_GUST = "#843C0C"       # Dark rust — gust marker + label (still wind family)
COLOR_SEA_LINE = "#2E75B6"        # Water blue — sea height line + label
COLOR_SCA_THRESHOLD = "#FFC000"   # Amber — SCA caution line at 18 kt
COLOR_REEF_THRESHOLD = "#C00000"  # Red — reef trigger line at 25 kt


def _score_color(score):
    """Map a 1-10 score to a fill color."""
    if score >= 7:
        return COLOR_GOOD
    elif score >= 4:
        return COLOR_NEUTRAL
    return COLOR_BAD


# ======================================================================
# Risk Bowtie radar chart — six-axis spider plot
# ======================================================================
def radar_chart_png_bytes(axes_data, title="Risk Profile", output_px=520):
    """Render a radar chart showing 6 axes scored 1-10.

    axes_data: list of dicts with keys 'name' and 'score' (1-10)
    Returns BytesIO containing PNG.
    """
    labels = [a["name"] for a in axes_data]
    scores = [a["score"] for a in axes_data]
    n = len(labels)

    # Angles for each axis (evenly distributed around the circle), starting
    # at the top and going clockwise. Use np.linspace + modulo to keep all
    # values in [0, 2π] — matplotlib polar plots get confused when angles
    # exceed 2π and may render only a partial wedge.
    base_angles = np.linspace(0, 2 * math.pi, n, endpoint=False)
    # Rotate so first axis is at top, going clockwise: angle = π/2 - i*step
    angles = [(math.pi / 2 - a) % (2 * math.pi) for a in base_angles]
    angles_closed = angles + [angles[0]]
    scores_closed = scores + [scores[0]]

    fig, ax = plt.subplots(figsize=(output_px / 100, output_px / 100),
                           subplot_kw=dict(polar=True), dpi=100)

    # Plot the polygon
    ax.plot(angles_closed, scores_closed, color=COLOR_PLAN_A, linewidth=2.5,
            marker="o", markersize=8, markerfacecolor=COLOR_PLAN_A,
            markeredgecolor="white", markeredgewidth=1.5, zorder=5)
    ax.fill(angles_closed, scores_closed, color=COLOR_PLAN_A, alpha=0.18, zorder=4)

    # Background rings showing score zones
    ax.fill_between(np.linspace(0, 2*math.pi, 100), 0, 4,
                    color=COLOR_BAD, alpha=0.08, zorder=1)
    ax.fill_between(np.linspace(0, 2*math.pi, 100), 4, 7,
                    color=COLOR_NEUTRAL, alpha=0.08, zorder=2)
    ax.fill_between(np.linspace(0, 2*math.pi, 100), 7, 10,
                    color=COLOR_GOOD, alpha=0.10, zorder=3)

    # Axis labels around the perimeter
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=10, color=COLOR_TITLE, fontweight="bold")

    # Radial grid 0-10 with integer ticks
    ax.set_yticks([2, 4, 6, 8, 10])
    ax.set_yticklabels(["2", "4", "6", "8", "10"], fontsize=8, color="#595959")
    ax.set_ylim(0, 10)
    ax.set_rlabel_position(225)  # Move radial labels out of the way

    # Spines + grid styling
    ax.grid(color=COLOR_GRID, linewidth=0.8, alpha=0.6)
    ax.spines["polar"].set_color(COLOR_GRID)
    ax.spines["polar"].set_linewidth(0.8)

    # Title
    ax.set_title(title, fontsize=12, color=COLOR_TITLE, fontweight="bold", pad=20)

    # Score annotations at each vertex (the actual number)
    for ang, sc, name in zip(angles, scores, labels):
        ax.annotate(
            f"{sc:.1f}",
            xy=(ang, sc),
            xytext=(0, 12),
            textcoords="offset points",
            ha="center", va="center",
            fontsize=9, fontweight="bold",
            color=_score_color(sc),
            bbox=dict(boxstyle="round,pad=0.25", facecolor="white",
                      edgecolor=_score_color(sc), linewidth=1.2),
        )

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=100, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf


# ======================================================================
# Multi-plan overlay radar — compare Plan A vs Plan B at a glance
# ======================================================================
def radar_overlay_png_bytes(plans_data, title="Plan Comparison", output_px=580):
    """Render an overlaid radar with multiple plans for at-a-glance comparison.

    plans_data: list of dicts with keys:
      - 'label' (e.g., 'Plan A - Fri 10 AM')
      - 'axes': list of {'name', 'score'}
      - 'color' (optional hex string)
    """
    if not plans_data:
        return None

    labels = [a["name"] for a in plans_data[0]["axes"]]
    n = len(labels)
    base_angles = np.linspace(0, 2 * math.pi, n, endpoint=False)
    angles = [(math.pi / 2 - a) % (2 * math.pi) for a in base_angles]
    angles_closed = angles + [angles[0]]

    colors_default = [COLOR_PLAN_A, COLOR_PLAN_B, "#9E480E", "#636363"]

    fig, ax = plt.subplots(figsize=(output_px / 100, output_px / 100),
                           subplot_kw=dict(polar=True), dpi=100)

    # Background zones
    ax.fill_between(np.linspace(0, 2*math.pi, 100), 0, 4,
                    color=COLOR_BAD, alpha=0.06, zorder=1)
    ax.fill_between(np.linspace(0, 2*math.pi, 100), 4, 7,
                    color=COLOR_NEUTRAL, alpha=0.06, zorder=2)
    ax.fill_between(np.linspace(0, 2*math.pi, 100), 7, 10,
                    color=COLOR_GOOD, alpha=0.08, zorder=3)

    # Plot each plan's polygon
    for i, plan in enumerate(plans_data):
        color = plan.get("color", colors_default[i % len(colors_default)])
        scores = [a["score"] for a in plan["axes"]]
        scores_closed = scores + [scores[0]]
        ax.plot(angles_closed, scores_closed, color=color, linewidth=2.2,
                marker="o", markersize=6, label=plan["label"],
                markeredgecolor="white", markeredgewidth=1.2, zorder=5+i)
        ax.fill(angles_closed, scores_closed, color=color, alpha=0.15,
                zorder=4+i)

    # Axis labels
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=10, color=COLOR_TITLE, fontweight="bold")
    ax.set_yticks([2, 4, 6, 8, 10])
    ax.set_yticklabels(["2", "4", "6", "8", "10"], fontsize=8, color="#595959")
    ax.set_ylim(0, 10)
    ax.set_rlabel_position(225)
    ax.grid(color=COLOR_GRID, linewidth=0.8, alpha=0.6)
    ax.spines["polar"].set_color(COLOR_GRID)
    ax.spines["polar"].set_linewidth(0.8)

    ax.set_title(title, fontsize=12, color=COLOR_TITLE, fontweight="bold", pad=20)
    ax.legend(loc="upper right", bbox_to_anchor=(1.30, 1.10), fontsize=9,
              frameon=True, edgecolor=COLOR_GRID)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=100, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf


# ======================================================================
# Wind/sea timeline strip — sparkline across the passage
# ======================================================================
def _short_location_from_wp_name(wp_name: str, max_len: int = 14) -> str:
    """Derive a short city/location label from a verbose WP name, for use
    as a second-line x-axis tick label on the timeline strip.

    Strategy:
      1. "Current position (in Stream E of <loc>)" → "<loc>"
      2. "Off <name>" → strip "Off "
      3. Drop parenthetical content
      4. Drop trailing zone-boundary tags ("AMZ650/AMZ450 zone boundary 30.7N")
      5. Drop trailing latitude markers ("28.5N")
      6. Drop "abeam (...)" descriptors
      7. "<a> / <b>" → "<a>"
      8. Drop common trailing nautical suffixes aggressively:
            Inlet Sea Buoy <id>, Sea Buoy <id>, Harbor Entrance, NMS, Tower
      9. If still over max_len, try dropping trailing "Beach" / "Inlet"
     10. If still over max_len, truncate with ".."

    Examples (max_len=14):
      "Current position (in Stream E of Vero Beach)" → "Vero Beach"
      "Off Cape Canaveral 28.5N"                    → "Cape Canaveral"
      "Off Jacksonville Beach"                      → "Jacksonville"
      "Off Savannah / Grays Reef NMS"               → "Savannah"
      "Charleston Sea Buoy 2CL"                     → "Charleston"
      "Charleston Harbor Entrance"                  → "Charleston"
      "Lake Worth Inlet Sea Buoy LW"                → "Lake Worth"

    Note: this heuristic may produce identical labels for WPs that share
    a city name (e.g., both Charleston WPs above collapse to "Charleston").
    Use the YAML waypoint `chart_label` field to override per-WP for
    disambiguation.
    """
    import re
    if not wp_name:
        return ""
    s = wp_name.strip()

    if s.lower().startswith("current position"):
        m = re.search(r'\(([^)]+)\)', s)
        if m:
            inner = m.group(1)
            of_match = re.search(r'\bof\s+(.+)$', inner, re.IGNORECASE)
            if of_match:
                return of_match.group(1).strip()
            return inner.strip()
        return "Current pos."

    if s.lower().startswith("off "):
        s = s[4:].strip()

    s = re.sub(r'\s*\([^)]*\)', '', s).strip()
    s = re.sub(r'\s*AMZ\d+/AMZ\d+\s+zone\s+boundary.*$', '', s, flags=re.IGNORECASE).strip()
    s = re.sub(r'\s+\d+\.?\d*\s*[NS]\s*$', '', s).strip()
    s = re.sub(r'\s+abeam(\s+.*)?$', '', s, flags=re.IGNORECASE).strip()

    if '/' in s:
        s = s.split('/')[0].strip()

    s = re.sub(r'\s+Inlet\s+Sea\s+Buoy(\s+\w+)?\s*$', '', s, flags=re.IGNORECASE).strip()
    s = re.sub(r'\s+Sea\s+Buoy(\s+\w+)?\s*$', '', s, flags=re.IGNORECASE).strip()
    s = re.sub(r'\s+Harbor\s+Entrance\s*$', '', s, flags=re.IGNORECASE).strip()
    s = re.sub(r'\s+NMS\s*$', '', s, flags=re.IGNORECASE).strip()
    s = re.sub(r'\s+Tower\s*$', '', s, flags=re.IGNORECASE).strip()

    if len(s) > max_len:
        s_alt = re.sub(r'\s+Beach\s*$', '', s, flags=re.IGNORECASE).strip()
        if s_alt and len(s_alt) < len(s):
            s = s_alt
    if len(s) > max_len:
        s_alt = re.sub(r'\s+Inlet\s*$', '', s, flags=re.IGNORECASE).strip()
        if s_alt and len(s_alt) < len(s):
            s = s_alt

    if len(s) > max_len:
        s = s[:max_len-2].rstrip() + '..'

    return s


def timeline_strip_png_bytes(legs, output_w_px=1200, output_h_px=460):
    """Horizontal timeline showing wind speed (top panel) and sea height
    (bottom panel) across the passage, with shared x-axis.

    Refactored 6/21/26 from single-axes / twinx layout to stacked subplots.
    The two-scales-on-one-chart pattern was crowding the data; stacking lets
    each series have its own y-scale and breathing room while keeping WPs
    aligned vertically so the skipper can read wind and sea at each
    waypoint by just running their eye down the page.

    Color story (intuitive):
      - Sea = blue (water)
      - Wind sustained = warm peach/orange
      - Wind gusts = dark rust (still wind family but more prominent)
      - SCA threshold (18 kt) = amber dashed reference line
      - Reef threshold (25 kt) = red dashed reference line

    legs: list of Leg objects with attributes wp_id, cum_time_hr,
          wind_kt_high, sea_ft_high, gust_kt, eta_str
    """
    if not legs:
        return None

    # Pull series; default missing values to 0 so plotting doesn't break
    wind = [l.wind_kt_high if l.wind_kt_high else 0 for l in legs]
    gusts = [l.gust_kt if l.gust_kt else 0 for l in legs]
    seas = [l.sea_ft_high if l.sea_ft_high else 0 for l in legs]
    wp_labels = [l.wp_id for l in legs]

    # Two-line tick labels: WP id on top, short city/location label below.
    # Prefer the explicit chart_label from YAML if present; otherwise derive
    # from wp_name via the heuristic. Falls back to just WP id if neither
    # produces a non-empty string.
    tick_labels = []
    for leg in legs:
        loc = (getattr(leg, "chart_label", "") or "").strip()
        if not loc:
            loc = _short_location_from_wp_name(getattr(leg, "wp_name", "") or "")
        if loc:
            tick_labels.append(f"{leg.wp_id}\n{loc}")
        else:
            tick_labels.append(leg.wp_id)

    # Two stacked panels sharing x-axis. Top is taller because wind has
    # more vertical content (bars + gust caps + two threshold lines + labels).
    fig, (ax_wind, ax_sea) = plt.subplots(
        2, 1, figsize=(output_w_px / 100, output_h_px / 100), dpi=100,
        sharex=True, gridspec_kw={"height_ratios": [2.0, 1.0], "hspace": 0.08},
    )
    fig.subplots_adjust(left=0.06, right=0.96, top=0.90, bottom=0.10)

    bar_width = 0.55
    x = np.arange(len(wp_labels))

    # ======================================================================
    # TOP PANEL — Wind sustained bars + gust caps + threshold reference lines
    # ======================================================================
    ax_wind.bar(x, wind, width=bar_width, color=COLOR_WIND_BAR,
                edgecolor=COLOR_WIND_EDGE, linewidth=1.4,
                label="Wind sustained (kt)", zorder=3)
    # Wind speed labels INSIDE the bars (clearer than below)
    for xi, wi in zip(x, wind):
        if wi > 0:
            ax_wind.annotate(f"{wi:g}", xy=(xi, wi / 2), ha="center", va="center",
                             fontsize=11, color=COLOR_WIND_EDGE, fontweight="bold",
                             zorder=4)

    # Gust marks — dark rust horizontal line above the bar + label
    has_gusts = any(gu and gu > wi for wi, gu in zip(wind, gusts))
    for i, (wi, gu) in enumerate(zip(wind, gusts)):
        if gu and gu > wi:
            ax_wind.plot([x[i] - bar_width/2 - 0.06, x[i] + bar_width/2 + 0.06],
                         [gu, gu], color=COLOR_WIND_GUST, linewidth=2.6, zorder=5)
            ax_wind.plot([x[i], x[i]], [wi, gu], color=COLOR_WIND_GUST,
                         linewidth=1.2, linestyle=":", alpha=0.7, zorder=4)
            ax_wind.annotate(f"gust {gu:g}", xy=(x[i], gu), xytext=(0, 6),
                             textcoords="offset points", ha="center",
                             fontsize=10, color=COLOR_WIND_GUST, fontweight="bold")

    # Synthetic legend handle for the gust marks
    gust_handle = None
    if has_gusts:
        from matplotlib.lines import Line2D
        gust_handle = Line2D([0], [0], color=COLOR_WIND_GUST, linewidth=2.6,
                             label="Wind gust (kt)")

    # Threshold reference lines — SCA at 18, reef at 25
    ax_wind.axhline(y=18, color=COLOR_SCA_THRESHOLD, linewidth=1.4, linestyle="--",
                    alpha=0.8, zorder=2)
    ax_wind.text(len(x) - 0.4, 18.6, "SCA threshold 18 kt", fontsize=9,
                 color="#9C5700", ha="right", va="bottom", fontweight="bold")
    ax_wind.axhline(y=25, color=COLOR_REEF_THRESHOLD, linewidth=1.4, linestyle="--",
                    alpha=0.7, zorder=2)
    ax_wind.text(len(x) - 0.4, 25.6, "Reef trigger 25 kt", fontsize=9,
                 color=COLOR_REEF_THRESHOLD, ha="right", va="bottom", fontweight="bold")

    # Wind axis styling
    ax_wind.set_ylabel("Wind speed (kt)", fontsize=11, color=COLOR_WIND_EDGE,
                      fontweight="bold")
    ax_wind.set_ylim(0, max(32, (max(wind + gusts + [0]) * 1.25)))
    ax_wind.tick_params(axis="y", labelcolor=COLOR_WIND_EDGE, labelsize=10)
    ax_wind.grid(axis="y", color=COLOR_GRID, linewidth=0.5, alpha=0.5, zorder=1)
    ax_wind.set_axisbelow(True)

    # Wind-panel legend (just wind series + gust handle if present)
    wind_handles, wind_labels = ax_wind.get_legend_handles_labels()
    if gust_handle is not None:
        wind_handles.append(gust_handle)
        wind_labels.append("Wind gust (kt)")
    ax_wind.legend(wind_handles, wind_labels, loc="upper left", fontsize=10,
                   frameon=True, edgecolor=COLOR_GRID, framealpha=0.95,
                   ncol=len(wind_labels))

    # Title on the top panel
    ax_wind.set_title("Wind & Sea across passage", fontsize=12,
                      color=COLOR_TITLE, fontweight="bold", loc="left", pad=8)

    # ======================================================================
    # BOTTOM PANEL — Sea height line
    # ======================================================================
    ax_sea.plot(x, seas, color=COLOR_SEA_LINE, linewidth=3, marker="o",
                markersize=10, markeredgecolor="white", markeredgewidth=1.5,
                label="Sea height (ft)", zorder=6)
    # Sea-height labels — placed above the marker (no longer competing with
    # gust callouts since they're in a different panel now)
    for xi, si in zip(x, seas):
        if si > 0:
            ax_sea.annotate(f"{si:g} ft", xy=(xi, si), xytext=(0, 8),
                            textcoords="offset points", ha="center",
                            fontsize=10, color=COLOR_SEA_LINE, fontweight="bold")

    ax_sea.set_ylabel("Sea height (ft)", fontsize=11, color=COLOR_SEA_LINE,
                     fontweight="bold")
    ax_sea.set_ylim(0, max(7, (max(seas + [0]) * 1.5)))
    ax_sea.tick_params(axis="y", labelcolor=COLOR_SEA_LINE, labelsize=10)
    ax_sea.grid(axis="y", color=COLOR_GRID, linewidth=0.5, alpha=0.5, zorder=1)
    ax_sea.set_axisbelow(True)

    # Sea-panel legend (just sea height)
    ax_sea.legend(loc="upper left", fontsize=10, frameon=True,
                  edgecolor=COLOR_GRID, framealpha=0.95)

    # ======================================================================
    # Shared x-axis — WP labels only on the bottom panel
    # Labels rotated 30° to prevent overlap of multi-word city names
    # (e.g. "Cape Canaveral" next to "Daytona Beach" doesn't fit horizontally
    # at 10 WPs in a 1200 px chart)
    # ======================================================================
    ax_sea.set_xticks(x)
    ax_sea.set_xticklabels(tick_labels, fontsize=10, color=COLOR_TITLE,
                           fontweight="bold", rotation=30, ha="right",
                           rotation_mode="anchor")
    ax_sea.tick_params(axis="x", labelsize=10)
    # Top panel hides its own x tick labels (sharex handles the axis values,
    # but we still want to suppress the tick label text on the top panel)
    plt.setp(ax_wind.get_xticklabels(), visible=False)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=100, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf


# ======================================================================
# Mini polar — small chart showing boat polar at this TWS with TWA marker
# ======================================================================
def mini_polar_png_bytes(tws, twa, design="D1170", output_px=200, tack="starboard"):
    """Tiny polar plot showing boat speed curve at given TWS,
    with a marker dot at the current TWA on the correct tack.

    tack: "starboard" → dot on the right side of the polar (wind from starboard)
          "port"      → dot on the left side  (wind from port)

    The polar curve itself is mirrored on both sides since the boat performs
    symmetrically — but the dot indicates which tack the boat is actually on
    for this leg.
    """
    from .polar import polar_speed

    # Sample the polar from 30° to 180° in 5° steps
    angles_deg = np.arange(30, 181, 5)
    speeds = [polar_speed(tws, a, design) for a in angles_deg]

    # Convert to radians for the polar plot
    # The chart uses standard "wind from top" — 0° = into the wind
    angles_rad = np.radians(angles_deg)

    fig, ax = plt.subplots(figsize=(output_px / 100, output_px / 100),
                           subplot_kw=dict(polar=True), dpi=100)

    # Plot the polar curve (mirrored across the vertical axis to show both sides)
    # We use 0° = into the wind, going clockwise
    ax.set_theta_zero_location("N")
    ax.set_theta_direction(-1)  # Clockwise

    ax.plot(angles_rad, speeds, color=COLOR_PLAN_A, linewidth=2)
    # Mirror on the other side
    ax.plot(-angles_rad, speeds, color=COLOR_PLAN_A, linewidth=2)
    ax.fill(np.concatenate([angles_rad, -angles_rad[::-1]]),
            np.concatenate([speeds, speeds[::-1]]),
            color=COLOR_PLAN_A, alpha=0.15)

    # Current TWA marker — place on the correct side based on tack.
    # In matplotlib polar with theta_zero=N, direction=-1:
    #   positive theta goes clockwise (right = starboard)
    #   negative theta goes counterclockwise (left = port)
    twa_rad = math.radians(twa)
    if tack == "port":
        marker_theta = -twa_rad  # Mirror to the left side
    else:
        marker_theta = twa_rad
    current_speed = polar_speed(tws, twa, design)
    ax.plot(marker_theta, current_speed, marker="o", markersize=9,
            markerfacecolor=COLOR_BAD, markeredgecolor="white",
            markeredgewidth=1.8, zorder=10)

    # Tick labels at the cardinal points of the polar.
    # Show angle labels on BOTH sides so the dot's side is unambiguous.
    ax.set_xticks(np.radians([0, 90, 180, 270]))
    ax.set_xticklabels(["0°", "90°\nstbd", "180°", "90°\nport"], fontsize=7,
                       color="#595959")
    ax.set_yticks([4, 8])
    ax.set_yticklabels(["4", "8"], fontsize=7, color="#595959")
    ax.set_ylim(0, max(speeds) * 1.15)
    ax.grid(color=COLOR_GRID, linewidth=0.5, alpha=0.5)
    ax.spines["polar"].set_color(COLOR_GRID)
    ax.spines["polar"].set_linewidth(0.5)
    ax.set_title(f"TWS {tws} kt", fontsize=8, color=COLOR_TITLE,
                 fontweight="bold", pad=10)

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=100, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf
