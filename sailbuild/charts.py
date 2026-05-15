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
def timeline_strip_png_bytes(legs, output_w_px=1200, output_h_px=340):
    """Horizontal timeline showing wind speed (bars) and sea height (line)
    across the passage.

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

    fig, ax1 = plt.subplots(figsize=(output_w_px / 100, output_h_px / 100), dpi=100)
    fig.subplots_adjust(left=0.06, right=0.94, top=0.85, bottom=0.16)

    bar_width = 0.55
    x = np.arange(len(wp_labels))

    # === Wind sustained bars — warm peach with darker orange edge ===
    ax1.bar(x, wind, width=bar_width, color=COLOR_WIND_BAR,
            edgecolor=COLOR_WIND_EDGE, linewidth=1.4,
            label="Wind sustained (kt)", zorder=3)
    # Wind speed labels INSIDE the bars (clearer than below)
    for xi, wi in zip(x, wind):
        if wi > 0:
            ax1.annotate(f"{wi:g}", xy=(xi, wi / 2), ha="center", va="center",
                         fontsize=11, color=COLOR_WIND_EDGE, fontweight="bold",
                         zorder=4)

    # === Gust marks — dark rust horizontal line above the bar + label ===
    has_gusts = any(gu and gu > wi for wi, gu in zip(wind, gusts))
    for i, (wi, gu) in enumerate(zip(wind, gusts)):
        if gu and gu > wi:
            # Horizontal cap mark at gust height
            ax1.plot([x[i] - bar_width/2 - 0.06, x[i] + bar_width/2 + 0.06],
                     [gu, gu], color=COLOR_WIND_GUST, linewidth=2.6, zorder=5)
            # Connecting tick lines from bar top to gust mark
            ax1.plot([x[i], x[i]], [wi, gu], color=COLOR_WIND_GUST,
                     linewidth=1.2, linestyle=":", alpha=0.7, zorder=4)
            ax1.annotate(f"gust {gu:g}", xy=(x[i], gu), xytext=(0, 6),
                         textcoords="offset points", ha="center",
                         fontsize=10, color=COLOR_WIND_GUST, fontweight="bold")

    # Synthetic legend handle for the gust marks (matplotlib doesn't auto-pick up the plot() calls)
    if has_gusts:
        from matplotlib.lines import Line2D
        gust_handle = Line2D([0], [0], color=COLOR_WIND_GUST, linewidth=2.6,
                             label="Wind gust (kt)")
    else:
        gust_handle = None

    # === Threshold reference lines — SCA at 18, reef at 25 ===
    ax1.axhline(y=18, color=COLOR_SCA_THRESHOLD, linewidth=1.4, linestyle="--",
                alpha=0.8, zorder=2)
    ax1.text(len(x) - 0.4, 18.6, "SCA threshold 18 kt", fontsize=9,
             color="#9C5700", ha="right", va="bottom", fontweight="bold")
    ax1.axhline(y=25, color=COLOR_REEF_THRESHOLD, linewidth=1.4, linestyle="--",
                alpha=0.7, zorder=2)
    ax1.text(len(x) - 0.4, 25.6, "Reef trigger 25 kt", fontsize=9,
             color=COLOR_REEF_THRESHOLD, ha="right", va="bottom", fontweight="bold")

    # === Wind axis styling ===
    ax1.set_xticks(x)
    ax1.set_xticklabels(wp_labels, fontsize=12, color=COLOR_TITLE, fontweight="bold")
    ax1.set_ylabel("Wind speed (kt)", fontsize=11, color=COLOR_WIND_EDGE,
                   fontweight="bold")
    ax1.set_ylim(0, max(32, (max(wind + gusts + [0]) * 1.25)))
    ax1.tick_params(axis="y", labelcolor=COLOR_WIND_EDGE, labelsize=10)
    ax1.tick_params(axis="x", labelsize=12)
    ax1.grid(axis="y", color=COLOR_GRID, linewidth=0.5, alpha=0.5, zorder=1)
    ax1.set_axisbelow(True)

    # === Sea height — water blue line on secondary axis ===
    ax2 = ax1.twinx()
    ax2.plot(x, seas, color=COLOR_SEA_LINE, linewidth=3, marker="o",
             markersize=10, markeredgecolor="white", markeredgewidth=1.5,
             label="Sea height (ft)", zorder=6)
    # Sea-height labels — placed below the marker so they don't collide
    # with gust callouts (which sit above the bars)
    for xi, si in zip(x, seas):
        if si > 0:
            ax2.annotate(f"{si:g} ft", xy=(xi, si), xytext=(0, -16),
                         textcoords="offset points", ha="center",
                         fontsize=10, color=COLOR_SEA_LINE, fontweight="bold")

    ax2.set_ylabel("Sea height (ft)", fontsize=11, color=COLOR_SEA_LINE,
                   fontweight="bold")
    ax2.set_ylim(0, max(9, (max(seas + [0]) * 1.6)))
    ax2.tick_params(axis="y", labelcolor=COLOR_SEA_LINE, labelsize=10)

    # === Combined legend ===
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    all_handles = h1 + ([gust_handle] if gust_handle else []) + h2
    all_labels = l1 + (["Wind gust (kt)"] if gust_handle else []) + l2
    ax1.legend(all_handles, all_labels, loc="upper left", fontsize=10,
               frameon=True, edgecolor=COLOR_GRID, framealpha=0.95,
               ncol=len(all_labels))

    # === Title ===
    ax1.set_title("Wind & Sea across passage", fontsize=12,
                  color=COLOR_TITLE, fontweight="bold", loc="left", pad=8)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=100, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    plt.close(fig)
    buf.seek(0)
    return buf


# ======================================================================
# Mini polar — small chart showing boat polar at this TWS with TWA marker
# ======================================================================
def mini_polar_png_bytes(tws, twa, design="D1170", output_px=200):
    """Tiny polar plot showing boat speed curve at given TWS,
    with a marker dot at the current TWA.
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

    # Current TWA marker (use whichever side is closer to 0-180)
    twa_rad = math.radians(twa)
    current_speed = polar_speed(tws, twa, design)
    ax.plot(twa_rad, current_speed, marker="o", markersize=9,
            markerfacecolor=COLOR_BAD, markeredgecolor="white",
            markeredgewidth=1.8, zorder=10)

    # Minimal styling for small size
    ax.set_xticks(np.radians([0, 45, 90, 135, 180]))
    ax.set_xticklabels(["0°", "45°", "90°", "135°", "180°"], fontsize=7,
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
