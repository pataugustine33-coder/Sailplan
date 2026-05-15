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
def timeline_strip_png_bytes(legs, output_w_px=1200, output_h_px=240):
    """Horizontal timeline showing wind speed (bars) and sea height (line)
    across the passage.

    legs: list of Leg objects with attributes wp_id, cum_time_hr,
          wind_kt_high, sea_ft_high, gust_kt, eta_str
    """
    if not legs:
        return None

    # Filter to legs with meaningful data (skip terminal WP without outbound metrics)
    times = [l.cum_time_hr for l in legs]
    wind = [l.wind_kt_high if l.wind_kt_high else 0 for l in legs]
    gusts = [l.gust_kt if l.gust_kt else 0 for l in legs]
    seas = [l.sea_ft_high if l.sea_ft_high else 0 for l in legs]
    wp_labels = [l.wp_id for l in legs]

    fig, ax1 = plt.subplots(figsize=(output_w_px / 100, output_h_px / 100), dpi=100)
    fig.subplots_adjust(left=0.05, right=0.95, top=0.85, bottom=0.20)

    # Wind speed bars (with gust overlay where present)
    bar_width = 0.6
    x = np.arange(len(times))

    # Sustained wind bars (lighter blue)
    bars = ax1.bar(x, wind, width=bar_width, color="#9DC3E6",
                   edgecolor=COLOR_PLAN_A, linewidth=1.0,
                   label="Wind (kt)", zorder=3)
    # Gust overlay (darker outline, slimmer bar inset)
    for i, (wi, gu) in enumerate(zip(wind, gusts)):
        if gu and gu > wi:
            ax1.plot([x[i] - bar_width/2 - 0.05, x[i] + bar_width/2 + 0.05],
                     [gu, gu], color=COLOR_BAD, linewidth=2.2, zorder=5)
            ax1.annotate(f"g{gu:g}", xy=(x[i], gu), xytext=(0, 4),
                         textcoords="offset points", ha="center",
                         fontsize=8, color=COLOR_BAD, fontweight="bold")

    # Wind threshold lines for context
    ax1.axhline(y=18, color=COLOR_NEUTRAL, linewidth=0.8, linestyle="--",
                alpha=0.6, zorder=2)
    ax1.text(len(x) - 0.5, 18.4, "SCA threshold", fontsize=7, color="#9C5700",
             ha="right", va="bottom")
    ax1.axhline(y=25, color=COLOR_BAD, linewidth=0.8, linestyle="--",
                alpha=0.6, zorder=2)

    ax1.set_xticks(x)
    ax1.set_xticklabels(wp_labels, fontsize=9, color=COLOR_TITLE, fontweight="bold")
    ax1.set_ylabel("Wind (kt)", fontsize=9, color=COLOR_PLAN_A, fontweight="bold")
    ax1.set_ylim(0, max(30, max(wind + gusts + [0]) * 1.15))
    ax1.tick_params(axis="y", labelcolor=COLOR_PLAN_A, labelsize=8)
    ax1.grid(axis="y", color=COLOR_GRID, linewidth=0.5, alpha=0.5, zorder=1)
    ax1.set_axisbelow(True)

    # Sea height as a line on secondary axis
    ax2 = ax1.twinx()
    ax2.plot(x, seas, color=COLOR_PLAN_B, linewidth=2.5, marker="o",
             markersize=7, markeredgecolor="white", markeredgewidth=1.2,
             label="Sea (ft)", zorder=6)
    # Sea-height labels
    for xi, si in zip(x, seas):
        if si > 0:
            ax2.annotate(f"{si:g}'", xy=(xi, si), xytext=(0, -14),
                         textcoords="offset points", ha="center",
                         fontsize=8, color=COLOR_PLAN_B, fontweight="bold")

    ax2.set_ylabel("Sea height (ft)", fontsize=9, color=COLOR_PLAN_B, fontweight="bold")
    ax2.set_ylim(0, max(8, max(seas + [0]) * 1.4))
    ax2.tick_params(axis="y", labelcolor=COLOR_PLAN_B, labelsize=8)

    # Combined legend
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, loc="upper left", fontsize=8,
               frameon=True, edgecolor=COLOR_GRID, framealpha=0.95)

    # Title
    ax1.set_title("Wind & Sea across passage", fontsize=10,
                  color=COLOR_TITLE, fontweight="bold", loc="left")

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
