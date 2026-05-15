"""
Course-up wind/sea compass rose for Plan tab WP rows.

Renders a small SVG rose that orients NORTH-UP and shows:
  - A black boat triangle in the center pointing toward course direction
  - Cardinal letters N / E / S / W around the perimeter
  - A solid BLUE arrow showing wind FROM direction (tail at perimeter, head inward)
  - A dashed GREEN arrow showing sea FROM direction (tail at perimeter, head inward,
    offset perpendicular to wind arrow so they don't overlap when colinear)

The rose is purely a visual orientation aid — wind/sea numeric data already lives
in adjacent columns G/H/I/J/K/N. So the rose itself has no captions inside it.

Public API:
  make_rose_svg(course_deg, wind_from_deg, sea_from_deg) -> str
  rose_png_bytes(course_deg, wind_from_deg, sea_from_deg, output_px=160) -> BytesIO
  HAVE_CAIROSVG -> bool

If cairosvg isn't installed, rose_png_bytes raises RuntimeError. The caller
(_embed_rose in tabs/plan.py) wraps this in try/except and degrades to "—".
"""
import io
import math

try:
    import cairosvg
    HAVE_CAIROSVG = True
except ImportError:
    HAVE_CAIROSVG = False

# Canvas geometry — keep these small constants in one place
_CX = 70           # center x
_CY = 70           # center y
_R_OUTER = 55      # outer ring radius
_R_LABEL = 62      # cardinal label radius
_R_ARROW_TAIL = 52 # where arrow tails start (just inside outer ring)
_R_ARROW_HEAD = 16 # where arrow heads end (just outside boat triangle)
_PERP_OFFSET = 6   # perpendicular offset for sea arrow to avoid overlap with wind


def _arrow_points(from_deg_rel: float, r_tail: float, r_head: float, perp_offset: float):
    """Compute (x1, y1, x2, y2) for an arrow whose TAIL sits on the perimeter at
    bearing `from_deg_rel` from center, and whose HEAD points inward toward center.

    `from_deg_rel` is degrees clockwise from up (north-up convention after course
    rotation has been applied). `perp_offset` shifts the entire arrow sideways
    perpendicular to its axis — used to separate wind and sea arrows when they're
    colinear (e.g. wind from N, sea from N would otherwise overlap exactly).
    """
    theta = math.radians(from_deg_rel)
    # Tail at perimeter (further from center)
    tx = _CX + r_tail * math.sin(theta)
    ty = _CY - r_tail * math.cos(theta)
    # Head just outside the boat triangle (closer to center)
    hx = _CX + r_head * math.sin(theta)
    hy = _CY - r_head * math.cos(theta)
    # Apply perpendicular offset (rotate 90° from arrow direction)
    perp_theta = theta + math.pi / 2
    dx = perp_offset * math.sin(perp_theta)
    dy = -perp_offset * math.cos(perp_theta)
    return (tx + dx, ty + dy, hx + dx, hy + dy)


def make_rose_svg(course_deg: int, wind_from_deg: int, sea_from_deg: int) -> str:
    """Build the SVG string for the rose.

    COURSE-UP rendering: the boat triangle always points straight up to the top
    of the rose (it's "what you see looking forward"). Cardinal letters N/E/S/W
    rotate around the perimeter to show where compass directions sit RELATIVE to
    the boat's heading. Wind and sea arrows show their bearings relative to
    course — so a wind FROM dead ahead (TWA 0°) has its tail at the top of the
    rose; a wind from the port quarter (TWA 135° port) has its tail at the
    lower-left.

    All three angles are passed in compass degrees (0=N, 90=E, 180=S, 270=W);
    the function converts to relative angles internally.
    """
    # Convert absolute bearings to course-relative angles.
    # 0° relative = directly ahead of the boat (top of rose).
    wind_rel = (wind_from_deg - course_deg) % 360
    sea_rel = (sea_from_deg - course_deg) % 360

    parts = []
    parts.append(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 140 140" '
        'width="140" height="140">'
    )

    # Arrow-head marker definitions
    parts.append(
        '<defs>'
        '<marker id="wh" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" '
        'markerHeight="5" orient="auto-start-reverse">'
        '<path d="M0,0 L10,5 L0,10 z" fill="#1f4ed8"/>'
        '</marker>'
        '<marker id="sh" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" '
        'markerHeight="5" orient="auto-start-reverse">'
        '<path d="M0,0 L10,5 L0,10 z" fill="#0f7a30"/>'
        '</marker>'
        '</defs>'
    )

    # Outer ring (light gray)
    parts.append(
        f'<circle cx="{_CX}" cy="{_CY}" r="{_R_OUTER}" fill="#fafafa" '
        f'stroke="#888" stroke-width="1"/>'
    )

    # Cardinal tick marks at N/E/S/W positions (rotated to course-relative)
    for tick_compass_deg in (0, 90, 180, 270):
        tick_rel = (tick_compass_deg - course_deg) % 360
        t = math.radians(tick_rel)
        x1 = _CX + (_R_OUTER - 4) * math.sin(t)
        y1 = _CY - (_R_OUTER - 4) * math.cos(t)
        x2 = _CX + _R_OUTER * math.sin(t)
        y2 = _CY - _R_OUTER * math.cos(t)
        parts.append(
            f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
            f'stroke="#888" stroke-width="1.5"/>'
        )

    # Cardinal labels — placed at their course-relative positions so the
    # navigator can see where N is relative to where they're heading.
    for label, compass_deg in (("N", 0), ("E", 90), ("S", 180), ("W", 270)):
        rel = (compass_deg - course_deg) % 360
        t = math.radians(rel)
        lx = _CX + _R_LABEL * math.sin(t)
        ly = _CY - _R_LABEL * math.cos(t) + 4  # +4 for text baseline centering
        # N label gets emphasized (slightly bolder/darker) since orientation
        # to true north is the most common mental reference
        weight = "bold" if label == "N" else "normal"
        color = "#222" if label == "N" else "#666"
        parts.append(
            f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="middle" '
            f'font-family="Calibri, Arial, sans-serif" font-size="11" '
            f'font-weight="{weight}" fill="{color}">{label}</text>'
        )

    # Boat triangle ALWAYS points straight up (course-up).
    # No rotation applied — the boat is fixed at the center pointing toward
    # the top of the rose, which represents "ahead".
    parts.append(
        f'<path d="M{_CX:.1f},{_CY - 13:.1f} '
        f'L{_CX - 6:.1f},{_CY + 8:.1f} '
        f'L{_CX + 6:.1f},{_CY + 8:.1f} z" '
        f'fill="#222" stroke="#222" stroke-width="0.5"/>'
    )

    # Wind FROM arrow (solid blue): tail at perimeter at relative bearing,
    # head inward at r=16. perp_offset of +6 separates this from the sea arrow.
    wx1, wy1, wx2, wy2 = _arrow_points(
        wind_rel, _R_ARROW_TAIL, _R_ARROW_HEAD, +_PERP_OFFSET
    )
    parts.append(
        f'<line x1="{wx1:.1f}" y1="{wy1:.1f}" x2="{wx2:.1f}" y2="{wy2:.1f}" '
        f'stroke="#1f4ed8" stroke-width="2.8" marker-end="url(#wh)"/>'
    )

    # Sea FROM arrow (dashed green): tail at perimeter, head inward.
    # perp_offset of -6 separates from wind arrow.
    sx1, sy1, sx2, sy2 = _arrow_points(
        sea_rel, _R_ARROW_TAIL, _R_ARROW_HEAD, -_PERP_OFFSET
    )
    parts.append(
        f'<line x1="{sx1:.1f}" y1="{sy1:.1f}" x2="{sx2:.1f}" y2="{sy2:.1f}" '
        f'stroke="#0f7a30" stroke-width="2.5" stroke-dasharray="4,2" '
        f'marker-end="url(#sh)"/>'
    )

    # Course label at the top of the rose, dark gray, indicates the direction
    # the boat is pointing (which is straight up in course-up orientation).
    parts.append(
        f'<text x="{_CX}" y="14" text-anchor="middle" '
        f'font-family="Calibri, Arial, sans-serif" font-size="9" '
        f'fill="#555">{course_deg:03d}°T</text>'
    )

    parts.append('</svg>')
    return ''.join(parts)


def rose_png_bytes(course_deg: int, wind_from_deg: int, sea_from_deg: int,
                   output_px: int = 160) -> io.BytesIO:
    """Render the rose to a PNG in memory.

    Returns BytesIO seeked to 0, ready to pass to openpyxl.drawing.image.Image.
    Raises RuntimeError if cairosvg is not installed.
    """
    if not HAVE_CAIROSVG:
        raise RuntimeError(
            "cairosvg required for rose rendering. Install with: pip install cairosvg"
        )
    svg = make_rose_svg(course_deg, wind_from_deg, sea_from_deg)
    buf = io.BytesIO()
    cairosvg.svg2png(
        bytestring=svg.encode('utf-8'),
        write_to=buf,
        output_width=output_px,
        output_height=output_px,
    )
    buf.seek(0)
    return buf
