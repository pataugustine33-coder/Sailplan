"""
Hallberg-Rassy polar grids (Frers VPP, half-load), with bilinear interpolation,
sea factor lookup, and apparent wind calculation.

This module is the boat-speed engine. Given (TWS, TWA, sea height, period,
calibration, design), it produces (polar speed, boat speed, sail mode, AWS, AWA).

Two design polars are wired in:
  - D1170 = HR 48 Mk II half-load (default for chs-beaufort etc.)
  - D1206 = HR 54 half-load (used by Vessel Comparison and any HR 54 passage)
"""
import math
import numpy as np


# Frers VPP polar table — TWS columns × TWA rows.
TWS_GRID = np.array([4, 6, 8, 10, 12, 14, 16, 20, 25])
TWA_GRID = np.array([45, 52, 60, 70, 80, 90, 100, 110, 120, 135, 150])

# HR 48 Mk II — D1170 half-load (the default working polar for the project).
VS_GRID_D1170 = np.array([
    [3.52, 5.08, 6.15, 6.85, 7.29, 7.56, 7.74, 7.96, 8.08],
    [3.98, 5.59, 6.63, 7.31, 7.69, 7.92, 8.04, 8.25, 8.43],
    [4.38, 5.98, 7.01, 7.63, 7.99, 8.20, 8.35, 8.56, 8.74],
    [4.69, 6.26, 7.26, 7.84, 8.21, 8.47, 8.63, 8.84, 9.04],
    [4.81, 6.37, 7.36, 7.93, 8.31, 8.60, 8.81, 9.07, 9.30],
    [4.77, 6.33, 7.33, 7.93, 8.32, 8.63, 8.88, 9.26, 9.54],
    [4.59, 6.18, 7.36, 8.00, 8.40, 8.66, 8.84, 9.30, 9.74],
    [4.41, 6.14, 7.28, 7.94, 8.38, 8.72, 8.96, 9.29, 9.76],
    [4.19, 5.88, 7.04, 7.78, 8.28, 8.66, 8.98, 9.48, 9.90],
    [3.51, 5.15, 6.38, 7.32, 7.92, 8.37, 8.74, 9.38, 10.12],
    [2.77, 4.20, 5.45, 6.44, 7.26, 7.83, 8.26, 8.93, 9.71],
])

# HR 54 — D1206 half-load. Source: /mnt/project/HR54speedtable.xls (Frers VPP).
VS_GRID_D1206 = np.array([
    [3.64, 5.26, 6.38, 7.10, 7.56, 7.85, 8.05, 8.28, 8.47],
    [4.12, 5.79, 6.87, 7.59, 7.99, 8.23, 8.37, 8.60, 8.80],
    [4.54, 6.20, 7.26, 7.93, 8.32, 8.54, 8.69, 8.92, 9.12],
    [4.85, 6.49, 7.53, 8.16, 8.56, 8.82, 8.99, 9.22, 9.43],
    [4.99, 6.60, 7.63, 8.25, 8.66, 8.97, 9.19, 9.46, 9.69],
    [4.95, 6.56, 7.60, 8.23, 8.66, 9.00, 9.26, 9.65, 9.92],
    [4.76, 6.39, 7.59, 8.30, 8.73, 9.02, 9.22, 9.69, 10.12],
    [4.51, 6.31, 7.51, 8.23, 8.71, 9.07, 9.33, 9.67, 10.14],
    [4.29, 6.05, 7.24, 8.04, 8.58, 9.00, 9.33, 9.84, 10.26],
    [3.57, 5.26, 6.54, 7.51, 8.17, 8.66, 9.06, 9.71, 10.46],
    [2.81, 4.25, 5.54, 6.57, 7.42, 8.05, 8.51, 9.23, 10.01],
])

# Design ID → polar grid lookup. Used by polar_speed(design=...).
POLARS = {
    "D1170": VS_GRID_D1170,
    "D1206": VS_GRID_D1206,
}


def polar_speed(tws: float, twa: float, design: str = "D1170") -> float:
    """Bilinear interpolation of pure polar speed at (TWS, TWA) for the given design.

    Clamps inputs to grid bounds — outside the grid we extrapolate to edge.
    `design` defaults to D1170 (HR 48) for backward compatibility with all
    existing call sites; pass "D1206" for HR 54 polar lookups.
    """
    grid = POLARS.get(design, VS_GRID_D1170)
    tws = max(TWS_GRID[0], min(TWS_GRID[-1], tws))
    twa = max(TWA_GRID[0], min(TWA_GRID[-1], twa))
    i_tws = max(0, min(len(TWS_GRID) - 2, np.searchsorted(TWS_GRID, tws) - 1))
    i_twa = max(0, min(len(TWA_GRID) - 2, np.searchsorted(TWA_GRID, twa) - 1))
    f_tws = (tws - TWS_GRID[i_tws]) / (TWS_GRID[i_tws + 1] - TWS_GRID[i_tws])
    f_twa = (twa - TWA_GRID[i_twa]) / (TWA_GRID[i_twa + 1] - TWA_GRID[i_twa])
    v00 = grid[i_twa, i_tws]
    v01 = grid[i_twa, i_tws + 1]
    v10 = grid[i_twa + 1, i_tws]
    v11 = grid[i_twa + 1, i_tws + 1]
    return ((1 - f_twa) * (1 - f_tws) * v00 +
            (1 - f_twa) * f_tws * v01 +
            f_twa * (1 - f_tws) * v10 +
            f_twa * f_tws * v11)


def select_sea_factor(twa: float, sea_ft: float, period_s: float, factors: dict) -> tuple[float, str]:
    """Choose a sea factor and return (value, label).

    Logic (per project calibration):
      - TWA < 60 with steep short-period chop (Hs/T >= 0.8, T < 6) → steep_chop_bow
      - TWA < 80                                                    → close_reach
      - TWA < 100                                                   → beam_reach
      - TWA < 150                                                   → broad_reach
      - TWA >= 150                                                  → broad_reach
    """
    if period_s and period_s > 0:
        hs_over_t = sea_ft / period_s
    else:
        hs_over_t = 0
    if twa < 60 and period_s < 6 and hs_over_t >= 0.8:
        return factors["steep_chop_bow"], "steep_chop_bow"
    if twa < 80:
        return factors["close_reach"], "close_reach"
    if twa < 100:
        return factors["beam_reach"], "beam_reach"
    return factors["broad_reach"], "broad_reach"


def apparent_wind(tws: float, twa_deg: float, bsp: float) -> tuple[float, float]:
    """Compute (AWS in kt, AWA in degrees 0-180)."""
    twa_rad = math.radians(twa_deg)
    fwd = tws * math.cos(twa_rad) + bsp
    lat = tws * math.sin(twa_rad)
    aws = math.sqrt(fwd ** 2 + lat ** 2)
    awa = abs(math.degrees(math.atan2(lat, fwd)))
    return aws, awa
