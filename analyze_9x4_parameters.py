#!/usr/bin/env python3
"""
analyze_9x4_parameters.py
=========================
Back-calculates New Glenn 9x4 stage parameters from Blue Origin's published
performance claims using a 3-constraint system:

  1. LEO payload  > 70,000 kg   (reusable S1)
  2. TLI payload  > 20,000 kg   (reusable S1)
  3. Height geometry: stage tank lengths must fit within the published ~120m vehicle

The GTO > 14,000 kg claim is not an independent constraint — it is automatically
satisfied when the TLI constraint is met (TLI requires more ΔV than GTO at the
same loss budget, but with lighter payload TLI naturally has more total stage ΔV).

All engine performance data is sourced from publicly available specifications.
Propellant masses and loss model parameters are derived analytically from the
above constraints. Results represent our best estimates given available information.

Usage:
    python3 analyze_9x4_parameters.py

Output:
    NG_9x4_Parameter_Analysis.xlsx  — documented analysis workbook
"""

import math
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from scipy.optimize import fsolve

# ── Import shared infrastructure from main script ─────────────────────────────
import sys
sys.path.insert(0, '.')
from generate_lv_design import (
    lookup_engine, blended_isp, G0, C,
    style_hdr, style_sec, apply_border_range,
)

# ── Physical constants ────────────────────────────────────────────────────────
MU_E  = 3.986004418e14   # m³/s²
R_E   = 6_371_000        # m
R_MOON = 384_400_000     # m  (mean Earth–Moon distance)

# ── Published Blue Origin claims (from public announcements) ──────────────────
# Sources: Blue Origin press releases, Wikipedia, Space Launch Schedule
TARGET_LEO_KG  = 70_000   # > 70,000 kg to LEO (reusable S1)
TARGET_GTO_KG  = 14_000   # > 14,000 kg to GTO (reusable S1) – NOT independent
TARGET_TLI_KG  = 20_000   # > 20,000 kg to TLI (reusable S1)
TARGET_HEIGHT_M = 120.0   # ≈ 120 m total vehicle height (public sources)

# ── Known engine parameters (Blue Origin public data) ─────────────────────────
eng1 = lookup_engine("BE-4 Block 2")   # 9 engines on S1
eng2 = lookup_engine("BE-3U Block 2")  # 4 engines on S2 (200,000 lbf version)
isp1 = blended_isp(eng1, 0.6)          # 327.4 s (60% atm fraction for S1)
isp2 = blended_isp(eng2, 1.0)          # 440.0 s (vacuum, S2)
vac1 = eng1["vac_isp"]                 # 339 s (for recovery prop calc)

# ── Structural model ──────────────────────────────────────────────────────────
# Dry mass = (structural mass scaling with propellant) + (engine system mass)
# Engine system mass = n_engines × engine_mass_from_DB × 1.15 (feed system factor)
# Structural mass (non-engine) calibrated from 7x2 published data
BE4_SYS   = 4760 * 1.15   # kg per BE-4 system
BE3U_SYS  = 2700 * 1.15   # kg per BE-3U system
S1_STRUCT_7X2 = 55_000 - 7 * BE4_SYS    # non-engine structural mass, 7x2 S1
S2_STRUCT_7X2 = 10_000 - 2 * BE3U_SYS   # non-engine structural mass, 7x2 S2

# ── Geometry constants (from 7x2 published dimensions) ───────────────────────
BODY_D   = 7.0             # m  core body diameter (8.7m is fairing, not core)
A_BODY   = math.pi * (BODY_D/2)**2   # 38.48 m²
S1_HT_7X2 = 57.5           # m  (Wikipedia, confirmed)
S2_HT_7X2 = 23.4           # m  (Wikipedia, confirmed)
NONSTAGE_HT = 98.0 - S1_HT_7X2 - S2_HT_7X2   # 17.1 m (fairing + adapters)
FAIRING_EXTRA = 5.0        # m  extra height for larger 8.7m fairing on 9x4

# 9x4 stage height budget
STAGE_BUDGET_9X4 = TARGET_HEIGHT_M - NONSTAGE_HT - FAIRING_EXTRA  # 97.9 m

# ── Bulk propellant densities ─────────────────────────────────────────────────
# Correct formula: rho_bulk = (O/F + 1) / (O/F/rho_ox + 1/rho_fuel)
def bulk_density(mr_of, rho_ox, rho_fuel):
    return (mr_of + 1) / (mr_of / rho_ox + 1 / rho_fuel)

# LOX/CH4 O/F calibrated from Wikipedia tank volumes (850 m³ LOX / 710 m³ CH4)
# Volume ratio = (m_lox/rho_lox) / (m_ch4/rho_ch4) = OF × rho_ch4/rho_lox
# 850/710 = OF × 423/1141  →  OF = 1.197 × 1141/423 = 3.23
RHO_CH4_BULK = bulk_density(3.23, 1141, 423)   # ≈ 814 kg/m³
RHO_LH2_BULK = bulk_density(5.50, 1141, 71)    # ≈ 344 kg/m³

# Tank fill fraction from 7x2 calibration
# Non-tank height fraction = 1 - (prop_volume / (A × stage_height))
FILL_FRAC = 0.95  # 95% usable tank volume (5% ullage/margins)

def prop_to_length(prop_kg, bulk_rho):
    """Convert propellant mass to cylindrical tank length at BODY_D diameter."""
    return prop_kg / (bulk_rho * A_BODY * FILL_FRAC)

# Calibrate non-tank fraction from 7x2 actual stage heights and our model prop masses
_L_s1_7x2 = prop_to_length(770_000, RHO_CH4_BULK)
_L_s2_7x2 = prop_to_length( 75_000, RHO_LH2_BULK)
S1_NONTANK_FRAC = (_L_s1_7x2 / S1_HT_7X2)   # proportion of S1 height occupied by tanks
S2_NONTANK_FRAC = (_L_s2_7x2 / S2_HT_7X2)   # proportion of S2 height occupied by tanks


# ── Core physics functions ────────────────────────────────────────────────────

def dry_masses(s1p, s2p):
    """Structural + engine system dry masses, scaling structural with propellant."""
    s1d = S1_STRUCT_7X2 * (s1p / 770_000) + 9 * BE4_SYS
    s2d = S2_STRUCT_7X2 * (s2p /  75_000) + 4 * BE3U_SYS
    return s1d, s2d

def total_stage_dv(s1p, s2p, payload_kg, rec_dv_ms=300):
    """Total stage ΔV (S1+S2) via direct Tsiolkovsky for given payload."""
    s1d, s2d = dry_masses(s1p, s2p)
    eff1 = s1p * 1.035                                     # 3.5% subcooling
    rec1 = s1d * (math.exp(rec_dv_ms / (vac1 * G0)) - 1)  # recovery propellant
    pl   = payload_kg + 200                                # payload + adapter
    s2g  = s2p + s2d
    dv2  = isp2 * G0 * math.log((s2p + s2d + pl) / (s2d + pl))
    dv1  = isp1 * G0 * math.log(
               (eff1 + s1d + s2g + pl) / ((s1d + rec1) + s2g + pl))
    return dv1 + dv2, dv1, dv2

def trajectory_losses(s1p, s2p, traj_pen, grav_coeff=1400, drag=120):
    """Trajectory losses: gravity (from TWR) + trajectory penalty + drag."""
    s1d, s2d = dry_masses(s1p, s2p)
    glow = s1p + s1d + s2p + s2d
    twr  = 9 * eng1["sl_thrust_kn"] * 1000 / (glow * G0)
    return grav_coeff / twr + traj_pen + drag, twr

def v_inject(s1p, s2p, payload_kg, traj_pen, grav_coeff=1400, **kw):
    tdv, dv1, dv2 = total_stage_dv(s1p, s2p, payload_kg, **kw)
    loss, twr = trajectory_losses(s1p, s2p, traj_pen, grav_coeff=grav_coeff)
    return tdv - loss, tdv, dv1, dv2, loss, twr

def apogee_km(v_inj_ms, h_perigee_km=300):
    """Maximum apogee altitude from direct perigee injection. vis-viva."""
    r_p = R_E + h_perigee_km * 1000
    denom = 2 * MU_E / r_p - v_inj_ms**2
    if denom <= 0:
        return float('inf')
    a = MU_E / denom
    return (2 * a - r_p - R_E) / 1000


# ── Orbital targets ───────────────────────────────────────────────────────────

def v_tli_at(h_perigee_km=300):
    r_p = R_E + h_perigee_km * 1000
    a   = (r_p + R_MOON) / 2
    return math.sqrt(MU_E * (2/r_p - 1/a))

V_LEO = 7_784   # m/s  target LEO orbital velocity (400 km)
V_GTO = 10_400  # m/s  GTO perigee injection (standard, ~10,200–10,500 m/s)
V_TLI = v_tli_at(300)  # ≈ 10,838 m/s at 300 km perigee


# ── Geometry max-prop calculations ───────────────────────────────────────────

def stage_heights_from_props(s1p, s2p):
    """Estimated stage heights for given propellant masses."""
    L_s1 = prop_to_length(s1p, RHO_CH4_BULK)
    L_s2 = prop_to_length(s2p, RHO_LH2_BULK)
    h_s1 = L_s1 / S1_NONTANK_FRAC
    h_s2 = L_s2 / S2_NONTANK_FRAC
    return h_s1, h_s2


# ── 3-constraint solver ───────────────────────────────────────────────────────

def solver_residuals(x):
    s1p, s2p, traj_pen = x
    if s1p < 400_000 or s2p < 50_000 or traj_pen < 0 or traj_pen > 1500:
        return [1e9, 1e9, 1e9]
    if s1p > 2_500_000 or s2p > 500_000:
        return [1e9, 1e9, 1e9]

    loss, _ = trajectory_losses(s1p, s2p, traj_pen)
    dv_leo, *_ = total_stage_dv(s1p, s2p, TARGET_LEO_KG)
    dv_tli, *_ = total_stage_dv(s1p, s2p, TARGET_TLI_KG)
    h_s1, h_s2 = stage_heights_from_props(s1p, s2p)

    r1 = dv_leo - loss - V_LEO          # LEO constraint
    r2 = dv_tli - loss - V_TLI          # TLI constraint
    r3 = (h_s1 + h_s2) - STAGE_BUDGET_9X4  # height constraint
    return [r1, r2, r3]

def run_solver():
    """Try multiple starting points; return best solutions."""
    solutions = []
    for s1_init in [770_000, 900_000, 1_050_000, 1_200_000, 1_400_000]:
        for s2_init in [75_000, 120_000, 170_000, 250_000]:
            for pen_init in [100, 250, 400, 600]:
                try:
                    sol = fsolve(solver_residuals, [s1_init, s2_init, pen_init],
                                 full_output=True, xtol=1e-5, maxfev=10_000)
                    x, _, ier, _ = sol
                    if ier == 1 and all(v > 0 for v in x):
                        res = max(abs(r) for r in solver_residuals(x))
                        if res < 10.0:
                            solutions.append((res, x.copy()))
                except Exception:
                    pass
    # Deduplicate
    solutions.sort()
    seen, best = set(), []
    for res, x in solutions:
        key = (round(x[0], -4), round(x[1], -3), round(x[2], 0))
        if key not in seen:
            seen.add(key)
            best.append((res, x))
    return best


# ── Sensitivity sweep ─────────────────────────────────────────────────────────

def sweep_performance(s1p, s2p, traj_pen):
    """Return performance at key payloads for a given parameter set."""
    rows = []
    for label, payload in [("LEO (70t claim)", 70_000), ("GTO (14t claim)", 14_000),
                            ("TLI (20t claim)", 20_000), ("Blue Moon Mk1", 21_000),
                            ("Max (0t)", 0)]:
        vi, tdv, dv1, dv2, loss, twr = v_inject(s1p, s2p, payload, traj_pen)
        h_apo = apogee_km(vi)
        rows.append((label, payload, tdv, vi, dv1, dv2, loss, twr, h_apo))
    return rows


# ── Excel workbook builder ────────────────────────────────────────────────────

def build_workbook(solutions):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_assumptions(wb)
    _build_geometry(wb)
    _build_calibration(wb)
    _build_solver_results(wb, solutions)
    _build_sensitivity(wb, solutions)

    tab_colors = {
        "Assumptions & Method": "1F3864",
        "Geometry Analysis":    "2E7D32",
        "7x2 Calibration":      "BF360C",
        "Solver Results":       "4527A0",
        "Sensitivity Analysis": "00695C",
    }
    for ws in wb.worksheets:
        if ws.title in tab_colors:
            ws.sheet_properties.tabColor = tab_colors[ws.title]

    return wb


def _wr(ws, r, label, unit, value, bold=False):
    """Write one label/unit/value row."""
    fg = C["pos"] if bold else C["ltgray"]
    c1 = ws.cell(r, 1, value=label)
    c1.font = Font(bold=bold, size=10); c1.fill = PatternFill("solid", fgColor=fg)
    c1.alignment = Alignment(horizontal="left", indent=1)
    c2 = ws.cell(r, 2, value=unit)
    c2.font = Font(size=9, color="777777"); c2.fill = PatternFill("solid", fgColor=fg)
    c2.alignment = Alignment(horizontal="center")
    c3 = ws.cell(r, 3, value=round(value, 2) if isinstance(value, float) else value)
    c3.font = Font(bold=bold, size=10, italic=True)
    c3.fill = PatternFill("solid", fgColor=C["pos"] if bold else C["calc"])
    c3.alignment = Alignment(horizontal="center")
    c3.number_format = "#,##0" if unit in ("kg", "m/s", "m", "t") else "0.00"


def _build_assumptions(wb):
    ws = wb.create_sheet("Assumptions & Method")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 80

    r = 1
    style_hdr(ws, r, 1, "NG 9x4 PARAMETER BACK-CALCULATION — ASSUMPTIONS AND METHOD", span=1)
    ws.row_dimensions[r].height = 26
    r += 1

    lines = [
        ("OBJECTIVE", True),
        ("Determine the most likely stage propellant masses and trajectory loss parameters for",  False),
        ("New Glenn 9x4 by inverting the published performance claims against a physics model.", False),
        ("All inputs are publicly available data or documented engineering assumptions.", False),
        ("", False),
        ("PUBLISHED PERFORMANCE CLAIMS (Blue Origin, public announcements)", True),
        (f"  LEO payload:   > {TARGET_LEO_KG:,} kg  (reusable S1, drone-ship landing)", False),
        (f"  GTO payload:   > {TARGET_GTO_KG:,} kg  (reusable S1)", False),
        (f"  TLI payload:   > {TARGET_TLI_KG:,} kg  (reusable S1)", False),
        (f"  Vehicle height: ≈ {TARGET_HEIGHT_M:.0f} m  (public imagery / press releases)", False),
        ("", False),
        ("INDEPENDENT CONSTRAINTS USED IN SOLVER", True),
        ("  1. LEO 70t  — direct injection to circular orbit at v_orbit = 7,784 m/s", False),
        ("  2. TLI 20t  — direct injection to trans-lunar trajectory at v_TLI ≈ 10,838 m/s", False),
        ("  NOTE: GTO 14t is NOT independent — it is automatically satisfied if TLI is met,", False),
        ("  because TLI (harder mission) with more payload bounds the GTO (easier) constraint.", False),
        ("  3. Height geometry — stage tank lengths must sum to the 9x4 stage height budget", False),
        ("     (~97.9 m, derived from 120m total minus fairing and non-stage structure).", False),
        ("", False),
        ("FIXED PARAMETERS (from public data)", True),
        (f"  BE-4 Block 2 SL Isp:         {eng1['sl_isp']} s         (Blue Origin public spec)", False),
        (f"  BE-4 Block 2 Vac Isp:        {eng1['vac_isp']} s         (Blue Origin public spec)", False),
        (f"  BE-4 effective blended Isp:  {isp1:.1f} s      (60% atm fraction, S1 ascent profile)", False),
        (f"  BE-3U Block 2 Vac Isp:       {eng2['vac_isp']} s         (Blue Origin public spec; thrust upgraded to 220,000 lbf)", False),
        (f"  S1 subcooling gain:          3.5%        (Blue Origin public spec for BE-4 Block 2)", False),
        (f"  S1 recovery ΔV reserve:      300 m/s     (less conservative than 7x2 500 m/s; 9x4 drone-ship profile)", False),
        (f"  Body diameter (tanks):       {BODY_D} m         (core, confirmed; 8.7 m is fairing diameter only)", False),
        (f"  Gravity loss coefficient:    1,400       (lower than 7x2's 1,750; 9x4 higher TWR enables more aggressive trajectory)", False),
        (f"  Drag loss:                   120 m/s     (industry standard)", False),
        ("", False),
        ("SOLVED PARAMETERS (unknowns)", True),
        ("  S1 nominal propellant mass  — constrained by LEO + height equations", False),
        ("  S2 propellant mass          — constrained by TLI + height equations", False),
        ("  Trajectory penalty (m/s)    — additive gravity loss for S1 recovery pitch-over;", False),
        ("                                0 = expendable, 500 = aggressive recovery (RTLS),", False),
        ("                                200-300 typical estimate for drone-ship landing", False),
        ("", False),
        ("STRUCTURAL MODEL", True),
        ("  Dry mass = (structural mass proportional to propellant) + (engine system mass)", False),
        ("  Engine system mass = n_engines × engine_DB_mass × 1.15 (feed system factor)", False),
        ("  Structural (non-engine) mass: calibrated from 7x2 published dry masses", False),
        (f"    7x2 S1: structural = {S1_STRUCT_7X2:,.0f} kg  (55,000 − 7 × BE4_sys)", False),
        (f"    7x2 S2: structural = {S2_STRUCT_7X2:,.0f} kg  (10,000 − 2 × BE3U_sys)", False),
        ("", False),
        ("DISCLAIMER", True),
        ("  All analysis uses only publicly available information — published payload figures,", False),
        ("  announced engine performance data, official dimensions, and open-source references.", False),
        ("  No proprietary, export-controlled, or non-public data has been used.", False),
        ("  Results represent informed engineering estimates, not authoritative specifications.", False),
    ]
    for text, bold in lines:
        cell = ws.cell(r, 1, value=text)
        cell.font = Font(bold=bold, size=10 if not bold else 11,
                         color="1F3864" if bold else "1F1F1F")
        if bold:
            cell.fill = PatternFill("solid", fgColor=C["sec"])
        ws.row_dimensions[r].height = 16
        r += 1


def _build_geometry(wb):
    ws = wb.create_sheet("Geometry Analysis")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16

    r = 1
    style_hdr(ws, r, 1, "TANK GEOMETRY ANALYSIS — Propellant Volume vs Stage Height", span=3)
    ws.row_dimensions[r].height = 24; r += 1

    # Bulk densities
    style_sec(ws, r, 1, "BULK PROPELLANT DENSITIES  (correct volumetric formula)", span=3)
    ws.row_dimensions[r].height = 18; r += 1
    _wr(ws, r, "LOX/CH4 bulk density  (O/F = 3.23, calibrated from Wikipedia tank volumes)", "kg/m³", RHO_CH4_BULK); r += 1
    _wr(ws, r, "LOX/LH2 bulk density  (O/F = 5.5, standard)", "kg/m³", RHO_LH2_BULK); r += 1
    _wr(ws, r, "Body diameter (7m core — 8.7m is fairing only)", "m", BODY_D); r += 1
    _wr(ws, r, "Cross-section area", "m²", A_BODY); r += 1
    _wr(ws, r, "Tank fill fraction (ullage allowance)", "%", FILL_FRAC * 100); r += 2

    # 7x2 calibration
    style_sec(ws, r, 1, "7x2 GEOMETRY CALIBRATION", span=3)
    ws.row_dimensions[r].height = 18; r += 1
    _wr(ws, r, "S1 height (Wikipedia, confirmed)", "m", S1_HT_7X2); r += 1
    _wr(ws, r, "S2 height (Wikipedia, confirmed)", "m", S2_HT_7X2); r += 1
    _wr(ws, r, "Non-stage height (fairing + adapters)", "m", NONSTAGE_HT); r += 1
    L1 = prop_to_length(770_000, RHO_CH4_BULK)
    L2 = prop_to_length( 75_000, RHO_LH2_BULK)
    _wr(ws, r, "S1 tank length for 770t prop  (model estimate)", "m", L1); r += 1
    _wr(ws, r, "S2 tank length for 75t prop  (model estimate)", "m", L2); r += 1
    _wr(ws, r, "S1 tank fraction of stage height", "%", L1 / S1_HT_7X2 * 100); r += 1
    _wr(ws, r, "S2 tank fraction of stage height", "%", L2 / S2_HT_7X2 * 100); r += 2

    note = ws.cell(r, 1,
        "Note: low tank fractions reflect that stage height includes engine section, interstage,"
        " skirts, and avionics — not just propellant tanks.")
    note.font = Font(size=9, italic=True, color="555555")
    ws.merge_cells(f"A{r}:C{r}"); r += 2

    # 9x4 geometry
    style_sec(ws, r, 1, "9x4 GEOMETRY CONSTRAINTS", span=3)
    ws.row_dimensions[r].height = 18; r += 1
    _wr(ws, r, "Total vehicle height (public sources)", "m", TARGET_HEIGHT_M); r += 1
    _wr(ws, r, "Non-stage height (fairing now 8.7m, +5m vs 7x2)", "m", NONSTAGE_HT + FAIRING_EXTRA); r += 1
    _wr(ws, r, "Total stage height budget", "m", STAGE_BUDGET_9X4, bold=True); r += 1

    # Max prop vs height for different splits
    r += 1
    style_sec(ws, r, 1, "MAX PROPELLANT vs S1/S2 HEIGHT SPLIT  (7m body, 95% fill)", span=3)
    ws.row_dimensions[r].height = 18; r += 1
    for ci, h in enumerate(["S1 height (m)", "S2 height (m)", "Max S1 prop (t)",
                              "Max S2 prop (t)", "Stage total (m)"], 1):
        style_hdr(ws, r, ci, h, dark=False)
        ws.column_dimensions[get_column_letter(ci)].width = 16
    r += 1
    for s1_ht in [55, 60, 65, 70, 75, 80, 85]:
        s2_ht = STAGE_BUDGET_9X4 - s1_ht
        if s2_ht < 10: continue
        max_s1 = s1_ht * S1_NONTANK_FRAC * A_BODY * RHO_CH4_BULK * FILL_FRAC / 1000
        max_s2 = s2_ht * S2_NONTANK_FRAC * A_BODY * RHO_LH2_BULK * FILL_FRAC / 1000
        fg = C["ltgray"] if s1_ht % 10 == 0 else C["white"]
        for ci, v in enumerate([s1_ht, round(s2_ht, 1), round(max_s1, 0),
                                  round(max_s2, 0), round(s1_ht + s2_ht, 1)], 1):
            cell = ws.cell(r, ci, value=v)
            cell.fill = PatternFill("solid", fgColor=fg)
            cell.alignment = Alignment(horizontal="center")
            cell.font = Font(size=10)
            cell.number_format = "#,##0"
        apply_border_range(ws, r, r, 1, 5)
        ws.row_dimensions[r].height = 15; r += 1


def _build_calibration(wb):
    ws = wb.create_sheet("7x2 Calibration")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16

    r = 1
    style_hdr(ws, r, 1, "7x2 MODEL CALIBRATION — Verifying Against Published 45t LEO", span=3)
    ws.row_dimensions[r].height = 24; r += 1

    vi, tdv, dv1, dv2, loss, twr = v_inject(770_000, 75_000, 45_000, traj_pen=500, grav_coeff=1750, rec_dv_ms=500)
    s1d, s2d = dry_masses(770_000, 75_000)

    style_sec(ws, r, 1, "7x2 STAGE PARAMETERS  (from Vehicle Design sheet defaults)", span=3)
    ws.row_dimensions[r].height = 18; r += 1
    _wr(ws, r, "S1 nominal propellant", "kg", 770_000); r += 1
    _wr(ws, r, "S1 effective propellant (3.5% subcooling)", "kg", 770_000 * 1.035); r += 1
    _wr(ws, r, "S1 dry mass", "kg", s1d); r += 1
    _wr(ws, r, "S1 effective Isp (blended, 60% atm)", "s", isp1); r += 1
    _wr(ws, r, "S2 propellant", "kg", 75_000); r += 1
    _wr(ws, r, "S2 dry mass", "kg", s2d); r += 1
    _wr(ws, r, "S2 Isp (vacuum)", "s", isp2); r += 2

    style_sec(ws, r, 1, "PERFORMANCE WITH 45t PAYLOAD (direct Tsiolkovsky)", span=3)
    ws.row_dimensions[r].height = 18; r += 1
    _wr(ws, r, "S1 ΔV", "m/s", dv1); r += 1
    _wr(ws, r, "S2 ΔV", "m/s", dv2); r += 1
    _wr(ws, r, "Total stage ΔV", "m/s", tdv, bold=True); r += 1
    _wr(ws, r, "Liftoff TWR (stage masses only)", "—", twr); r += 1
    _wr(ws, r, "Gravity loss  (1750/TWR)", "m/s", 1750/twr); r += 1
    _wr(ws, r, "Trajectory penalty (drone-ship recovery estimate)", "m/s", 500); r += 1
    _wr(ws, r, "Drag loss", "m/s", 120); r += 1
    _wr(ws, r, "Total losses", "m/s", loss, bold=True); r += 1
    _wr(ws, r, "v_inject  (= total ΔV − losses)", "m/s", vi); r += 1
    _wr(ws, r, "LEO orbital velocity (400 km)", "m/s", 7_784); r += 1
    margin = vi - 7_784
    _wr(ws, r, "ΔV margin (positive = can reach orbit)", "m/s", margin, bold=(margin >= 0)); r += 2

    note = ws.cell(r, 1,
        "The 7x2 model is calibrated to match the published 45t LEO payload via the ΔV "
        "fraction heuristic in the design sheet. The direct Tsiolkovsky approach used here "
        "gives slightly different max payload because the fraction heuristic does not constrain "
        "the S1 Tsiolkovsky equation independently. Both approaches are documented for transparency.")
    note.font = Font(size=9, italic=True, color="555555")
    note.fill = PatternFill("solid", fgColor=C["ltgray"])
    note.alignment = Alignment(horizontal="left", indent=1, wrap_text=True)
    ws.merge_cells(f"A{r}:C{r}")
    ws.row_dimensions[r].height = 40


def _build_solver_results(wb, solutions):
    ws = wb.create_sheet("Solver Results")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 18

    r = 1
    style_hdr(ws, r, 1, "SOLVER RESULTS — Back-Calculated 9x4 Parameters", span=3)
    ws.row_dimensions[r].height = 24; r += 1

    style_sec(ws, r, 1, "CONSTRAINTS USED", span=3)
    ws.row_dimensions[r].height = 18; r += 1
    _wr(ws, r, f"LEO target (70t payload → v_inject = {V_LEO:,} m/s)", "m/s", V_LEO); r += 1
    _wr(ws, r, f"TLI target (20t payload → v_inject = {V_TLI:,.1f} m/s)", "m/s", V_TLI); r += 1
    _wr(ws, r, f"Stage height budget (from {TARGET_HEIGHT_M:.0f} m total vehicle)", "m", STAGE_BUDGET_9X4); r += 2

    if not solutions:
        cell = ws.cell(r, 1, value="⚠  No converged solution found with current model assumptions.")
        cell.font = Font(bold=True, size=11, color="CC0000")
        cell.fill = PatternFill("solid", fgColor=C["neg"])
        ws.merge_cells(f"A{r}:C{r}"); r += 2

        cell = ws.cell(r, 1,
            "This indicates the three constraints (LEO 70t, TLI 20t, height 120m) cannot be "
            "simultaneously satisfied with the current model assumptions (7m body diameter, "
            "Isp = 327.4/440 s, grav_coeff = 1400, drag = 120 m/s). One or more of the "
            "following must differ from our assumptions:\n"
            "  • Trajectory penalty significantly lower than 500 m/s for reusable 9x4\n"
            "  • BE-3U Block 2 Isp higher than published 440 s (expander cycle can improve with thrust)\n"
            "  • BE-4 Block 2 effective Isp higher than 327.4 s (unannounced improvement)\n"
            "  • Stage height ratio biased toward S2 beyond what geometry implies\n"
            "  • Core body diameter slightly larger than 7 m\n"
            "See Sensitivity Analysis sheet for the closest-approach parameter combinations.")
        cell.font = Font(size=10, italic=True)
        cell.fill = PatternFill("solid", fgColor=C["ltgray"])
        cell.alignment = Alignment(horizontal="left", indent=1, wrap_text=True)
        ws.merge_cells(f"A{r}:C{r}")
        ws.row_dimensions[r].height = 100
        r += 2
    else:
        for i, (res, x) in enumerate(solutions[:3], 1):
            s1p, s2p, pen = x
            s1d, s2d = dry_masses(s1p, s2p)
            glow = (s1p + s1d + s2p + s2d) / 1000
            loss, twr = trajectory_losses(s1p, s2p, pen)
            h_s1, h_s2 = stage_heights_from_props(s1p, s2p)

            style_sec(ws, r, 1, f"SOLUTION {i}  (residual = {res:.1f} m/s)", span=3)
            ws.row_dimensions[r].height = 18; r += 1
            _wr(ws, r, "S1 nominal propellant", "kg", s1p, bold=True); r += 1
            _wr(ws, r, "S1 dry mass (engineered)", "kg", s1d); r += 1
            _wr(ws, r, "S1 implied stage height", "m", h_s1); r += 1
            _wr(ws, r, "S2 propellant", "kg", s2p, bold=True); r += 1
            _wr(ws, r, "S2 dry mass (engineered)", "kg", s2d); r += 1
            _wr(ws, r, "S2 implied stage height", "m", h_s2); r += 1
            _wr(ws, r, "GLOW (no payload)", "t", glow, bold=True); r += 1
            _wr(ws, r, "Trajectory penalty", "m/s", pen, bold=True); r += 1
            _wr(ws, r, "Liftoff TWR", "—", twr); r += 1
            _wr(ws, r, "Total losses", "m/s", loss); r += 2

            for lbl, payload, v_tgt in [("LEO 70t", 70_000, V_LEO),
                                         ("GTO 14t", 14_000, V_GTO),
                                         ("TLI 20t", 20_000, V_TLI)]:
                vi, *_ = v_inject(s1p, s2p, payload, pen)
                ok = "✓" if vi >= v_tgt - 5 else f"✗ short {v_tgt-vi:.0f} m/s"
                _wr(ws, r, f"  {lbl}: v_inject = {vi:,.0f} m/s  (target {v_tgt:,.0f})  {ok}",
                    "m/s", vi); r += 1
            r += 1


def _build_sensitivity(wb, solutions):
    ws = wb.create_sheet("Sensitivity Analysis")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    for ci in range(2, 10):
        ws.column_dimensions[get_column_letter(ci)].width = 14

    r = 1
    style_hdr(ws, r, 1, "SENSITIVITY — What Parameters Allow TLI 20t Reusable?", span=8)
    ws.row_dimensions[r].height = 24; r += 1

    note = ws.cell(r, 1,
        "Sweep over trajectory penalty and S2 Isp to show what combinations "
        "allow the 9x4 (with height-constrained propellant masses) to reach TLI with 20t reusable. "
        "S1 and S2 prop masses held at geometry maximum for each cell's assumed stage split.")
    note.font = Font(size=9, italic=True, color="555555")
    note.fill = PatternFill("solid", fgColor=C["ltgray"])
    note.alignment = Alignment(horizontal="left", indent=1, wrap_text=True)
    ws.merge_cells(f"A{r}:H{r}")
    ws.row_dimensions[r].height = 36; r += 2

    # Use the best available prop masses (from solver or geometry max)
    if solutions:
        s1p_use, s2p_use, _ = solutions[0][1]
    else:
        # Geometry max with S1/S2 height split 70/30
        s1_ht = STAGE_BUDGET_9X4 * 0.70
        s2_ht = STAGE_BUDGET_9X4 * 0.30
        s1p_use = s1_ht * S1_NONTANK_FRAC * A_BODY * RHO_CH4_BULK * FILL_FRAC
        s2p_use = s2_ht * S2_NONTANK_FRAC * A_BODY * RHO_LH2_BULK * FILL_FRAC

    style_sec(ws, r, 1,
              f"v_inject for TLI with 20t payload   "
              f"(S1={s1p_use/1000:.0f}t, S2={s2p_use/1000:.0f}t)  "
              f"— target = {V_TLI:,.0f} m/s", span=8)
    ws.row_dimensions[r].height = 18; r += 1

    isp2_vals = [440, 445, 450, 455, 460]
    pen_vals  = [500, 400, 300, 200, 100, 0]

    # Header row
    ws.cell(r, 1, value="Traj penalty \\ S2 Isp →").font = Font(bold=True, size=10)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["sec"])
    for ci, isp in enumerate(isp2_vals, 2):
        style_hdr(ws, r, ci, f"{isp} s", dark=False)
    ws.row_dimensions[r].height = 22; r += 1

    for pen in pen_vals:
        ws.cell(r, 1, value=f"{pen} m/s").font = Font(size=10)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
        ws.cell(r, 1).alignment = Alignment(horizontal="center")
        for ci, isp2_val in enumerate(isp2_vals, 2):
            # Compute with overridden isp2
            s1d, s2d = dry_masses(s1p_use, s2p_use)
            eff1  = s1p_use * 1.035
            rec1  = s1d * (math.exp(300 / (vac1 * G0)) - 1)
            pl    = 20_000 + 200
            s2g   = s2p_use + s2d
            dv2   = isp2_val * G0 * math.log((s2p_use + s2d + pl) / (s2d + pl))
            dv1   = isp1 * G0 * math.log(
                        (eff1 + s1d + s2g + pl) / ((s1d + rec1) + s2g + pl))
            loss, _ = trajectory_losses(s1p_use, s2p_use, pen)
            vi_val  = dv1 + dv2 - loss

            deficit = vi_val - V_TLI
            if deficit >= 0:
                fg = C["pos"]
            elif deficit >= -500:
                fg = "FFE0B2"   # light orange: getting close
            else:
                fg = C["neg"]

            cell = ws.cell(r, ci, value=round(vi_val))
            cell.fill = PatternFill("solid", fgColor=fg)
            cell.alignment = Alignment(horizontal="center")
            cell.font = Font(size=10, bold=(deficit >= 0))
            cell.number_format = "#,##0"
        apply_border_range(ws, r, r, 1, len(isp2_vals) + 1)
        ws.row_dimensions[r].height = 15; r += 1

    r += 1
    note2 = ws.cell(r, 1,
        f"Green = achieves TLI 20t (v_inject ≥ {V_TLI:,.0f} m/s).  "
        "Orange = within 500 m/s of TLI target.  "
        "Red = does not reach TLI.")
    note2.font = Font(size=9, italic=True, color="444444")
    note2.fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.merge_cells(f"A{r}:H{r}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("NG 9x4 Parameter Back-Calculation")
    print("=" * 60)
    print(f"  TLI injection velocity target: {V_TLI:,.1f} m/s  (300 km perigee)")
    print(f"  Stage height budget:           {STAGE_BUDGET_9X4:.1f} m")
    print(f"  7x2 calibration: tank fracs   S1={S1_NONTANK_FRAC*100:.1f}%  S2={S2_NONTANK_FRAC*100:.1f}%")
    print()

    print("Running solver (LEO 70t + TLI 20t + height geometry)...")
    solutions = run_solver()

    if solutions:
        print(f"Found {len(solutions)} solution(s):\n")
        for i, (res, x) in enumerate(solutions[:3], 1):
            s1p, s2p, pen = x
            print(f"  Solution {i} (residual {res:.1f} m/s):")
            print(f"    S1 prop = {s1p/1000:.0f} t,  S2 prop = {s2p/1000:.0f} t,  penalty = {pen:.0f} m/s")
            vi_tli, *_ = v_inject(s1p, s2p, 20_000, pen)
            vi_mk1, *_ = v_inject(s1p, s2p, 21_000, pen)
            h_apo = apogee_km(vi_mk1)
            print(f"    TLI v_inject = {vi_tli:,.0f} m/s,  Blue Moon apogee = {h_apo:,.0f} km")
    else:
        print("No converged solution. Publishing closest-approach sensitivity table.")
        print("See Sensitivity Analysis sheet for details.\n")

    print("Building analysis workbook...")
    wb = build_workbook(solutions)
    outfile = "NG_9x4_Parameter_Analysis.xlsx"
    wb.save(outfile)
    print(f"✓ Saved: {outfile}")
    return solutions


if __name__ == "__main__":
    main()
