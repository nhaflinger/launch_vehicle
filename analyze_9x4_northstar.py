#!/usr/bin/env python3
"""
analyze_9x4_northstar.py
========================
Re-estimates New Glenn 9x4 performance with Blue Origin's new Northstar
first-stage engine (BE-4 derivative, ~715,000 lbf, ~30% lighter than BE-4).

Method — calibrate on what Blue Origin has actually flown, then extrapolate:

  1. CALIBRATE on the 7x2: 7 x BE-4 (550,000 lbf SL), 770t / 75t propellant,
     55t / 10t dry.  Solve for the single loss term that makes the model
     reproduce Blue Origin's claimed 45t to LEO (reusable).  Losses are split as
         loss = GRAV_COEFF / TWR_liftoff + K
     so only the gravity-loss part responds to the 9x4's different TWR.
     The 7x2's 13t GTO claim is used the same way to back out the effective
     "GTO" injection energy Blue Origin quotes against.

  2. ENGINE SWAP: hold the previous 9x4 tank estimate (847t / 111t) fixed and
     compare BE-4 Block 2 vs Northstar — isolates what the engine alone buys.

  3. BACK-SOLVE: find the S1/S2 propellant loads that best satisfy
     Blue Origin's 70t LEO + 20t TLI claims, for each engine.

  4. HEIGHT CHECK: does that propellant fit in a ~120 m, 7 m-diameter stack?
     Two height models are shown — the previous proportional model
     (stage height ∝ tank length) and a fixed-plus-tank model (engine bay,
     interstage, skirts are fixed length; only the barrel section grows).

  5. SENSITIVITY on the Northstar parameters that are not yet published.

Usage:
    python3 analyze_9x4_northstar.py

Output:
    NG_9x4_Northstar_Analysis.xlsx   (new file — does not touch existing workbooks)
"""

import math
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from scipy.optimize import brentq, least_squares

from generate_lv_design import lookup_engine, blended_isp, G0, C, style_hdr, style_sec, apply_border_range
import analyze_9x4_parameters as P   # geometry constants, height model, orbital targets

OUTFILE = "NG_9x4_Northstar_Analysis.xlsx"
LBF_TO_KN = 4.4482216e-3

# ── Claims ────────────────────────────────────────────────────────────────────
CLAIM_7X2_LEO = 45_000
CLAIM_7X2_GTO = 13_000
CLAIM_LEO     = 70_000
CLAIM_GTO     = 14_000
CLAIM_TLI     = 20_000

V_LEO = P.V_LEO          # 7,784 m/s
V_TLI = P.V_TLI          # ~10,838 m/s @ 300 km perigee

# ── Engines (publicly quoted thrusts, interpreted as sea level) ──────────────
# NOTE: the Engine DB stores BE-4's 2,400 kN / BE-4 Block 2's 2,847 kN in the
# *vacuum* column with a lower SL figure.  Blue Origin quotes these as sea-level
# thrust (550,000 / 640,000 lbf).  For a like-for-like comparison with Northstar
# (715,000 lbf) we use the quoted figures as SL thrust here.
ENGINES_S1 = {
    "BE-4":         {"sl_kn": 550_000 * LBF_TO_KN, "mass": 4760},
    "BE-4 Block 2": {"sl_kn": 640_000 * LBF_TO_KN, "mass": 4760},
    "Northstar":    {"sl_kn": lookup_engine("Northstar")["sl_thrust_kn"],
                     "mass":  lookup_engine("Northstar")["mass_kg"]},
}
_be4  = lookup_engine("BE-4")
ISP1  = blended_isp(_be4, 0.6)            # 331.4 s — Northstar Isp assumed = BE-4
VAC1  = _be4["vac_isp"]                   # 339 s, for recovery-prop calc
ISP2  = blended_isp(lookup_engine("BE-3U Block 2"), 1.0)   # 440 s

GRAV_COEFF = 1750   # same TWR sensitivity as the main design tool
SYS_FACTOR = 1.15   # engine + feed system mass factor (same as previous analysis)
BE3U_SYS   = 2700 * SYS_FACTOR
S1_STRUCT_7X2 = 55_000 - 7 * 4760 * SYS_FACTOR
S2_STRUCT_7X2 = 10_000 - 2 * BE3U_SYS
ADAPTER = 200


# ── Vehicle model ─────────────────────────────────────────────────────────────

def make_vehicle(s1p, s2p, engine="Northstar", n1=9, n2=4, subcool=0.035,
                 rec_dv=500, eng_mass=None, sl_kn=None, isp1=ISP1, isp2=ISP2):
    e = ENGINES_S1[engine]
    eng_mass = e["mass"] if eng_mass is None else eng_mass
    sl_kn    = e["sl_kn"] if sl_kn is None else sl_kn
    s1d = S1_STRUCT_7X2 * s1p / 770_000 + n1 * eng_mass * SYS_FACTOR
    s2d = S2_STRUCT_7X2 * s2p /  75_000 + n2 * BE3U_SYS
    glow = s1p + s1d + s2p + s2d
    return {
        "s1p": s1p, "s2p": s2p, "s1d": s1d, "s2d": s2d, "glow": glow,
        "eff1": s1p * (1 + subcool), "rec1": s1d * (math.exp(rec_dv / (VAC1 * G0)) - 1),
        "twr": n1 * sl_kn * 1000 / (glow * G0), "isp1": isp1, "isp2": isp2,
        "engine": engine, "n1": n1, "sl_kn": sl_kn, "eng_mass": eng_mass,
    }


def stage_dv(v, payload):
    pl  = payload + ADAPTER
    s2g = v["s2p"] + v["s2d"]
    dv2 = v["isp2"] * G0 * math.log((s2g + pl) / (v["s2d"] + pl))
    dv1 = v["isp1"] * G0 * math.log((v["eff1"] + v["s1d"] + s2g + pl)
                                     / (v["s1d"] + v["rec1"] + s2g + pl))
    return dv1 + dv2


def losses(v, K):
    return GRAV_COEFF / v["twr"] + K


def max_payload(v, K, v_target):
    f = lambda p: stage_dv(v, p) - losses(v, K) - v_target
    if f(0) < 0:
        return 0.0
    return brentq(f, 0, 1_000_000)


def payloads(v, K, v_gto):
    return {"LEO": max_payload(v, K, V_LEO),
            "GTO": max_payload(v, K, v_gto),
            "TLI": max_payload(v, K, V_TLI)}


# ── Step 1: calibrate on the 7x2 ─────────────────────────────────────────────

V7X2 = make_vehicle(770_000, 75_000, engine="BE-4", n1=7, n2=2, subcool=0.0)
L7X2 = stage_dv(V7X2, CLAIM_7X2_LEO) - V_LEO          # total loss that gives 45t
K    = L7X2 - GRAV_COEFF / V7X2["twr"]                 # TWR-independent remainder
V_GTO_EFF = stage_dv(V7X2, CLAIM_7X2_GTO) - L7X2       # Blue's effective "GTO" energy
V_GTO_STD = P.V_GTO                                     # 10,400 m/s (prev. analysis)


# ── Height models ─────────────────────────────────────────────────────────────

def tank_lengths(s1p, s2p):
    return P.prop_to_length(s1p, P.RHO_CH4_BULK), P.prop_to_length(s2p, P.RHO_LH2_BULK)

_L1_7X2, _L2_7X2 = tank_lengths(770_000, 75_000)
S1_FIXED = P.S1_HT_7X2 - _L1_7X2     # engine bay, skirts, interstage, domes
S2_FIXED = P.S2_HT_7X2 - _L2_7X2

def height_proportional(s1p, s2p):
    return sum(P.stage_heights_from_props(s1p, s2p))

def height_fixed_plus_tank(s1p, s2p):
    L1, L2 = tank_lengths(s1p, s2p)
    return S1_FIXED + L1 + S2_FIXED + L2

BUDGET = P.STAGE_BUDGET_9X4


# ── Step 3: back-solve propellant for the claims ─────────────────────────────

def back_solve(engine, **kw):
    def res(x):
        v = make_vehicle(x[0] * 1e3, x[1] * 1e3, engine=engine, **kw)
        L = losses(v, K)
        return [(stage_dv(v, CLAIM_LEO) - L - V_LEO),
                (stage_dv(v, CLAIM_TLI) - L - V_TLI)]
    sol = least_squares(res, [1000, 130], bounds=([400, 50], [3000, 500]))
    v = make_vehicle(sol.x[0] * 1e3, sol.x[1] * 1e3, engine=engine, **kw)
    return v, res(sol.x)


def fill_height(engine, height_fn, ratio, **kw):
    """Scale S1/S2 propellant at a fixed S2:S1 ratio until the stack fills the height budget."""
    f = lambda s1p: height_fn(s1p, s1p * ratio) - BUDGET
    s1p = brentq(f, 100_000, 5_000_000)
    return make_vehicle(s1p, s1p * ratio, engine=engine, **kw)


# ── Workbook helpers ──────────────────────────────────────────────────────────

def _sheet(wb, title, widths):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def _note(ws, r, text, span, height=30):
    c = ws.cell(r, 1, value=text)
    c.font = Font(size=9, italic=True, color="444444")
    c.fill = PatternFill("solid", fgColor=C["ltgray"])
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    ws.row_dimensions[r].height = height
    return r + 1


def _table(ws, r, headers, rows, fmts, highlight=None):
    for ci, h in enumerate(headers, 1):
        style_hdr(ws, r, ci, h, dark=False)
    ws.row_dimensions[r].height = 30
    r += 1
    for row in rows:
        for ci, (val, fmt) in enumerate(zip(row, fmts), 1):
            cell = ws.cell(r, ci, value=val)
            cell.number_format = fmt
            cell.font = Font(size=10)
            cell.alignment = Alignment(horizontal="left" if ci == 1 else "center", indent=1 if ci == 1 else 0)
            fill = C["ltgray"] if ci == 1 else C["calc"]
            if highlight:
                fill = highlight(ci, val, row) or fill
            cell.fill = PatternFill("solid", fgColor=fill)
        apply_border_range(ws, r, r, 1, len(headers))
        r += 1
    return r


def _claim_color(target):
    def f(ci, val, row):
        if not isinstance(val, (int, float)):
            return None
        return C["pos"] if val >= target * 0.98 else ("FFE0B2" if val >= target * 0.9 else C["neg"])
    return f


def _payload_highlight(col_targets):
    def f(ci, val, row):
        if ci in col_targets and isinstance(val, (int, float)):
            t = col_targets[ci]
            return C["pos"] if val >= t * 0.98 else ("FFE0B2" if val >= t * 0.9 else C["neg"])
        return None
    return f


def _vrow(label, v, pls, extra=()):
    return [label, v["s1p"] / 1e3, v["s2p"] / 1e3, v["s1d"] / 1e3, v["glow"] / 1e3,
            v["twr"], v.get("loss_override", losses(v, K)), pls["LEO"] / 1e3, pls["GTO"] / 1e3, pls["TLI"] / 1e3, *extra]


VHEAD = ["Case", "S1 prop (t)", "S2 prop (t)", "S1 dry (t)", "GLOW (t)", "Liftoff TWR",
         "Losses (m/s)", "LEO (t)", "GTO* (t)", "TLI (t)"]
VFMT  = ["@", "#,##0", "#,##0", "0.0", "#,##0", "0.00", "#,##0", "0.0", "0.0", "0.0"]
VHL   = _payload_highlight({8: CLAIM_LEO / 1e3, 9: CLAIM_GTO / 1e3, 10: CLAIM_TLI / 1e3})


# ── Sheets ────────────────────────────────────────────────────────────────────

def build_summary(wb, R):
    ws = _sheet(wb, "Summary", [46, 13, 13, 13, 13, 13, 13, 11, 11, 11])
    r = 1
    style_hdr(ws, r, 1, "NEW GLENN 9x4 WITH NORTHSTAR — PERFORMANCE RE-ESTIMATE", span=10)
    ws.row_dimensions[r].height = 24; r += 2
    r = _note(ws, r,
        "Blue Origin claims for New Glenn 9x4 (reusable booster): >70t LEO, >14t GTO, >20t TLI. "
        "Our previous BE-4 Block 2 estimate fell well short (~48t LEO / ~10t TLI). This workbook swaps in "
        "Northstar (715,000 lbf SL, 30% lighter than BE-4) and re-calibrates the loss model against the "
        "flown 7x2's 45t LEO claim. Green = within 2% of claim, orange = within 10%, red = short.", 10, 48)
    r += 1
    style_sec(ws, r, 1, "KEY RESULTS", span=10); r += 1
    rows = [
        _vrow("Previous estimate: BE-4 Block 2, 847t/111t (old loss model)", R["prev"]["v"], R["prev"]["pl"]),
        _vrow("BE-4 Block 2, 847t/111t (7x2-calibrated losses)", R["swap"]["BE-4 Block 2"]["v"], R["swap"]["BE-4 Block 2"]["pl"]),
        _vrow("Northstar, 847t/111t (engine swap only)", R["swap"]["Northstar"]["v"], R["swap"]["Northstar"]["pl"]),
        _vrow("Northstar, 7m core stretched to fill 120m (upper bound)", R["hmax"]["Northstar"]["v"], R["hmax"]["Northstar"]["pl"]),
        _vrow("Northstar, back-solved to 70t LEO + 20t TLI  ★", R["solve"]["Northstar"]["v"], R["solve"]["Northstar"]["pl"]),
        _vrow("BE-4 Block 2, back-solved to 70t LEO + 20t TLI", R["solve"]["BE-4 Block 2"]["v"], R["solve"]["BE-4 Block 2"]["pl"]),
    ]
    r = _table(ws, r, VHEAD, rows, VFMT, VHL)
    r += 1
    ns = R["solve"]["Northstar"]
    r = _note(ws, r,
        f"★ Best-fit Northstar 9x4: S1 ≈ {ns['v']['s1p']/1e3:,.0f}t, S2 ≈ {ns['v']['s2p']/1e3:,.0f}t propellant, "
        f"GLOW ≈ {ns['v']['glow']/1e3:,.0f}t. Stack height: {ns['h_fixed']:.0f} m with the fixed+tank model "
        f"(budget {BUDGET:.0f} m → fits) vs {ns['h_prop']:.0f} m with the old proportional model (does not fit). "
        f"The 70t LEO and 20t TLI claims cannot both be hit exactly in this model — the best compromise lands "
        f"~{ns['pl']['LEO']/1e3:.0f}t LEO / ~{ns['pl']['TLI']/1e3:.1f}t TLI (≈{abs(ns['res'][0]):.0f} m/s either way).", 10, 62)
    r = _note(ws, r,
        f"* GTO uses Blue Origin's effective GTO energy backed out of the 7x2 13t claim "
        f"(v_inject ≈ {V_GTO_EFF:,.0f} m/s, i.e. a GTO-1800-class orbit, not a minimum-energy GTO). "
        f"On that definition every 9x4 case beats the 14t claim by a wide margin — the 14t figure looks very conservative "
        f"(or is quoted to a different orbit) relative to the LEO and TLI figures.", 10, 48)
    r = _note(ws, r,
        "What closes the gap, in order of size: (1) calibrating losses on the flown 7x2 instead of assuming a 500 m/s "
        "recovery penalty; (2) a height model where engine bay / interstage don't grow with propellant, which lets the 9x4 "
        "carry ~25–30% more propellant in ~120 m; (3) Northstar itself — about +6t LEO at fixed tankage (≈ +15t lighter S1, "
        "+12% thrust → lower gravity loss) and ~120t less propellant needed to hit the claims.", 10, 48)


def build_assumptions(wb):
    ws = _sheet(wb, "Assumptions", [48, 14, 16, 60])
    r = 1
    style_hdr(ws, r, 1, "ASSUMPTIONS & CALIBRATION", span=4); ws.row_dimensions[r].height = 24; r += 2
    style_sec(ws, r, 1, "ENGINE DATA", span=4); r += 1
    ns = lookup_engine("Northstar")
    rows = [
        ["BE-4 SL thrust",               "kN", ENGINES_S1["BE-4"]["sl_kn"],         "550,000 lbf, Blue Origin quote (SL)"],
        ["BE-4 Block 2 SL thrust",       "kN", ENGINES_S1["BE-4 Block 2"]["sl_kn"], "640,000 lbf, Blue Origin quote (SL)"],
        ["Northstar SL thrust",          "kN", ns["sl_thrust_kn"],                  "~715,000 lbf announced; assumed to be the SL figure"],
        ["Northstar vac thrust",         "kN", ns["vac_thrust_kn"],                 "Assumed: SL thrust × Isp ratio 339/320"],
        ["BE-4 mass",                    "kg", 4760,                                "Engine DB"],
        ["Northstar mass",               "kg", ns["mass_kg"],                       "30% lighter than BE-4 (announced)"],
        ["S1 Isp (blended, 60% atm)",    "s",  ISP1,                                "320 + 0.6×(339−320). Northstar assumed = BE-4 (339 vac / 320 SL) — not published"],
        ["S2 Isp (BE-3U Block 2)",       "s",  ISP2,                                "Engine DB"],
        ["Engine system mass factor",    "—",  SYS_FACTOR,                          "Engine mass × 1.15 for feed lines, TVC, etc."],
    ]
    r = _table(ws, r, ["Parameter", "Unit", "Value", "Source / note"], rows, ["@", "@", "#,##0.0", "@"])
    r += 1
    style_sec(ws, r, 1, "7x2 CALIBRATION (BE-4, 770t / 75t, 55t / 10t dry, 500 m/s recovery)", span=4); r += 1
    rows = [
        ["7x2 liftoff TWR",                         "—",   V7X2["twr"],            "7 × 550,000 lbf / stage GLOW"],
        ["Total loss that reproduces 45t LEO",      "m/s", L7X2,                   "Solved"],
        [f"  gravity part ({GRAV_COEFF} / TWR)",    "m/s", GRAV_COEFF / V7X2["twr"], "Scales with TWR on the 9x4"],
        ["  remainder K (drag + trajectory)",        "m/s", K,                      "Held constant for the 9x4"],
        ["Effective GTO injection velocity",        "m/s", V_GTO_EFF,              "Velocity at which 7x2 delivers its claimed 13t"],
        ["Min-energy GTO (300 km perigee)",         "m/s", P.V_GTO,                "For comparison (previous analysis)"],
        ["LEO injection velocity",                  "m/s", V_LEO,                  "Previous analysis"],
        ["TLI injection velocity (300 km perigee)", "m/s", V_TLI,                  "Vis-viva"],
    ]
    r = _table(ws, r, ["Parameter", "Unit", "Value", "Note"], rows, ["@", "@", "#,##0.0", "@"])
    r += 1
    r = _note(ws, r,
        "Why calibrate this way: the previous 9x4 analysis assumed a 500 m/s trajectory penalty on top of a 1,400/TWR "
        f"gravity loss. Applied to the 7x2, that model under-predicts the 7x2's own 45t LEO claim. Fitting to the 7x2 gives a "
        f"remainder K ≈ {K:,.0f} m/s, i.e. essentially no extra recovery-profile penalty beyond drag. Only the TWR-driven "
        "gravity term changes between 7x2 and 9x4.", 4, 60)
    r = _note(ws, r,
        "Engine DB caveat: the existing BE-4 / BE-4 Block 2 rows store the quoted thrust in the Vac column and a lower value in "
        "the SL column. This analysis uses the quoted figures as SL thrust for all three engines so the TWR comparison is fair. "
        "The Northstar DB row uses 715,000 lbf as SL.", 4, 44)


def build_engine_swap(wb, R):
    ws = _sheet(wb, "Engine Swap", [40, 13, 13, 13, 13, 13, 13, 11, 11, 11])
    r = 1
    style_hdr(ws, r, 1, "ENGINE SWAP — same 9x4 tankage, different S1 engine", span=10)
    ws.row_dimensions[r].height = 24; r += 2
    r = _note(ws, r, "All cases: 9 S1 engines, 4 × BE-3U Block 2, 3.5% subcooling, 500 m/s recovery reserve, 7x2-calibrated losses.", 10)
    r += 1
    for (s1p, s2p) in R["swap_tanks"]:
        style_sec(ws, r, 1, f"S1 = {s1p/1e3:,.0f}t,  S2 = {s2p/1e3:,.0f}t", span=10); r += 1
        rows = []
        for eng in ("BE-4", "BE-4 Block 2", "Northstar"):
            v = make_vehicle(s1p, s2p, engine=eng)
            rows.append(_vrow(eng, v, payloads(v, K, V_GTO_EFF)))
        r = _table(ws, r, VHEAD, rows, VFMT, VHL) + 1


def build_back_solve(wb, R):
    ws = _sheet(wb, "Back-Solve & Height", [44, 13, 13, 13, 13, 13, 13, 11, 11, 11, 15, 15])
    r = 1
    style_hdr(ws, r, 1, "BACK-SOLVED PROPELLANT FOR 70t LEO + 20t TLI — and does it fit?", span=12)
    ws.row_dimensions[r].height = 24; r += 2
    r = _note(ws, r,
        f"Least-squares fit of S1 and S2 propellant to both claims. Height budget for the two stages = {BUDGET:.1f} m "
        f"(≈120 m vehicle minus fairing/adapters, 7 m core). Fixed sections calibrated from 7x2: S1 {S1_FIXED:.1f} m, "
        f"S2 {S2_FIXED:.1f} m.", 12, 36)
    r += 1
    head = VHEAD + ["Height, old model (m)", "Height, fixed+tank (m)"]
    fmt  = VFMT + ["0.0", "0.0"]

    def hl(ci, val, row):
        if ci in (11, 12) and isinstance(val, (int, float)):
            return C["pos"] if val <= BUDGET else C["neg"]
        return VHL(ci, val, row)

    rows = []
    for label, key in [("Northstar (3.5% subcool, 500 m/s rec)", "Northstar"),
                       ("Northstar (no subcool)", "Northstar_nosub"),
                       ("Northstar (300 m/s rec)", "Northstar_rec300"),
                       ("BE-4 Block 2 (3.5% subcool, 500 m/s rec)", "BE-4 Block 2")]:
        s = R["solve"][key]
        rows.append(_vrow(label, s["v"], s["pl"], (s["h_prop"], s["h_fixed"])))
    r = _table(ws, r, head, rows, fmt, hl) + 1

    style_sec(ws, r, 1, f"HEADROOM — tanks stretched to fill the {BUDGET:.0f} m budget at the Northstar best-fit S2:S1 ratio", span=12); r += 1
    rows = []
    for eng in ("BE-4 Block 2", "Northstar"):
        for lbl, key in [("fixed+tank height model", "hmax"), ("old proportional height model", "hmax_prop")]:
            h = R[key][eng]
            v = h["v"]
            rows.append(_vrow(f"{eng} — {lbl}", v, h["pl"],
                              (height_proportional(v["s1p"], v["s2p"]), height_fixed_plus_tank(v["s1p"], v["s2p"]))))
    r = _table(ws, r, head, rows, fmt, hl) + 1
    r = _note(ws, r,
        "The old proportional height model treats the whole stage as if it stretched with the tanks (S1 tanks are only ~45% "
        "of the 7x2 S1's 57.5 m), so every extra tonne of propellant cost >2× its tank length in height. That model is what "
        "made 70t look geometrically impossible. With fixed engine bay / interstage lengths the claims fit in ~120 m on a 7 m core. "
        "Headroom rows are upper bounds: the model has no upper-stage gravity loss and assumes dry mass scales linearly.", 12, 48)


def build_sensitivity(wb, R):
    ws = _sheet(wb, "Sensitivity", [34, 12, 12, 12, 12, 12, 12])
    r = 1
    style_hdr(ws, r, 1, "SENSITIVITY — unpublished Northstar parameters (best-fit tankage held fixed)", span=7)
    ws.row_dimensions[r].height = 24; r += 2
    base = R["solve"]["Northstar"]["v"]
    s1p, s2p = base["s1p"], base["s2p"]

    def block(title, label, values, kw_fn, fmt):
        nonlocal r
        style_sec(ws, r, 1, title, span=7); r += 1
        rows = []
        for val in values:
            v = make_vehicle(s1p, s2p, engine="Northstar", **kw_fn(val))
            pl = payloads(v, K, V_GTO_EFF)
            rows.append([fmt(val), v["twr"], losses(v, K), pl["LEO"] / 1e3, pl["GTO"] / 1e3, pl["TLI"] / 1e3])
        r = _table(ws, r, [label, "Liftoff TWR", "Losses (m/s)", "LEO (t)", "GTO* (t)", "TLI (t)"], rows,
                   ["@", "0.00", "#,##0", "0.0", "0.0", "0.0"],
                   _payload_highlight({4: CLAIM_LEO / 1e3, 5: CLAIM_GTO / 1e3, 6: CLAIM_TLI / 1e3})) + 1

    block("Mass reduction vs BE-4", "Northstar mass", [0.0, 0.1, 0.2, 0.3, 0.4],
          lambda x: {"eng_mass": 4760 * (1 - x)}, lambda x: f"−{x:.0%}  ({4760*(1-x):,.0f} kg)")
    block("Thrust interpretation", "715 klbf is…", ["SL", "Vac"],
          lambda x: {"sl_kn": 715_000 * LBF_TO_KN * (1 if x == "SL" else 320 / 339)},
          lambda x: f"{x} thrust")
    block("S1 blended Isp", "S1 Isp", [322.0, 327.0, ISP1, 336.0],
          lambda x: {"isp1": x}, lambda x: f"{x:.1f} s")
    block("S1 recovery ΔV reserve", "Recovery ΔV", [0, 300, 500, 700],
          lambda x: {"rec_dv": x}, lambda x: f"{x} m/s")
    _note(ws, r, "Each row changes one parameter from the best-fit Northstar case; everything else held constant.", 7)


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    R = {}
    # Previous estimate, re-evaluated with the old loss model (1400/TWR + 500 + 120, 300 m/s rec)
    vp = make_vehicle(847_000, 111_000, engine="BE-4 Block 2", rec_dv=300,
                      sl_kn=lookup_engine("BE-4 Block 2")["sl_thrust_kn"])
    old_loss = lambda v: 1400 / v["twr"] + 500 + 120
    def old_pl(v, tgt):
        f = lambda p: stage_dv(v, p) - old_loss(v) - tgt
        return brentq(f, 0, 1e6) if f(0) > 0 else 0.0
    R["prev"] = {"v": vp, "pl": {"LEO": old_pl(vp, V_LEO), "GTO": old_pl(vp, V_GTO_EFF), "TLI": old_pl(vp, V_TLI)}}
    R["prev"]["v"] = {**vp, "loss_override": old_loss(vp)}   # show old-model loss in tables

    R["swap_tanks"] = [(847_000, 111_000), (915_000, 138_000)]
    R["swap"] = {}
    for eng in ("BE-4 Block 2", "Northstar"):
        v = make_vehicle(847_000, 111_000, engine=eng)
        R["swap"][eng] = {"v": v, "pl": payloads(v, K, V_GTO_EFF)}

    R["solve"] = {}
    for key, eng, kw in [("Northstar", "Northstar", {}),
                         ("Northstar_nosub", "Northstar", {"subcool": 0.0}),
                         ("Northstar_rec300", "Northstar", {"rec_dv": 300}),
                         ("BE-4 Block 2", "BE-4 Block 2", {})]:
        v, res = back_solve(eng, **kw)
        R["solve"][key] = {"v": v, "res": res, "pl": payloads(v, K, V_GTO_EFF),
                           "h_prop": height_proportional(v["s1p"], v["s2p"]),
                           "h_fixed": height_fixed_plus_tank(v["s1p"], v["s2p"])}
    ns = R["solve"]["Northstar"]["v"]
    R["swap_tanks"][1] = (round(ns["s1p"], -3), round(ns["s2p"], -3))

    ratio = ns["s2p"] / ns["s1p"]
    R["hmax"], R["hmax_prop"] = {}, {}
    for eng in ("BE-4 Block 2", "Northstar"):
        for key, fn in [("hmax", height_fixed_plus_tank), ("hmax_prop", height_proportional)]:
            v = fill_height(eng, fn, ratio)
            R[key][eng] = {"v": v, "pl": payloads(v, K, V_GTO_EFF)}
    return R


def main():
    print("NG 9x4 + Northstar re-estimate")
    print("=" * 64)
    print(f"  7x2 calibration: TWR {V7X2['twr']:.2f}, loss for 45t LEO = {L7X2:,.0f} m/s "
          f"(gravity {GRAV_COEFF/V7X2['twr']:,.0f} + K {K:,.0f})")
    print(f"  Effective GTO v_inject (7x2 13t): {V_GTO_EFF:,.0f} m/s")
    R = run()
    def show(lbl, d):
        v, pl = d["v"], d["pl"]
        print(f"  {lbl:44s} S1 {v['s1p']/1e3:5.0f}t S2 {v['s2p']/1e3:4.0f}t TWR {v['twr']:.2f}  "
              f"LEO {pl['LEO']/1e3:5.1f}  GTO {pl['GTO']/1e3:5.1f}  TLI {pl['TLI']/1e3:5.1f}")
    print()
    show("Previous (BE-4 B2, old loss model)", R["prev"])
    show("BE-4 B2 847/111 (calibrated)", R["swap"]["BE-4 Block 2"])
    show("Northstar 847/111 (calibrated)", R["swap"]["Northstar"])
    show("Northstar fill height (fixed+tank)", R["hmax"]["Northstar"])
    show("Northstar fill height (old model)", R["hmax_prop"]["Northstar"])
    for k in ("Northstar", "Northstar_nosub", "Northstar_rec300", "BE-4 Block 2"):
        s = R["solve"][k]
        show(f"Back-solve {k}", s)
        print(f"  {'':44s} height: old {s['h_prop']:.0f} m, fixed+tank {s['h_fixed']:.0f} m (budget {BUDGET:.0f}); "
              f"dry S1 {s['v']['s1d']/1e3:.1f}t S2 {s['v']['s2d']/1e3:.1f}t; GLOW {s['v']['glow']/1e3:,.0f}t")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_summary(wb, R)
    build_assumptions(wb)
    build_engine_swap(wb, R)
    build_back_solve(wb, R)
    build_sensitivity(wb, R)
    for ws, color in zip(wb.worksheets, ["1F3864", "BF360C", "2E7D32", "4527A0", "00695C"]):
        ws.sheet_properties.tabColor = color

    if os.path.exists(OUTFILE):
        raise SystemExit(f"{OUTFILE} already exists — refusing to overwrite. Rename or remove it first.")
    wb.save(OUTFILE)
    print(f"\n✓ Saved: {OUTFILE}")
    return R


if __name__ == "__main__":
    main()
