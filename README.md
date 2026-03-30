# Launch Vehicle Design Tool

A Python script that generates an Excel workbook for conceptual launch vehicle design and performance analysis, grounded in real-world engine and vehicle data.

## Usage

```bash
python3 generate_lv_design.py
```

Outputs `lv_design_<vehicle_name>.xlsx` (filename derived from the Vehicle Name field). Requires `openpyxl`:

```bash
pip3 install openpyxl
```

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Vehicle Design** | Main working sheet — fill in yellow cells, blue cells calculate automatically |
| **Engine DB** | 40 real engines with performance data (Isp, thrust, mass, chamber pressure) |
| **Vehicle DB** | 13 reference vehicles with per-stage mass breakdowns |
| **Propellants** | Propellant properties and subsystem mass fraction model |
| **Comparison** | Quick-compare table across all reference vehicles |
| **README** | In-workbook user guide |

---

## How to Use

1. Go to the **Vehicle Design** tab — this is your main working sheet.
2. Fill in **Mission Requirements** (yellow cells): vehicle name, target orbit, payload adapter mass, orbital velocity, drag loss, gravity loss coefficient, trajectory penalty, GTO orbital ΔV, and number of stages.
3. For each stage, configure:
   - **Engine**: type the exact engine name from the Engine DB tab — Isp will auto-populate via VLOOKUP.
   - **Number of Engines**: how many of that engine on this stage.
   - **Propellant Combination**: must match exactly (e.g. `LOX/RP-1`, `LOX/LH2`, `LOX/CH4`, `Solid HTPB`, `NTO/MMH`).
   - **Propellant Mass**: total usable propellant mass for this stage in kg.
   - **Vac Isp Override**: leave blank to use the Engine DB value; enter a number to override.
4. **Subsystem masses** auto-calculate from parametric models in the Propellants tab. Yellow cells can be manually overridden. Use **Dry Mass Override** if you have a known dry mass from a reference vehicle.
5. For **reusable stages** (e.g. Falcon 9 S1), enter a **Recovery ΔV Reserve**:
   - This is the total ΔV budget for recovery burns (boost-back + entry + landing).
   - Typical values: ~750 m/s for RTLS (return-to-launch-site), ~500 m/s for ASDS (drone ship).
   - The tool computes Recovery Propellant Reserved = `dry_mass × (e^(ΔV/Isp/g₀) − 1)`.
   - This propellant is treated as dead weight during ascent (still onboard at MECO).
   - Set to 0 for expendable stages.
6. The **Performance** section calculates automatically:
   - Delta-V per stage via the Tsiolkovsky equation.
   - Burnout mass includes recovery-reserved propellant for reusable stages.
   - TWR uses sea-level thrust for Stage 1, vacuum thrust for upper stages.

---

## Physics Model

**Delta-V** per stage (Tsiolkovsky rocket equation):
```
ΔV = Isp × g₀ × ln(m₀ / m_f)
```

**Max payload** solved analytically from the last stage:
```
m_payload = prop / (exp(ΔV_alloc / (Isp × g₀)) − 1) − dry_mass − adapter_mass
```

**Gravity loss** computed from liftoff TWR plus an additive trajectory penalty:
```
Gravity Loss = (Coefficient / TWR_liftoff) + Trajectory Penalty
```

**Mission ΔV**:
```
LEO Mission ΔV = Orbital Velocity + Drag Loss + Gravity Loss
GTO Mission ΔV = GTO Orbital ΔV + Drag Loss + Gravity Loss
```

**Effective Isp** accounts for atmospheric pressure during Stage 1 burn:
```
Effective Isp = SL_Isp + Atm_Fraction × (Vac_Isp − SL_Isp)
```
Stage 1 default atmospheric fraction = 0.6. Upper stages default to 1.0 (pure vacuum).

**Subcooling** (propellant densification) modeled as a percentage gain on nominal propellant mass, applied to performance calculations only — subsystem mass fractions use nominal tank capacity.

---

## Tips & Rules of Thumb

- LEO missions typically need 9,000–9,500 m/s total ΔV (includes ~1,500 m/s gravity + drag losses).
- GTO missions need ~11,500–12,500 m/s depending on inclination.
- First stage TWR should be 1.2–1.5 at liftoff; too low = slow ascent; too high = unnecessary mass.
- Payload fraction for real vehicles: 1.5–4% to LEO is typical.
- Structural mass fraction (dry/propellant) is typically 5–15%; lower is better engineering.
- LOX/LH2 gives the highest Isp (~450 s) but low density means large, heavy tanks.
- LOX/RP-1 gives good density and SL performance; workhorse of most first stages.
- LOX/CH4 is the emerging choice for reusability; good Isp + density balance.
- Reusable first stages carry a payload penalty: ~30–40% less to LEO vs expendable.
- Recovery ΔV budget: RTLS ~750 m/s (boost-back ~300 + entry ~200 + landing ~250); ASDS ~450–500 m/s (no boost-back; entry ~200 + landing ~250).

---

## Reference Data

### Engines (selected)

| Engine | Vehicle | Propellant | Vac Isp (s) | Vac Thrust (kN) |
|---|---|---|---|---|
| Merlin 1D | Falcon 9 S1 | LOX/RP-1 | 311 | 934 |
| Merlin 1D Vac | Falcon 9 S2 | LOX/RP-1 | 348 | 934 |
| RS-25 | SLS Core | LOX/LH2 | 452 | 2090 |
| BE-4 | New Glenn S1 | LOX/CH4 | 339 | 2400 |
| BE-4 Block 2 | New Glenn Block 2 S1 | LOX/CH4 | 339 | 2847 |
| BE-3U | New Glenn S2 | LOX/LH2 | 440 | 710 |
| BE-3U Block 2 | New Glenn Block 2 S2 | LOX/LH2 | 440 | 890 |
| Raptor 2 | Starship / Super Heavy | LOX/CH4 | 350 | 2350 |
| F-1 | Saturn V S-IC | LOX/RP-1 | 304 | 7740 |

### Reference Vehicles (selected)

| Vehicle | Payload LEO (kg) | Payload GTO (kg) | GLOW (kg) |
|---|---|---|---|
| Falcon 9 Block 5 | 22,800 | 8,300 | 549,054 |
| Falcon Heavy | 63,800 | 26,700 | 1,420,788 |
| New Glenn 7x2 Baseline | 45,000 | 13,000 | 1,000,000 |
| New Glenn 9x4 | 70,000+ | 14,000+ | ~1,300,000 |
| SLS Block 1 | 95,000 | — | 2,608,000 |
| Saturn V | 130,000 | — | 2,970,000 |
| Starship | 150,000 | — | 5,000,000 |

---

## Limitations

- Simplified mass models — early concept accuracy is ±20–30%.
- Gravity loss model is empirical, not trajectory-integrated.
- Parallel staging (Soyuz, Ariane 5 strap-ons firing simultaneously with core) is not modeled — treat as series stages for ΔV calculation.
- No trajectory shaping, throttling timeline, or staging timing.
- New Glenn 9x4 propellant masses are estimates (not published by Blue Origin).

---

## Simplifying Assumptions & Calibration Parameters

This model combines physics-based equations with several parameters that cannot be derived from first principles alone. Some defaults are reasonable engineering approximations. Others were chosen specifically to bring outputs into agreement with Falcon 9 Block 5 published performance figures and should be treated as guesses until calibrated against a vehicle you trust.

### Tunable Parameters

| Parameter | Default | Confidence | Notes |
|---|---|---|---|
| Gravity Loss Coefficient | 1,750 | **LOW** | Empirical. Fitted to approximate typical vehicles. Needs recalibration per vehicle family. |
| Trajectory Penalty | 500 m/s | **GUESS** | Reverse-engineered from F9 ASDS LEO payload figure. No independent physical basis. Set to 0 for expendable vehicles. Do not apply to other vehicles without recalibrating against known performance. |
| Stage ΔV Fractions (LEO) | 0.40 / 0.60 | **GUESS** | Reverse-engineered to hit a target F9 payload number. Real optimal splits depend on mass ratios and Isp. Re-derive per vehicle by iterating until the ΔV allocation check = 0. |
| Stage ΔV Fractions (GTO) | 0.33 / 0.67 | **GUESS** | Same issue as LEO fractions. Calibrate independently against a known GTO figure. |
| Recovery ΔV Reserve | 500 m/s | **MODERATE** | Plausible range for ASDS. Varies with S1 mass and landing site distance. Not verified against flight telemetry. |
| GTO Orbital ΔV | 10,400 m/s | **MODERATE** | Varies ±300–600 m/s with launch site, inclination, and target apogee. Set per mission rather than using the default. |
| Atm Fraction (S1) | 0.6 | **MODERATE** | Common rule of thumb. Varies with ascent profile and nozzle design. Can be off by 5–10 s effective Isp. |
| Drag Loss | 120 m/s | **MODERATE** | Reasonable for a typical trajectory. Varies with vehicle diameter and max-q throttling strategy. |
| Subsystem mass fractions (Propellants tab) | various | **MODERATE** | Statistical fits to a small number of real vehicles. ±20–30% accuracy claimed; individual subsystems may be worse. |

### Structural Simplifications (not tunable)

These are inherent to the model architecture, not user inputs:

- **Effective Isp** is modeled as a fixed linear blend between SL and vacuum — not an altitude-integrated curve.
- **Recovery burns** are modeled as a single equivalent burn from dry mass, not multiple discrete burns (entry + landing). Single-burn model slightly underestimates total recovery propellant required.
- **Fairing jettison** is not modeled as a staging event — add fairing mass to Payload Adapter as a workaround, accepting it is treated as dead weight to orbit.
- **Isp does not vary with throttle level** — real engines shift Isp by ~1–3% when throttled.
- **GLOW for TWR/gravity loss calculation excludes payload mass** — makes liftoff TWR slightly optimistic.
- **Parallel staging** (strap-on boosters burning simultaneously with core) is not supported.
