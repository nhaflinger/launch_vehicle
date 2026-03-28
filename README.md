# Launch Vehicle Design Tool

A Python script that generates an Excel workbook for conceptual launch vehicle design and performance analysis, grounded in real-world engine and vehicle data.

## Usage

```bash
python3 generate_lv_design.py
```

Opens `launch_vehicle_design.xlsx`. Requires `openpyxl`:

```bash
pip3 install openpyxl
```

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **Vehicle Design** | Main working sheet — fill in yellow cells, blue cells calculate automatically |
| **Engine DB** | 39 real engines with performance data (Isp, thrust, mass, chamber pressure) |
| **Vehicle DB** | 13 reference vehicles with per-stage mass breakdowns |
| **Propellants** | Propellant properties and subsystem mass fraction model |
| **Comparison** | Quick-compare table across all reference vehicles |
| **README** | In-workbook user guide |

## How to Use

1. Open **Vehicle Design** and fill in the yellow **Mission Requirements** cells
2. For each stage, enter the engine name (must match Engine DB exactly), engine count, propellant combination, and propellant mass
3. The **Mission ΔV Analysis** section computes gravity loss from Stage 1 TWR — higher thrust automatically reduces gravity loss and increases payload
4. Subsystem masses are estimated automatically from parametric models; override any cell with a known value
5. **MAX PAYLOAD TO LEO** and **MAX PAYLOAD TO GTO** calculate live in the Performance section

## Physics Model

**Delta-V** per stage via the Tsiolkovsky rocket equation:

```
ΔV = Isp × g₀ × ln(m₀ / m_f)
```

**Max payload** solved analytically from the last stage:

```
m_payload = prop / (exp(ΔV_alloc / (Isp × g₀)) − 1) − dry_mass − adapter_mass
```

**Gravity loss** computed from liftoff TWR:

```
Gravity Loss = Coefficient / TWR_liftoff   (default coefficient = 1750)
```

**Mission ΔV** = Orbital Velocity + Drag Loss + Gravity Loss (all calculated, not hardcoded)

**Effective Isp** accounts for atmospheric pressure during Stage 1 burn:

```
Effective Isp = SL_Isp + Atm_Fraction × (Vac_Isp − SL_Isp)
```

Stage 1 default atmospheric fraction = 0.6 (60% of burn at altitude). Upper stages default to 1.0 (pure vacuum).

**Subcooling** (propellant densification) modeled as a percentage gain on nominal propellant mass, applied to performance calculations only — subsystem mass fractions use nominal tank capacity.

## Reference Data

### Engines (selected)
| Engine | Vehicle | Propellant | Vac Isp (s) | Vac Thrust (kN) |
|---|---|---|---|---|
| Merlin 1D | Falcon 9 S1 | LOX/RP-1 | 311 | 934 |
| RS-25 | SLS Core | LOX/LH2 | 452 | 2090 |
| BE-4 | New Glenn S1 | LOX/CH4 | 339 | 2400 |
| BE-4 Block 2 | New Glenn Block 2 | LOX/CH4 | 339 | 2847 |
| BE-3U | New Glenn S2 | LOX/LH2 | 440 | 710 |
| BE-3U Block 2 | New Glenn Block 2 S2 | LOX/LH2 | 440 | 890 |
| Raptor 2 | Starship | LOX/CH4 | 350 | 2350 |
| F-1 | Saturn V S-IC | LOX/RP-1 | 304 | 7740 |

### Reference Vehicles (selected)
| Vehicle | Payload LEO (kg) | Payload GTO (kg) | GLOW (kg) |
|---|---|---|---|
| Falcon 9 Block 5 | 22,800 | 8,300 | 549,054 |
| New Glenn 7x2 Baseline | 45,000 | 13,000 | 1,000,000 |
| New Glenn 9x4 | 70,000+ | 14,000+ | ~1,300,000 |
| SLS Block 1 | 95,000 | — | 2,608,000 |
| Saturn V | 130,000 | — | 2,970,000 |

## Limitations

- Simplified mass models — early concept accuracy is ±20–30%
- Gravity loss model is empirical, not trajectory-integrated
- Parallel staging (Soyuz, Ariane 5 strap-ons firing with core) not modeled — treat as series stages
- No trajectory shaping, throttling, or staging timing
- New Glenn 9x4 propellant masses are estimates (not published by Blue Origin)
