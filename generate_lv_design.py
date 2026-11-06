#!/usr/bin/env python3
"""
Launch Vehicle Design Spreadsheet Generator
Usage: python3 generate_lv_design.py
Output: launch_vehicle_design.xlsx
"""

import math
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── Constants ─────────────────────────────────────────────────────────────────
OUTPUT_FILE = "launch_vehicle_design.xlsx"
G0 = 9.80665
MAX_STAGES = 3

# ── Colors ────────────────────────────────────────────────────────────────────
C = {
    "hdr":    "1F3864",   # dark blue header
    "sub":    "2F5496",   # medium blue subheader
    "input":  "FFF2CC",   # yellow: user editable
    "calc":   "DAEEF3",   # light blue: formula/calculated
    "sec":    "D6DCE4",   # gray: section label
    "stage":  ["E2EFDA", "FCE4D6", "EAD1DC"],  # per-stage tints
    "pos":    "C6EFCE",
    "neg":    "FFC7CE",
    "white":  "FFFFFF",
    "ltgray": "F2F2F2",
}

# ── Engine Database ───────────────────────────────────────────────────────────
# name, vehicle_application, propellant, isp_vac_s, isp_sl_s,
# thrust_vac_kn, thrust_sl_kn, mass_kg, chamber_press_bar, exp_ratio, notes
ENGINES = [
    # LOX/RP-1
    ("Merlin 1D",       "Falcon 9 S1",          "LOX/RP-1",  311, 282,  934,  845,  470, 97,  16,   "9 engines on F9 first stage; sea-level optimized"),
    ("Merlin 1D Vac",   "Falcon 9 S2",          "LOX/RP-1",  348, None, 934,  None, 490, 97,  165,  "Vacuum-optimized fixed nozzle extension"),
    ("RD-180",          "Atlas V CBC",           "LOX/RP-1",  338, 311, 4152, 3828, 5480,257, 36.4, "Dual-chamber single-turbopump; Russian-designed"),
    ("RD-107A",         "Soyuz Strap-on",        "LOX/RP-1",  320, 263, 1020,  838, 1190, 60, 21.0, "4 strap-on boosters; 4 fixed + 4 vernier chambers"),
    ("RD-108A",         "Soyuz Core",            "LOX/RP-1",  320, 257,  990,  792, 1278, 60, 21.0, "Core stage; 4 main + 4 vernier chambers"),
    ("RD-0124",         "Soyuz S3 / Angara S2",  "LOX/RP-1",  359, 311,  294,  None, 480,157, 30.0, "High-performance staged combustion upper stage"),
    ("F-1",             "Saturn V S-IC",         "LOX/RP-1",  304, 263, 7740, 6770, 8391, 70, 16.0, "5 engines on Saturn V first stage; largest US engine"),
    ("RD-191",          "Angara S1",             "LOX/RP-1",  337, 311, 2090, 1922, 2200,263, 37.0, "Single-chamber RD-180 derivative"),
    ("NK-33 (AJ-26)",   "Antares S1",            "LOX/RP-1",  331, 297, 1638, 1510, 1222,146, 27.0, "Soviet N1 engine; modernized as AJ-26"),
    ("Rutherford",      "Electron S1",           "LOX/RP-1",  311, 303,   24,   22,   35,145, 11.0, "9 engines; electric-pump cycle; additive manufactured"),
    ("Rutherford Vac",  "Electron S2",           "LOX/RP-1",  343, None,  25, None,   35,145, 250,  "Vacuum variant; electric-pump cycle"),
    # LOX/LH2
    ("RS-25 (SSME)",    "SLS Core / Shuttle",    "LOX/LH2",   452, 366, 2090, 1860, 3177,206, 69.0, "3x on SLS; 4x on Shuttle; high-performance SSME"),
    ("RL-10A-4-2",      "Atlas V Centaur",       "LOX/LH2",   451, None,  99, None,  168, 28, 84.0, "Expander cycle; multiple restarts"),
    ("RL-10B-2",        "Delta IV DCSS",         "LOX/LH2",   462, None, 110, None,  277, 44,285,   "Extendable carbon-carbon nozzle"),
    ("RL-10C-1",        "Vulcan Centaur",        "LOX/LH2",   450, None, 102, None,  190, 34, 84.0, "Updated RL-10; simplified nozzle"),
    ("RL-10C-3",        "SLS ICPS",              "LOX/LH2",   460, None, 110, None,  190, 34,150,   "High-expansion ratio variant"),
    ("Vulcain 2",       "Ariane 5 EPC",          "LOX/LH2",   431, 318, 1340,  960, 1686,115, 60.0, "Gas-generator cycle; Ariane 5 core"),
    ("Vinci",           "Ariane 6 ULPM",         "LOX/LH2",   465, None, 180, None,  550, 61,240,   "Expander cycle; restartable upper stage"),
    ("HM7B",            "Ariane 5 ESC-A",        "LOX/LH2",   446, None,  65, None,  165, 35, 83.0, "Gas-generator; multiple restarts"),
    ("J-2",             "Saturn V S-II / S-IVB", "LOX/LH2",   421, None,1033, None, 1788, 52, 27.5, "5x on S-II; 1x on S-IVB; restartable"),
    ("RS-68A",          "Delta IV Heavy CBC",    "LOX/LH2",   412, 365, 3370, 2949, 6597, 97, 21.5, "Gas-generator; simple design; largest H2 engine"),
    ("BE-3U",           "New Glenn S2",          "LOX/LH2",   440, None, 710, None, 2700, 90,200,   "Blue Origin upper stage engine; expander cycle; 2x on NG 7x2"),
    ("BE-3U Block 2",   "New Glenn Block 2 S2",  "LOX/LH2",   440, None, 890, None, 2700, 90,200,   "Upgraded BE-3U; 890 kN vac (200,000 lbf); original Block 2 announced spec; 2x on NG 7x2; 4x on NG 9x4"),
    ("BE-3U Block 2+",  "New Glenn Block 2 S2",  "LOX/LH2",   440, None, 979, None, 2700, 90,200,   "Further upgraded BE-3U; 979 kN vac (220,000 lbf) announced; up from 890 kN Block 2; Isp held at 440 s pending confirmation"),
    # LOX/CH4
    ("Raptor 2",        "Starship / Super Heavy","LOX/CH4",   350, 327, 2350, 2200, 1500,330, 32.0, "Full-flow staged combustion; 33x on Super Heavy"),
    ("Raptor Vac",      "Starship upper",        "LOX/CH4",   380, None,2200, None, 1600,330,107,   "Vacuum variant; large fixed nozzle; 3x on Starship"),
    ("BE-4",            "Vulcan / New Glenn S1", "LOX/CH4",   339, 320, 2400, 2100, 4760,134, 40.0, "Oxygen-rich staged combustion; 2x Vulcan / 7x NG 7x2 baseline"),
    ("BE-4 Block 2",    "New Glenn Block 2 S1",  "LOX/CH4",   339, 320, 2847, 2490, 4760,134, 40.0, "Upgraded BE-4 with subcooling; 2,847 kN SL (640,000 lbf); use with Subcooling Gain ~3.5%"),
    ("BE-4 Block 2 (no subcool)", "New Glenn Block 2 S1", "LOX/CH4", 339, 320, 2780, 2430, 4760,134, 40.0, "Upgraded BE-4 design improvements only, no subcooling; 2,780 kN SL (625,000 lbf); use with Subcooling Gain 0%"),
    ("Northstar",       "New Glenn 9x4 S1",      "LOX/CH4",   339, 320, 3369, 3181, 3332,134, 40.0, "BE-4 derivative announced Nov 2026; ~715,000 lbf (3,181 kN) SL; 30% lighter than BE-4 (0.70 x 4,760 kg); Isp, Pc, expansion ratio assumed = BE-4 (not published); vac thrust scaled by Isp ratio 339/320"),
    # Hypergolic
    ("AJ-10-190",       "Orion SPS / STS OMS",   "NTO/MMH",   316, None,  27, None,  118,  9, 49.0, "Pressure-fed; restartable; long heritage"),
    ("Vikas",           "PSLV / GSLV S2",        "NTO/UDMH",  293, 270,  800,  725,  634, 58, 16.0, "Indian engine; Ariane Viking derivative"),
    ("YF-20 (cluster)", "Long March S1",         "NTO/UDMH",  289, 259, 2962, 2799, 4000, 65, 14.0, "4-chamber cluster used on CZ-2/3/4 first stages"),
    ("Aestus",          "Ariane 5 ATV",          "NTO/MMH",   324, None,  29, None,  111, 11, 84.0, "Pressure-fed; restartable; ESA upper stage"),
    # Solid
    ("SRB (SLS 5-seg)", "SLS Block 1",           "Solid HTPB",269, 242,16000,14900,90718, 62,  7.5, "Largest SRB ever flown; 2x on SLS"),
    ("RSRM (Shuttle)",  "Space Shuttle",         "Solid HTPB",268, 242,14678,12453,90714, 62,  7.7, "4-segment Shuttle SRB"),
    ("GEM-63XL",        "Vulcan strap-on",       "Solid HTPB",279, 252, 2140, 1930, 7000, 60, 11.0, "Solid strap-on; also used on Atlas V 500"),
    ("P80",             "Ariane 6 S1",           "Solid HTPB",278, 255, 3040, 2820, 5650, 95, 16.0, "Ariane 6 first stage solid motor"),
    ("Zefiro-23",       "Vega S2",               "Solid HTPB",287, None,1120, None, 1315, 87, 30.0, "Vega second stage solid motor"),
    ("Zefiro-9",        "Vega S3",               "Solid HTPB",296, None, 314, None,  260, 98, 54.0, "Vega third stage solid motor"),
    ("Star-48BV",       "Various kick stages",   "Solid HTPB",292, None,  68, None,  119, 40, 55.0, "Small solid kick motor; spin-stabilized"),
    ("Castor 30B",      "Antares S2",            "Solid HTPB",302, None, 357, None, 1437, 80, 57.0, "Antares second stage solid motor"),
]

# ── Reference Vehicle Database ─────────────────────────────────────────────────
# Each vehicle has metadata and a list of stages:
#   (stage_name, engine_name, n_engines, propellant_mass_kg, dry_mass_kg, notes)
VEHICLES = [
    {
        "name": "Falcon 9 Block 5", "country": "USA", "operator": "SpaceX",
        "first_flight": 2018, "status": "Active",
        "payload_leo_kg": 22800, "payload_gto_kg": 8300, "payload_tli_kg": None,
        "glow_kg": 549054, "height_m": 70.0, "diameter_m": 3.66,
        "notes": "22,800 kg LEO expendable / 17,400 kg reusable; partially reusable first stage",
        "stages": [
            ("S1 Core",     "Merlin 1D",      9, 411000, 22200, "Reusable; includes grid fins, landing legs, interstage"),
            ("S2 Upper",    "Merlin 1D Vac",  1, 111500,  4000, "Single Merlin Vacuum engine"),
        ],
    },
    {
        "name": "Falcon Heavy", "country": "USA", "operator": "SpaceX",
        "first_flight": 2018, "status": "Active",
        "payload_leo_kg": 63800, "payload_gto_kg": 26700, "payload_tli_kg": None,
        "glow_kg": 1420788, "height_m": 70.0, "diameter_m": 12.2,
        "notes": "Expendable config; 3 F9 cores; center core reinforced",
        "stages": [
            ("Center Core", "Merlin 1D",      9, 411000, 25600, "Reinforced; cross-feed ready"),
            ("Side Booster (x2)", "Merlin 1D",9, 411000, 22200, "Two identical; mass shown per booster"),
            ("S2 Upper",    "Merlin 1D Vac",  1, 111500,  4000, "Single Merlin Vacuum"),
        ],
    },
    {
        "name": "Atlas V 401", "country": "USA", "operator": "ULA",
        "first_flight": 2002, "status": "Active (limited)",
        "payload_leo_kg": 9750, "payload_gto_kg": 4750, "payload_tli_kg": None,
        "glow_kg": 334500, "height_m": 58.3, "diameter_m": 3.81,
        "notes": "401 = 4m fairing, 0 SRBs, 1 RL-10; 551 config adds 5 GEM SRBs and dual RL-10",
        "stages": [
            ("CBC Booster",  "RD-180",        1, 284089, 21054, "Russian-built engine; kerosene/LOX"),
            ("Centaur III",  "RL-10A-4-2",    1,  20830,  2316, "High-energy upper stage; LOX/LH2"),
        ],
    },
    {
        "name": "Delta IV Heavy", "country": "USA", "operator": "ULA",
        "first_flight": 2004, "status": "Retired 2024",
        "payload_leo_kg": 28370, "payload_gto_kg": 14220, "payload_tli_kg": None,
        "glow_kg": 733400, "height_m": 72.0, "diameter_m": 15.0,
        "notes": "3 identical CBCs; heaviest US expendable before SLS",
        "stages": [
            ("Center CBC",   "RS-68A",         1, 202000, 26760, "LOX/LH2 core; mass per CBC"),
            ("Side CBC (x2)","RS-68A",         1, 202000, 26760, "Two identical side boosters; mass per booster"),
            ("DCSS",         "RL-10B-2",       1,  27220,  3490, "Delta Cryogenic Second Stage; LOX/LH2"),
        ],
    },
    {
        "name": "Ariane 5 ECA", "country": "Europe", "operator": "Arianespace",
        "first_flight": 2002, "status": "Retired 2023",
        "payload_leo_kg": 21000, "payload_gto_kg": 10865, "payload_tli_kg": None,
        "glow_kg": 777000, "height_m": 59.0, "diameter_m": 5.4,
        "notes": "2 EAP solid strap-ons + EPC LOX/LH2 core + ESC-A LOX/LH2 upper",
        "stages": [
            ("EAP Solids (x2)","P80",          1, 237000, 30000, "Solid strap-on; each booster; actual Isp ~274s; P80 shown as proxy"),
            ("EPC Core",     "Vulcain 2",       1, 158000, 14700, "LOX/LH2 core stage"),
            ("ESC-A Upper",  "HM7B",           1,  14700,  1400, "LOX/LH2 upper stage; restartable"),
        ],
    },
    {
        "name": "Soyuz-2.1b", "country": "Russia", "operator": "Roscosmos",
        "first_flight": 2006, "status": "Active",
        "payload_leo_kg": 8200, "payload_gto_kg": 2800, "payload_tli_kg": None,
        "glow_kg": 312000, "height_m": 46.3, "diameter_m": 10.3,
        "notes": "Classic 3-stage parallel burn; boosters + core ignite at liftoff",
        "stages": [
            ("Block B/V/G/D (x4)","RD-107A",  1,  39600,  3784, "4 strap-on boosters; burn with core; mass per booster"),
            ("Block A Core",  "RD-108A",       1,  90360,  6545, "Burns simultaneously with boosters from liftoff"),
            ("Block I (S3)",  "RD-0124",       4,  25400,  2355, "4-chamber upper stage; high Isp"),
        ],
    },
    {
        "name": "Saturn V", "country": "USA", "operator": "NASA",
        "first_flight": 1967, "status": "Retired 1973",
        "payload_leo_kg": 130000, "payload_gto_kg": None, "payload_tli_kg": 48600,
        "glow_kg": 2970000, "height_m": 110.6, "diameter_m": 10.1,
        "notes": "Apollo program; still the most powerful rocket ever flown",
        "stages": [
            ("S-IC First",   "F-1",            5,2077000,130000, "LOX/RP-1; 5 F-1 engines"),
            ("S-II Second",  "J-2",            5, 426000, 36200, "LOX/LH2; 5 J-2 engines"),
            ("S-IVB Third",  "J-2",            1, 106000, 10000, "LOX/LH2; 1 J-2; also used as TMI stage"),
        ],
    },
    {
        "name": "SLS Block 1", "country": "USA", "operator": "NASA",
        "first_flight": 2022, "status": "Active",
        "payload_leo_kg": 95000, "payload_gto_kg": None, "payload_tli_kg": 27000,
        "glow_kg": 2608000, "height_m": 98.0, "diameter_m": 8.4,
        "notes": "Artemis program; RS-25 core + 2 SLS SRBs + ICPS upper stage",
        "stages": [
            ("SLS SRB (x2)",  "SRB (SLS 5-seg)",1,590000, 90718, "5-segment solid strap-on; mass per booster"),
            ("Core Stage",    "RS-25 (SSME)",  4, 987000, 85000, "LOX/LH2; 4 RS-25D engines"),
            ("ICPS Upper",    "RL-10C-3",      1,  27220,  3500, "Interim Cryogenic Propulsion Stage; LOX/LH2"),
        ],
    },
    {
        "name": "New Glenn (7x2 Baseline)", "country": "USA", "operator": "Blue Origin",
        "first_flight": 2025, "status": "Active",
        "payload_leo_kg": 45000, "payload_gto_kg": 13000, "payload_tli_kg": None,
        "glow_kg": 1000000, "height_m": 98.0, "diameter_m": 7.0,
        "notes": "First flight Jan 2025; 7 BE-4 + 2 BE-3U; 7m fairing; Blue Origin claimed 45t LEO / 13t GTO",
        "stages": [
            ("S1 First",     "BE-4",           7, 770000, 55000, "LOX/CH4; 7 BE-4 baseline engines; reusable"),
            ("S2 Upper",     "BE-3U",          2,  75000, 10000, "LOX/LH2; 2 BE-3U baseline engines"),
        ],
    },
    {
        "name": "New Glenn Block 2 (7x2 Upgraded)", "country": "USA", "operator": "Blue Origin",
        "first_flight": 2026, "status": "Development",
        "payload_leo_kg": None, "payload_gto_kg": None, "payload_tli_kg": None,
        "glow_kg": 1000000, "height_m": 98.0, "diameter_m": 7.0,
        "notes": "Upgraded BE-4 Block 2 (2,847 kN SL w/ subcooling) + BE-3U Block 2 (890 kN); no updated payload figures released by Blue Origin",
        "stages": [
            ("S1 First",     "BE-4 Block 2",   7, 770000, 55000, "LOX/CH4; 7 BE-4 Block 2; 640,000 lbf each with subcooling; reusable"),
            ("S2 Upper",     "BE-3U Block 2",  2,  75000, 10000, "LOX/LH2; 2 BE-3U Block 2; 200,000 lbf each"),
        ],
    },
    {
        "name": "New Glenn 9x4", "country": "USA", "operator": "Blue Origin",
        "first_flight": 2027, "status": "Development",
        "payload_leo_kg": 70000, "payload_gto_kg": 14000, "payload_tli_kg": 20000,
        "glow_kg": 1044000, "height_m": 120.0, "diameter_m": 8.7,
        "notes": "Heavy-lift variant; 9 BE-4 Block 2 S1 + 4 BE-3U Block 2 S2; 8.7m fairing (body ~7m); ~120m total height; "
                 "S2-biased % scaling from 7x2 (S1 ×1.10, S2 ×1.48); "
                 "grav loss coeff 1400 (vs 1750 for 7x2) — higher TWR enables more aggressive ascent trajectory",
        "stages": [
            ("S1 First",     "BE-4 Block 2",   9, 847000, 68000, "LOX/CH4; 9 BE-4 Block 2; 847t = 770t × 1.10; reusable"),
            ("S2 Upper",     "BE-3U Block 2",  4, 111000, 18000, "LOX/LH2; 4 BE-3U Block 2; 111t = 75t × 1.48"),
        ],
    },
    {
        "name": "New Glenn 9x4 (Northstar)", "country": "USA", "operator": "Blue Origin",
        "first_flight": 2027, "status": "Development",
        "payload_leo_kg": 70000, "payload_gto_kg": 14000, "payload_tli_kg": 20000,
        "glow_kg": 1129000, "height_m": 120.0, "diameter_m": 8.7,
        "notes": "9x4 re-estimated with 9 Northstar engines (715,000 lbf SL, 30% lighter than BE-4); "
                 "prop masses back-solved in analyze_9x4_northstar.py to Blue Origin's 70t LEO / 20t TLI claims "
                 "using a loss model calibrated on the 7x2 45t LEO claim; S1/S2 masses are estimates",
        "stages": [
            ("S1 First",     "Northstar",      9, 919000, 54400, "LOX/CH4; 9 Northstar; 919t nominal + 3.5% subcooling; dry mass back-solved; reusable"),
            ("S2 Upper",     "BE-3U Block 2",  4, 137000, 19300, "LOX/LH2; 4 BE-3U Block 2; 137t back-solved"),
        ],
    },
    {
        "name": "Starship (IFT-6 config)", "country": "USA", "operator": "SpaceX",
        "first_flight": 2023, "status": "Development",
        "payload_leo_kg": 150000, "payload_gto_kg": None, "payload_tli_kg": None,
        "glow_kg": 5000000, "height_m": 121.0, "diameter_m": 9.0,
        "notes": "Fully reusable target; 150t LEO is aspirational with full reuse; ~200t expendable claimed",
        "stages": [
            ("Super Heavy",  "Raptor 2",      33,3400000,200000, "LOX/CH4; 33 Raptor engines; reusable booster"),
            ("Starship",     "Raptor 2",       6,1200000,100000, "LOX/CH4; 3 vac + 3 SL Raptors; reusable ship"),
        ],
    },
    {
        "name": "Vulcan Centaur", "country": "USA", "operator": "ULA",
        "first_flight": 2024, "status": "Active",
        "payload_leo_kg": 27200, "payload_gto_kg": 9250, "payload_tli_kg": None,
        "glow_kg": 547000, "height_m": 61.6, "diameter_m": 5.4,
        "notes": "Replaces Atlas V and Delta IV; 2 or 4 GEM-63XL strap-ons optional",
        "stages": [
            ("First Stage",  "BE-4",           2, 406000, 22400, "LOX/CH4; 2 BE-4 engines"),
            ("Centaur V",    "RL-10C-1",       2,  54000,  6300, "LOX/LH2; 1 or 2 RL-10C-1; dual-engine shown"),
        ],
    },
]

# ── Propellant Properties ──────────────────────────────────────────────────────
# name, oxidizer, fuel, ox_density_kg_m3, fuel_density_kg_m3, MR_ox_fuel,
# typical_isp_vac_s, typical_isp_sl_s, storable, notes
PROPELLANTS = [
    ("LOX/RP-1",      "LOX",  "RP-1 (Kerosene)", 1141,  820, 2.56, 330, 290, False, "Dense; good SL performance; not storable; workhorse combo"),
    ("LOX/LH2",       "LOX",  "LH2 (Hydrogen)",  1141,   71, 5.50, 455, 380, False, "Highest Isp; low density challenges; large tanks required"),
    ("LOX/CH4",       "LOX",  "CH4 (Methane)",   1141,  423, 3.55, 360, 330, False, "Good Isp + density; reusability-friendly; ISRU potential"),
    ("NTO/MMH",       "NTO",  "MMH",             1450,  880, 1.65, 320, None, True, "Hypergolic; storable; common for upper stages and RCS"),
    ("NTO/UDMH",      "NTO",  "UDMH",            1450,  793, 2.60, 310, None, True, "Hypergolic; storable; Russian/Chinese first stages"),
    ("NTO/Aerozine-50","NTO", "Aerozine-50",     1450,  900, 1.90, 320, None, True, "50/50 N2H4+UDMH; Titan II/IV, Apollo CSM SPS"),
    ("Solid HTPB",    "AP",   "HTPB Binder",     1790, 1050, 3.50, 285, 270, True,  "High density; simple ops; Isp 265-295s depending on formulation"),
    ("Solid PBAN",    "AP",   "PBAN Binder",     1790, 1000, 3.50, 270, 250, True,  "Shuttle/SLS SRB propellant; similar to HTPB"),
    ("H2O2/RP-1",     "H2O2","RP-1",             1440,  820, 7.00, 300, 270, False, "Green propellant; lower Isp than LOX/RP-1; non-cryogenic oxidizer"),
    ("LOX/Ethanol",   "LOX",  "Ethanol",         1141,  789, 1.30, 311, 260, False, "Historical (V-2, Redstone); largely superseded by RP-1"),
]

# ── Subsystem Mass Fraction Model ──────────────────────────────────────────────
# Fractions are expressed as % of propellant mass, per propellant class
# Derived from analysis of real vehicles (Falcon 9, Atlas V, Saturn V, Ariane 5, Soyuz)
# Engine mass is handled separately (from engine DB x engine count)
SUBSYSTEMS = [
    # subsystem, LOX/RP1%, LOX/LH2%, LOX/CH4%, Solid%, Hypergolic%, description
    ("Tank Structure",         3.5, 6.5, 3.8, 1.8, 2.5, "Propellant tanks, interstage, skirt structure; LH2 needs insulation"),
    ("Pressurization System",  0.4, 0.6, 0.4, 0.0, 0.5, "He pressurant tanks, regulators, lines; N/A for solids"),
    ("Feed System / Plumbing", 0.8, 1.0, 0.8, 0.2, 0.6, "Turbopumps, manifolds, valves, lines (exclusive of engine)"),
    ("Avionics / GNC",         0.3, 0.4, 0.3, 0.2, 0.4, "Flight computer, IMU, GPS, telemetry; typically fixed-ish mass"),
    ("Power System",           0.2, 0.3, 0.2, 0.1, 0.2, "Batteries, power distribution, wiring harness"),
    ("Thermal Control",        0.2, 1.0, 0.3, 0.1, 0.3, "MLI blankets, heaters, paint; LH2 stages significantly higher"),
    ("RCS / ACS",              0.3, 0.4, 0.3, 0.1, 0.2, "Reaction control thrusters, propellant, lines"),
    ("Separation System",      0.1, 0.1, 0.1, 0.1, 0.1, "Stage separation clampband / ordnance; fairing sep if applicable"),
    ("Recovery Hardware",      0.0, 0.0, 0.0, 0.0, 0.0, "Landing legs, grid fins, parachutes — add manually if reusable"),
    ("Residuals / Margin",     0.5, 0.8, 0.5, 0.3, 0.5, "Unusable propellant, trapped fluids, 3-5% mass growth margin"),
]


# ── Payload / Spacecraft Configs ───────────────────────────────────────────────
PAYLOAD_BLUE_MOON_MK1 = {
    "name":       "Blue Moon Mk1",
    "mass_kg":    21_000,   # declared wet mass
    "adapter_kg":    200,
}


# ── Vehicle Analysis Configs ───────────────────────────────────────────────────
# Each entry drives build_payload_analysis() and produces one output file.
# Stage types:
#   "standard"      — burns sequentially; standard Tsiolkovsky treatment.
#   "parallel_core" — strap-ons/SRBs fire at liftoff alongside a liquid core,
#                     then jettison at booster burnout. Phase 1 uses the
#                     thrust-weighted blended Isp; Phase 2 is core-only.
#
# Gravity / drag loss model (matches Vehicle Design sheet defaults):
#   gravity_loss = grav_loss_coeff / TWR_liftoff + traj_penalty_ms
#   drag_loss    = drag_loss_ms (constant)
#   losses       = gravity_loss + drag_loss
#
# Notes on GEM-63XL (Vulcan VC6):
#   Engine DB entry records peak SL thrust (1,930 kN) and Isp (252/279 s).
#   Propellant mass (35,100 kg/booster) and dry mass (7,000 kg/booster) are
#   from Northrop Grumman published data.  The DB thrust value represents peak
#   thrust; solid motors taper, so effective average burn time via our model
#   (~45 s) is shorter than the published ~94 s.  The ΔV is still correct
#   because it depends on mass ratio × Isp, not burn time.
#
# Notes on SLS SRBs:
#   Engine DB entry for "SRB (SLS 5-seg)" has engine mass = 90,718 kg (dry
#   casing); propellant (590,000 kg per booster) is tracked separately here.

ANALYSIS_NG_BLOCK2_7X2 = {
    "vehicle_name":    "New Glenn Block 2 (7x2)",
    "output_file":     "Blue_Moon_Mk1_NG_Block2_7x2.xlsx",
    "grav_loss_coeff": 1750,
    "traj_penalty_ms": 500,    # recovery pitch-over penalty
    "drag_loss_ms":    120,
    "stages": [
        {
            "name": "S1 First", "type": "standard",
            "engine": "BE-4 Block 2", "n_engines": 7,
            "prop_mass_kg": 770_000, "subcooling_pct": 3.5,
            "dry_mass_kg": 55_000, "rec_dv_ms": 500, "atm_fraction": 0.6,
        },
        {
            "name": "S2 Upper", "type": "standard",
            "engine": "BE-3U Block 2", "n_engines": 2,
            "prop_mass_kg": 75_000, "subcooling_pct": 0.0,
            "dry_mass_kg": 10_000, "rec_dv_ms": 0, "atm_fraction": 1.0,
        },
    ],
}

ANALYSIS_VULCAN_VC6 = {
    "vehicle_name":    "Vulcan Centaur VC6",
    "output_file":     "Blue_Moon_Mk1_Vulcan_VC6.xlsx",
    "grav_loss_coeff": 1750,
    "traj_penalty_ms": 0,      # expendable; no recovery pitch-over
    "drag_loss_ms":    120,
    "stages": [
        {
            "name": "S1 Core + GEM-63XL (×6)", "type": "parallel_core",
            "engine": "BE-4", "n_engines": 2,
            "prop_mass_kg": 406_000, "subcooling_pct": 0.0,
            "dry_mass_kg": 22_400, "rec_dv_ms": 0, "atm_fraction": 0.6,
            "boosters": {
                "engine": "GEM-63XL", "n_engines": 6,
                "prop_mass_kg": 210_600,   # 6 × 35,100 kg (NGA published)
                "dry_mass_kg":  42_000,    # 6 × 7,000 kg (Engine DB)
                "atm_fraction": 0.9,
            },
        },
        {
            "name": "Centaur V", "type": "standard",
            "engine": "RL-10C-1", "n_engines": 2,
            "prop_mass_kg": 54_000, "subcooling_pct": 0.0,
            "dry_mass_kg": 6_300, "rec_dv_ms": 0, "atm_fraction": 1.0,
        },
    ],
}

# ── Vehicle Design Sheet Defaults ─────────────────────────────────────────────
# Passed to build_design(wb, vd) to pre-populate the yellow input cells.
# All values are user-editable in the spreadsheet once generated.

NG_7X2_DESIGN = {
    "vehicle_name": "New Glenn Block 2 (7x2)",
    "traj_penalty":  500,
    "s_names":       ["S1 First",     "S2 Upper",     ""],
    "s_engines":     ["BE-4 Block 2", "BE-3U Block 2",""],
    "s_n_eng":       [7,               2,              ""],
    "s_prop_type":   ["LOX/CH4",       "LOX/LH2",      ""],
    "s_prop":        [770_000,          75_000,         ""],
    "s_subcool":     [3.5,              0.0,             0.0],
    "s_rec_dv":      [500,              0,               0],
    "s_atm_frac":    [0.6,              1.0,             1.0],
    "s_leo_frac":    [0.60,             0.40,            ""],
    "s_gto_frac":    [0.55,             0.45,            ""],
    "s_dry_ovr":     [55_000,           10_000,          ""],
}

NG_9X4_DESIGN = {
    "vehicle_name": "New Glenn 9x4",
    "traj_penalty":  500,
    "s_names":       ["S1 First",     "S2 Upper",     ""],
    "s_engines":     ["BE-4 Block 2", "BE-3U Block 2",""],
    "s_n_eng":       [9,               4,              ""],
    "s_prop_type":   ["LOX/CH4",       "LOX/LH2",      ""],
    "s_prop":        [847_000,          111_000,         ""],   # 7x2 × 1.10/1.48 (S2-biased split: S1+10%, S2+48%)
    "s_subcool":     [3.5,              0.0,             0.0],
    "s_rec_dv":      [300,              0,               0],
    "s_atm_frac":    [0.6,              1.0,             1.0],
    "s_leo_frac":    [0.60,             0.40,            ""],
    "s_gto_frac":    [0.55,             0.45,            ""],
    "s_dry_ovr":     [68_000,           18_000,          ""],
}


NG_9X4_NORTHSTAR_DESIGN = {
    # Back-solved in analyze_9x4_northstar.py to Blue Origin's 70t LEO / 20t TLI claims,
    # with losses calibrated on the 7x2's 45t LEO claim (1750/TWR + 120 drag + 65 trajectory).
    # ΔV fractions are set so the sheet's upper-stage ΔV equals the model's actual S2 ΔV.
    "vehicle_name": "New Glenn 9x4 Northstar",
    "traj_penalty":  65,
    "s_names":       ["S1 First",     "S2 Upper",     ""],
    "s_engines":     ["Northstar",    "BE-3U Block 2",""],
    "s_n_eng":       [9,               4,              ""],
    "s_prop_type":   ["LOX/CH4",       "LOX/LH2",      ""],
    "s_prop":        [919_000,          137_000,        ""],
    "s_subcool":     [3.5,              0.0,             0.0],
    "s_rec_dv":      [500,              0,               0],
    "s_atm_frac":    [0.6,              1.0,             1.0],
    "s_leo_frac":    [0.5424,           0.4576,          ""],
    "s_gto_frac":    [0.4564,           0.5436,          ""],
    "s_dry_ovr":     [54_400,           19_300,          ""],
}

ANALYSIS_SLS_BLOCK1 = {
    "vehicle_name":    "SLS Block 1",
    "output_file":     "Blue_Moon_Mk1_SLS_Block1.xlsx",
    "grav_loss_coeff": 1750,
    "traj_penalty_ms": 0,
    "drag_loss_ms":    120,
    "stages": [
        {
            "name": "Core Stage + SRBs (×2)", "type": "parallel_core",
            "engine": "RS-25 (SSME)", "n_engines": 4,
            "prop_mass_kg": 987_000, "subcooling_pct": 0.0,
            "dry_mass_kg": 85_000, "rec_dv_ms": 0, "atm_fraction": 0.5,
            "boosters": {
                "engine": "SRB (SLS 5-seg)", "n_engines": 2,
                "prop_mass_kg": 1_180_000,  # 2 × 590,000 kg
                "dry_mass_kg":    181_436,  # 2 × 90,718 kg (Engine DB dry casing)
                "atm_fraction": 0.8,
            },
        },
        {
            "name": "ICPS Upper", "type": "standard",
            "engine": "RL-10C-3", "n_engines": 1,
            "prop_mass_kg": 27_220, "subcooling_pct": 0.0,
            "dry_mass_kg": 3_500, "rec_dv_ms": 0, "atm_fraction": 1.0,
        },
    ],
}

ANALYSIS_NG_9X4 = {
    "vehicle_name":    "New Glenn 9x4",
    "output_file":     "Blue_Moon_Mk1_NG_9x4.xlsx",
    "grav_loss_coeff": 1400,   # lower coeff vs 7x2 (1750); 9x4 higher TWR allows more aggressive trajectory
    "traj_penalty_ms": 500,
    "drag_loss_ms":    120,
    "stages": [
        {
            "name": "S1 First", "type": "standard",
            "engine": "BE-4 Block 2", "n_engines": 9,
            "prop_mass_kg": 847_000, "subcooling_pct": 3.5,   # 7x2 × 1.10 (S2-biased split: S1+10%, S2+48%)
            "dry_mass_kg": 68_000, "rec_dv_ms": 300, "atm_fraction": 0.6,
        },
        {
            "name": "S2 Upper", "type": "standard",
            "engine": "BE-3U Block 2", "n_engines": 4,
            "prop_mass_kg": 111_000, "subcooling_pct": 0.0,   # 7x2 × 1.48 (S2-biased split: S1+10%, S2+48%)
            "dry_mass_kg": 18_000, "rec_dv_ms": 0, "atm_fraction": 1.0,
        },
    ],
}


ANALYSIS_NG_9X4_NORTHSTAR = {
    "vehicle_name":    "New Glenn 9x4 (Northstar)",
    "output_file":     "Blue_Moon_Mk1_NG_9x4_Northstar.xlsx",
    # Losses calibrated on the 7x2 45t LEO claim (see analyze_9x4_northstar.py):
    # 1750/TWR + K, K ≈ 185 m/s = 120 drag + 65 trajectory
    "grav_loss_coeff": 1750,
    "traj_penalty_ms": 65,
    "drag_loss_ms":    120,
    "stages": [
        {
            "name": "S1 First", "type": "standard",
            "engine": "Northstar", "n_engines": 9,
            "prop_mass_kg": 919_000, "subcooling_pct": 3.5,   # back-solved to 70t LEO / 20t TLI
            "dry_mass_kg": 54_400, "rec_dv_ms": 500, "atm_fraction": 0.6,
        },
        {
            "name": "S2 Upper", "type": "standard",
            "engine": "BE-3U Block 2", "n_engines": 4,
            "prop_mass_kg": 137_000, "subcooling_pct": 0.0,
            "dry_mass_kg": 19_300, "rec_dv_ms": 0, "atm_fraction": 1.0,
        },
    ],
}

# ── Analysis helpers ───────────────────────────────────────────────────────────
def lookup_engine(name):
    """Return engine dict from ENGINES list; raises ValueError if not found."""
    for eng in ENGINES:
        if eng[0] == name:
            return {
                "name":          eng[0],
                "vac_isp":       eng[3],
                "sl_isp":        eng[4],          # None for solid/vacuum-only
                "vac_thrust_kn": eng[5],
                "sl_thrust_kn":  eng[6],          # None for vacuum-only
                "mass_kg":       eng[7],
            }
    raise ValueError(f"Engine not found in ENGINES: '{name}'")


def blended_isp(engine, atm_fraction):
    """Effective Isp = SL_Isp + atm_fraction × (Vac_Isp − SL_Isp).
    Falls back to Vac_Isp when SL_Isp is unavailable (solids, upper stages)."""
    sl = engine["sl_isp"] or engine["vac_isp"]
    return sl + atm_fraction * (engine["vac_isp"] - sl)


def compute_dv_budget(vehicle_cfg, payload_cfg):
    """
    Compute the full ΔV budget for a given vehicle + payload combination.

    Returns a dict with keys:
        stages       – list of per-stage result dicts (in burn order)
        total_dv     – sum of all stage ΔVs (m/s)
        glow         – estimated GLOW using nominal prop masses, no payload (kg)
        twr          – liftoff TWR (stage masses only)
        grav_loss    – gravity loss (m/s)
        drag_loss    – drag loss (m/s)
        losses       – total losses (m/s)
        v_inject     – injection velocity at perigee (total_dv − losses, m/s)

    Stage result dict keys (both types):
        name, type, dv, m0, mf, isp_eff, rec_prop, eff_dry_asc,
        eff_prop, dry_mass, n_engines, engine
    Additional keys for "parallel_core":
        dv_phase1, dv_phase2, isp_parallel, prop_core_p1,
        booster_prop, booster_dry, booster_engine
    """
    payload   = payload_cfg["mass_kg"]
    adapter   = payload_cfg["adapter_kg"]
    stage_cfgs = vehicle_cfg["stages"]

    # ── Resolve engines and compute per-stage gross masses ────────────────────
    # Gross mass = eff_prop + dry  (+ booster prop + dry for parallel_core)
    # "above" mass for stage i = sum of stage_gross[i+1:] + payload + adapter
    stage_info = []   # resolved stage data
    stage_gross = []  # actual gross mass for "above" calculations

    for sc in stage_cfgs:
        eng = lookup_engine(sc["engine"])
        eff_prop = sc["prop_mass_kg"] * (1 + sc.get("subcooling_pct", 0) / 100)
        rec_prop = 0.0
        if sc.get("rec_dv_ms", 0) > 0:
            rec_prop = sc["dry_mass_kg"] * (
                math.exp(sc["rec_dv_ms"] / (eng["vac_isp"] * G0)) - 1)
        isp = blended_isp(eng, sc.get("atm_fraction", 1.0))

        info = {
            "name": sc["name"], "type": sc["type"],
            "engine": eng, "n_engines": sc["n_engines"],
            "prop_mass_kg": sc["prop_mass_kg"],   # nominal (for GLOW)
            "eff_prop": eff_prop, "rec_prop": rec_prop,
            "dry_mass": sc["dry_mass_kg"],
            "eff_dry_asc": sc["dry_mass_kg"] + rec_prop,
            "isp_eff": isp,
            "rec_dv_ms": sc.get("rec_dv_ms", 0),
        }

        if sc["type"] == "parallel_core":
            bsc  = sc["boosters"]
            beng = lookup_engine(bsc["engine"])
            bisp = blended_isp(beng, bsc.get("atm_fraction", 1.0))
            info["booster_engine"]  = beng
            info["booster_n_eng"]   = bsc["n_engines"]
            info["booster_prop"]    = bsc["prop_mass_kg"]
            info["booster_dry"]     = bsc["dry_mass_kg"]
            info["booster_isp_eff"] = bisp
            gross = (eff_prop + sc["dry_mass_kg"]
                     + bsc["prop_mass_kg"] + bsc["dry_mass_kg"])
        else:
            gross = eff_prop + sc["dry_mass_kg"]

        stage_info.append(info)
        stage_gross.append(gross)

    # ── Compute ΔV per stage ──────────────────────────────────────────────────
    results = []
    for i, si in enumerate(stage_info):
        above = sum(stage_gross[i + 1:]) + payload + adapter
        eng   = si["engine"]

        if si["type"] == "standard":
            m0 = si["eff_prop"] + si["dry_mass"] + above
            mf = si["eff_dry_asc"] + above
            dv = si["isp_eff"] * G0 * math.log(m0 / mf)
            results.append({**si, "m0": m0, "mf": mf, "dv": dv})

        elif si["type"] == "parallel_core":
            beng  = si["booster_engine"]
            bisp  = si["booster_isp_eff"]
            T_core  = si["n_engines"]      * (eng["sl_thrust_kn"]  or eng["vac_thrust_kn"])  * 1_000
            T_boost = si["booster_n_eng"]  * (beng["sl_thrust_kn"] or beng["vac_thrust_kn"]) * 1_000
            T_total = T_core + T_boost

            mdot_core  = T_core  / (si["isp_eff"] * G0)
            mdot_boost = T_boost / (bisp          * G0)
            t_boost    = si["booster_prop"] / mdot_boost
            prop_core_p1 = mdot_core * t_boost

            # Phase 1: core + boosters burn simultaneously
            isp_parallel = T_total / ((mdot_core + mdot_boost) * G0)
            m0_p1 = (si["eff_prop"] + si["dry_mass"]
                     + si["booster_prop"] + si["booster_dry"] + above)
            mf_p1 = m0_p1 - si["booster_prop"] - si["booster_dry"] - prop_core_p1
            dv_p1 = isp_parallel * G0 * math.log(m0_p1 / mf_p1)

            # Phase 2: core burns remaining propellant alone
            m0_p2 = mf_p1
            mf_p2 = si["eff_dry_asc"] + above
            dv_p2 = si["isp_eff"] * G0 * math.log(m0_p2 / mf_p2)

            dv = dv_p1 + dv_p2
            results.append({
                **si,
                "m0": m0_p1, "mf": mf_p2, "dv": dv,
                "dv_phase1": dv_p1, "dv_phase2": dv_p2,
                "isp_parallel": isp_parallel,
                "prop_core_p1": prop_core_p1,
            })

    # ── TWR & losses (use nominal prop for GLOW, consistent with design sheet) ─
    glow = sum(
        si["prop_mass_kg"] + si["dry_mass"]
        + (si.get("booster_prop", 0) + si.get("booster_dry", 0))
        for si in stage_info
    )
    s0   = stage_info[0]
    s0eng = s0["engine"]
    T_liftoff = s0["n_engines"] * (s0eng["sl_thrust_kn"] or s0eng["vac_thrust_kn"]) * 1_000
    if s0["type"] == "parallel_core":
        beng0 = s0["booster_engine"]
        T_liftoff += s0["booster_n_eng"] * (beng0["sl_thrust_kn"] or beng0["vac_thrust_kn"]) * 1_000

    twr       = T_liftoff / (glow * G0)
    grav_loss = vehicle_cfg["grav_loss_coeff"] / twr + vehicle_cfg.get("traj_penalty_ms", 0)
    drag_loss = vehicle_cfg.get("drag_loss_ms", 120)
    losses    = grav_loss + drag_loss
    total_dv  = sum(r["dv"] for r in results)
    v_inject  = total_dv - losses

    return {
        "stages":    results,
        "total_dv":  total_dv,
        "glow":      glow,
        "twr":       twr,
        "grav_loss": grav_loss,
        "drag_loss": drag_loss,
        "losses":    losses,
        "v_inject":  v_inject,
    }


def style_hdr(ws, row, col, text, span=1, dark=True):
    cell = ws.cell(row=row, column=col, value=text)
    bg = C["hdr"] if dark else C["sub"]
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.font = Font(bold=True, color="FFFFFF", size=11)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    return cell


def style_sec(ws, row, col, text, span=1):
    cell = ws.cell(row=row, column=col, value=text)
    cell.fill = PatternFill("solid", fgColor=C["sec"])
    cell.font = Font(bold=True, size=10, color="1F1F1F")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    return cell


def style_input(ws, row, col, value=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = PatternFill("solid", fgColor=C["input"])
    cell.font = Font(size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    return cell


def style_calc(ws, row, col, value=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = PatternFill("solid", fgColor=C["calc"])
    cell.font = Font(size=10, italic=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    return cell


def style_label(ws, row, col, text):
    cell = ws.cell(row=row, column=col, value=text)
    cell.fill = PatternFill("solid", fgColor=C["ltgray"])
    cell.font = Font(size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    return cell


def thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def apply_border_range(ws, min_row, max_row, min_col, max_col):
    b = thin_border()
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(r, c).border = b


def col(n):
    return get_column_letter(n)


# ─────────────────────────────────────────────────────────────────────────────
# SHEET: Engine Database
# ─────────────────────────────────────────────────────────────────────────────
def build_engine_db(wb):
    ws = wb.create_sheet("Engine DB")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    headers = [
        "Engine Name", "Vehicle Application", "Propellant", "Vac Isp (s)", "SL Isp (s)",
        "Vac Thrust (kN)", "SL Thrust (kN)", "Engine Mass (kg)", "Chamber Press (bar)",
        "Expansion Ratio", "Notes"
    ]
    widths = [18, 22, 12, 12, 12, 16, 16, 16, 18, 15, 45]

    style_hdr(ws, 1, 1, "ROCKET ENGINE DATABASE — Real-World Performance Data", span=len(headers))
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        style_hdr(ws, 2, ci, h, dark=False)
        ws.column_dimensions[col(ci)].width = w

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 30

    # Group by propellant
    groups = {}
    for eng in ENGINES:
        groups.setdefault(eng[2], []).append(eng)

    r = 3
    for prop_group, engines in groups.items():
        style_sec(ws, r, 1, f"── {prop_group} ──", span=len(headers))
        ws.row_dimensions[r].height = 18
        r += 1
        for eng in engines:
            for ci, val in enumerate(eng, 1):
                cell = ws.cell(r, ci, value=val if val is not None else "—")
                cell.font = Font(size=10)
                cell.alignment = Alignment(horizontal="center" if ci > 2 else "left",
                                           vertical="center", wrap_text=(ci == 11))
                if ci in (1, 2, 11):
                    cell.alignment = Alignment(horizontal="left", vertical="center",
                                               wrap_text=True, indent=1)
            apply_border_range(ws, r, r, 1, len(headers))
            ws.row_dimensions[r].height = 15
            r += 1

    ws.auto_filter.ref = f"A2:{col(len(headers))}2"
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# SHEET: Vehicle Database
# ─────────────────────────────────────────────────────────────────────────────
def build_vehicle_db(wb):
    ws = wb.create_sheet("Vehicle DB")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"

    style_hdr(ws, 1, 1, "REFERENCE LAUNCH VEHICLE DATABASE — Real Vehicle Mass & Performance Data", span=12)

    # Vehicle summary headers
    vh = ["Vehicle", "Country", "Operator", "First Flight", "Status",
          "Payload LEO (kg)", "Payload GTO (kg)", "Payload TLI (kg)",
          "GLOW (kg)", "Height (m)", "Diameter (m)", "Notes"]
    vw = [22, 10, 14, 12, 14, 16, 16, 16, 14, 11, 12, 45]
    for ci, (h, w) in enumerate(zip(vh, vw), 1):
        style_hdr(ws, 2, ci, h, dark=False)
        ws.column_dimensions[col(ci)].width = w

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 30

    r = 3
    for v in VEHICLES:
        row_vals = [v["name"], v["country"], v["operator"], v["first_flight"],
                    v["status"], v["payload_leo_kg"], v["payload_gto_kg"],
                    v.get("payload_tli_kg", None), v["glow_kg"],
                    v["height_m"], v["diameter_m"], v["notes"]]
        for ci, val in enumerate(row_vals, 1):
            cell = ws.cell(r, ci, value=val if val is not None else "—")
            cell.font = Font(bold=(ci == 1), size=10)
            cell.alignment = Alignment(horizontal="left" if ci in (1,2,3,5,12) else "center",
                                       vertical="center", wrap_text=(ci == 12), indent=1)
            cell.fill = PatternFill("solid", fgColor=C["ltgray"] if r % 2 == 0 else C["white"])
        apply_border_range(ws, r, r, 1, len(vh))
        ws.row_dimensions[r].height = 15
        r += 1

    # Stage breakdown section
    r += 1
    style_hdr(ws, r, 1, "PER-STAGE MASS BREAKDOWN", span=9)
    r += 1
    sh = ["Vehicle", "Stage", "Engine", "# Engines", "Propellant Mass (kg)",
          "Dry Mass (kg)", "Gross Mass (kg)", "Struct. Frac. (%)", "Notes"]
    sw = [22, 20, 18, 10, 20, 15, 15, 15, 45]
    for ci, (h, w) in enumerate(zip(sh, sw), 1):
        style_hdr(ws, r, ci, h, dark=False)
    ws.row_dimensions[r].height = 25
    r += 1

    for v in VEHICLES:
        for si, stg in enumerate(v["stages"]):
            sname, eng, n_eng, prop_mass, dry_mass, notes = stg
            gross = prop_mass + dry_mass
            sf = (dry_mass - 0) / prop_mass * 100  # simplified: dry/prop
            row_vals = [v["name"] if si == 0 else "", sname, eng, n_eng,
                        prop_mass, dry_mass, gross, round(sf, 2), notes]
            for ci, val in enumerate(row_vals, 1):
                cell = ws.cell(r, ci, value=val)
                cell.font = Font(bold=(ci == 1 and si == 0), size=10)
                cell.alignment = Alignment(horizontal="left" if ci in (1,2,3,9) else "center",
                                           vertical="center", wrap_text=(ci == 9), indent=1)
                cell.fill = PatternFill("solid", fgColor=C["ltgray"] if r % 2 == 0 else C["white"])
            apply_border_range(ws, r, r, 1, len(sh))
            ws.row_dimensions[r].height = 15
            r += 1

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# SHEET: Propellants
# ─────────────────────────────────────────────────────────────────────────────
def build_propellants(wb):
    ws = wb.create_sheet("Propellants")
    ws.sheet_view.showGridLines = False

    style_hdr(ws, 1, 1, "PROPELLANT PROPERTIES", span=10)

    headers = ["Propellant", "Oxidizer", "Fuel", "Ox Density (kg/m³)",
               "Fuel Density (kg/m³)", "Mix Ratio (O/F)", "Bulk Density* (kg/m³)",
               "Typical Vac Isp (s)", "Storable?", "Notes"]
    widths = [18, 12, 18, 18, 18, 15, 18, 18, 10, 50]
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        style_hdr(ws, 2, ci, h, dark=False)
        ws.column_dimensions[col(ci)].width = w

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 30

    for ri, p in enumerate(PROPELLANTS, 3):
        name, ox, fuel, ox_d, f_d, mr, isp_vac, isp_sl, storable, notes = p
        bulk = (ox_d * mr + f_d * 1) / (mr + 1)
        row_vals = [name, ox, fuel, ox_d, f_d, mr, round(bulk, 0),
                    isp_vac, "Yes" if storable else "No", notes]
        for ci, val in enumerate(row_vals, 1):
            cell = ws.cell(ri, ci, value=val)
            cell.font = Font(size=10)
            cell.alignment = Alignment(horizontal="left" if ci in (1,2,3,9,10) else "center",
                                       vertical="center", wrap_text=(ci == 10), indent=1)
            cell.fill = PatternFill("solid", fgColor=C["ltgray"] if ri % 2 == 0 else C["white"])
        apply_border_range(ws, ri, ri, 1, len(headers))
        ws.row_dimensions[ri].height = 15

    r = len(PROPELLANTS) + 4
    ws.cell(r, 1, value="* Bulk density = mass-weighted average of oxidizer + fuel at stated mixture ratio").font = Font(italic=True, size=9, color="555555")

    # Subsystem mass fraction model
    r += 2
    style_hdr(ws, r, 1, "SUBSYSTEM MASS FRACTION MODEL — Fractions as % of Propellant Mass", span=7)
    r += 1
    sh = ["Subsystem", "LOX/RP-1 (%)", "LOX/LH2 (%)", "LOX/CH4 (%)",
          "Solid (%)", "Hypergolic (%)", "Description"]
    sw2 = [22, 13, 13, 13, 10, 14, 55]
    for ci, (h, w) in enumerate(zip(sh, sw2), 1):
        style_hdr(ws, r, ci, h, dark=False)
        ws.column_dimensions[col(ci)].width = max(ws.column_dimensions[col(ci)].width, w)
    r += 1

    totals = [0.0] * 5
    for sub in SUBSYSTEMS:
        name_s, rp1, lh2, ch4, sol, hyp, desc = sub
        for ci, val in enumerate([name_s, rp1, lh2, ch4, sol, hyp, desc], 1):
            cell = ws.cell(r, ci, value=val)
            cell.font = Font(size=10)
            cell.alignment = Alignment(horizontal="left" if ci in (1, 7) else "center",
                                       vertical="center", wrap_text=(ci == 7), indent=1)
            cell.fill = PatternFill("solid", fgColor=C["ltgray"] if r % 2 == 0 else C["white"])
        apply_border_range(ws, r, r, 1, 7)
        ws.row_dimensions[r].height = 15
        totals = [totals[i] + v for i, v in enumerate([rp1, lh2, ch4, sol, hyp])]
        r += 1

    # Totals row
    ws.cell(r, 1, value="TOTAL (excl. engines)").font = Font(bold=True, size=10)
    for ci, t in enumerate(totals, 2):
        cell = ws.cell(r, ci, value=round(t, 2))
        cell.font = Font(bold=True, size=10)
        cell.fill = PatternFill("solid", fgColor=C["sec"])
        cell.alignment = Alignment(horizontal="center")
    apply_border_range(ws, r, r, 1, 7)

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# SHEET: Vehicle Design (main working sheet)
# ─────────────────────────────────────────────────────────────────────────────
def build_design(wb, vd=None):
    vd = vd or NG_7X2_DESIGN
    ws = wb.create_sheet("Vehicle Design", 0)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    # Column layout
    # A: labels   B: value/unit   C: (blank buffer)
    # D..J: Stage 1-3 (three cols each: label, value, spacer)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 4
    for i in range(MAX_STAGES):
        ws.column_dimensions[col(4 + i * 3)].width = 18
        ws.column_dimensions[col(5 + i * 3)].width = 16
        ws.column_dimensions[col(6 + i * 3)].width = 3

    # Row 1: Title
    style_hdr(ws, 1, 1, "LAUNCH VEHICLE DESIGN TOOL", span=4 + MAX_STAGES * 3)
    ws.row_dimensions[1].height = 28

    # Row 2: Legend
    ws.merge_cells(f"A2:{col(4 + MAX_STAGES * 3)}2")
    ws.cell(2, 1, "Yellow = Input   |   Blue = Calculated   |   "
                  "Engine name must match Engine DB exactly   |   "
                  "ΔV Allocation per stage must sum to Mission ΔV").font = Font(size=9, italic=True)

    # Row 3: Stage column headers
    ws.cell(3, 1, value="PARAMETER").font = Font(bold=True, size=10)
    ws.cell(3, 2, value="UNITS").font = Font(bold=True, size=10)
    ws.cell(3, 1).fill = PatternFill("solid", fgColor=C["sec"])
    ws.cell(3, 2).fill = PatternFill("solid", fgColor=C["sec"])
    ws.row_dimensions[3].height = 20

    for si in range(MAX_STAGES):
        c_label = 4 + si * 3
        c_val   = 5 + si * 3
        ws.merge_cells(start_row=3, start_column=c_label,
                       end_row=3, end_column=c_val)
        cell = ws.cell(3, c_label, value=f"Stage {si+1}")
        cell.fill = PatternFill("solid", fgColor=C["stage"][si])
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── Section: Mission ──────────────────────────────────────────────────────
    r = 5
    style_sec(ws, r, 1, "MISSION REQUIREMENTS", span=2)
    r += 1

    mission_inputs = [
        ("Vehicle Name",                   "text",  vd["vehicle_name"],  None),
        ("Target Orbit",                   "text",  "LEO 400km / GTO",   None),
        ("Payload Adapter Mass",           "kg",    200,          "Separation hardware mass between payload and vehicle"),
        ("Orbital Velocity (LEO)",         "m/s",   7784,         "Target orbital velocity — 7784 m/s for 400km LEO; 7726 m/s for 200km"),
        ("Drag Loss",                      "m/s",   120,          "Aerodynamic drag loss — typically 100-150 m/s; relatively insensitive to thrust"),
        ("Gravity Loss Coefficient",          "—",    1750,         "Base gravity loss = coeff / TWR_liftoff; ~1750 typical for expendable; increase to match reusable vehicles"),
        ("Trajectory Penalty",               "m/s",  vd["traj_penalty"], "Additive gravity loss for recovery pitch profile; 0=expendable, calibrate to match known vehicle — ~500 m/s estimated for NG Block 2 downrange recovery"),
        ("GTO Orbital ΔV",                 "m/s",   10400,        "Velocity needed at GTO perigee — ~10,200-10,500 m/s depending on inclination and target apogee"),
        ("Number of Stages",               "1-3",   2,            "Active stages (enter 1-3)"),
    ]

    mission_rows = {}
    for label, unit, default, note in mission_inputs:
        ws.cell(r, 1, value=label).font = Font(size=10)
        ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
        style_input(ws, r, 2, default)
        ws.cell(r, 2).number_format = "#,##0" if unit in ("kg", "m/s") else "@"
        if note:
            ws.cell(r, 3, value=f"← {note}").font = Font(size=8, italic=True, color="777777")
            ws.merge_cells(start_row=r, start_column=3,
                           end_row=r, end_column=4 + MAX_STAGES * 3 - 1)
        mission_rows[label] = r
        r += 1

    adapter_row           = mission_rows["Payload Adapter Mass"]
    orb_vel_row           = mission_rows["Orbital Velocity (LEO)"]
    drag_loss_row         = mission_rows["Drag Loss"]
    grav_coeff_row        = mission_rows["Gravity Loss Coefficient"]
    traj_penalty_row      = mission_rows["Trajectory Penalty"]
    gto_orb_dv_row        = mission_rows["GTO Orbital ΔV"]


    # LEO and GTO mission ΔV will be calculated after stage config (needs TWR)
    # placeholders assigned after stage config loop
    dv_budget_row     = None  # set later
    dv_gto_budget_row = None  # set later

    # ── Section: Stage Configuration ─────────────────────────────────────────
    r += 1
    style_sec(ws, r, 1, "STAGE CONFIGURATION", span=2 + MAX_STAGES * 3)
    ws.row_dimensions[r].height = 18
    r += 1

    stage_rows = {}
    stage_params = [
        ("Stage Name",               "text",  vd["s_names"]),
        ("Engine (from Engine DB)",  "name",  vd["s_engines"]),
        ("Number of Engines",        "count", vd["s_n_eng"]),
        ("Propellant Combination",   "type",  vd["s_prop_type"]),
        ("Propellant Mass (nominal)", "kg",   vd["s_prop"]),
        ("Subcooling Gain (%)",       "%",    vd["s_subcool"]),
        ("Effective Propellant Mass", "kg",   None),   # nominal × (1 + gain/100)
        ("Recovery ΔV Reserve (m/s)", "m/s", vd["s_rec_dv"]),
        ("Recovery Propellant Reserved", "kg", None),  # deferred
        ("LEO ΔV Fraction",           "0-1", vd["s_leo_frac"]),
        ("LEO ΔV Allocation",         "m/s", None),   # deferred
        ("Vac Isp Override",          "s",   ["", "", ""]),
        ("Vac Isp (from DB)",         "s",   None),   # VLOOKUP col 4
        ("SL Isp (from DB)",          "s",   None),   # VLOOKUP col 5
        ("Atm Fraction (0-1)",        "—",   vd["s_atm_frac"]),
        ("Effective Isp (used)",      "s",   None),   # SL + frac*(Vac-SL), or override
        ("GTO ΔV Fraction",           "0-1", vd["s_gto_frac"]),
        ("GTO ΔV Allocation",         "m/s", None),   # deferred
    ]

    for param, unit, defaults in stage_params:
        ws.cell(r, 1, value=param).font = Font(size=10)
        ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
        ws.cell(r, 2, value=unit).font = Font(size=9, color="777777")
        ws.cell(r, 2).alignment = Alignment(horizontal="center")
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
        stage_rows[param] = r

        for si in range(MAX_STAGES):
            c_val = 5 + si * 3
            if param in ("LEO ΔV Allocation", "GTO ΔV Allocation"):
                # Formula filled after dv_budget_row / dv_gto_budget_row is known (deferred)
                frac_key = "LEO ΔV Fraction" if param == "LEO ΔV Allocation" else "GTO ΔV Fraction"
                frac_c = f"{col(c_val)}{stage_rows[frac_key]}"
                cell = style_calc(ws, r, c_val, value="")
                cell.number_format = "#,##0"
                if not hasattr(ws, '_dv_alloc_deferred'):
                    ws._dv_alloc_deferred = []
                ws._dv_alloc_deferred.append((r, c_val, frac_c, param))
            elif param == "Effective Propellant Mass":
                nom_c  = f"{col(c_val)}{stage_rows['Propellant Mass (nominal)']}"
                gain_c = f"{col(c_val)}{stage_rows['Subcooling Gain (%)']}"
                formula = f'=IFERROR({nom_c}*(1+{gain_c}/100),"")'
                cell = style_calc(ws, r, c_val, value=formula)
                cell.font = Font(bold=True, size=10, italic=True)
                cell.number_format = "#,##0"
            elif param == "Recovery Propellant Reserved":
                # Deferred — formula depends on eff_dry_row which isn't set yet
                cell = style_calc(ws, r, c_val, value="")
                cell.number_format = "#,##0"
                if not hasattr(ws, '_recovery_prop_deferred'):
                    ws._recovery_prop_deferred = []
                ws._recovery_prop_deferred.append((r, c_val))
            elif param == "Vac Isp (from DB)":
                eng_cell = f"{col(c_val)}{stage_rows['Engine (from Engine DB)']}"
                ovr_cell = f"{col(c_val)}{stage_rows['Vac Isp Override']}"
                formula = (f'=IF({ovr_cell}<>"",{ovr_cell},'
                           f'IFERROR(VLOOKUP({eng_cell},\'Engine DB\'!$A:$K,4,0),"Enter engine"))')
                cell = style_calc(ws, r, c_val, value=formula)
                cell.number_format = "0.0"
            elif param == "SL Isp (from DB)":
                eng_cell = f"{col(c_val)}{stage_rows['Engine (from Engine DB)']}"
                # Col 5 = SL Isp; returns "—" if not applicable (upper stages, solids)
                formula = f'=IFERROR(VLOOKUP({eng_cell},\'Engine DB\'!$A:$K,5,0),"—")'
                cell = style_calc(ws, r, c_val, value=formula)
                cell.number_format = "0.0"
            elif param == "Effective Isp (used)":
                vac_c  = f"{col(c_val)}{stage_rows['Vac Isp (from DB)']}"
                sl_c   = f"{col(c_val)}{stage_rows['SL Isp (from DB)']}"
                frac_c = f"{col(c_val)}{stage_rows['Atm Fraction (0-1)']}"
                # If SL Isp unavailable (solid/upper), use Vac Isp directly
                formula = (f'=IFERROR('
                           f'IF(ISNUMBER({sl_c}),'
                           f'{sl_c}+{frac_c}*({vac_c}-{sl_c}),'
                           f'{vac_c}),{vac_c})')
                cell = style_calc(ws, r, c_val, value=formula)
                cell.font = Font(bold=True, size=10, italic=True)
                cell.number_format = "0.0"
            else:
                val = defaults[si] if si < len(defaults) else ""
                cell = style_input(ws, r, c_val, val)
                cell.number_format = "#,##0" if unit in ("kg", "m/s") else "@"
            cell.fill = PatternFill("solid", fgColor=C["calc"] if defaults is None or param in ("Vac Isp (from DB)", "SL Isp (from DB)", "Effective Isp (used)", "Effective Propellant Mass", "Recovery Propellant Reserved", "LEO ΔV Allocation", "GTO ΔV Allocation") else C["input"])

        r += 1

    # LEO ΔV allocation check row
    dv_alloc_row = stage_rows["LEO ΔV Allocation"]

    # ── Mission ΔV Analysis (calculated from TWR) ─────────────────────────────
    # TWR uses sum of all stage gross masses (prop + dry) as GLOW approximation.
    # This avoids a circular reference with payload. Error is small (<3%) since
    # payload is ~2-4% of GLOW. Stage 1 SL thrust used for gravity loss calculation.
    r += 1
    style_sec(ws, r, 1, "MISSION ΔV ANALYSIS  —  Gravity loss computed from liftoff TWR", span=2 + MAX_STAGES * 3)
    ws.row_dimensions[r].height = 18
    r += 1

    # Stage gross masses (prop nominal + dry) — needed for GLOW estimate
    # We use nominal prop mass here (before subcooling) as that's the tank-full condition at liftoff
    nom_prop_row = stage_rows["Propellant Mass (nominal)"]

    def make_label_calc(ws, r, label, unit):
        ws.cell(r, 1, value=label).font = Font(size=10)
        ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
        ws.cell(r, 2, value=unit).font = Font(size=9, color="777777")
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
        ws.cell(r, 2).alignment = Alignment(horizontal="center")

    # Estimated GLOW (sum of all stage prop+dry, no payload — conservative TWR)
    make_label_calc(ws, r, "Est. GLOW (stage masses only, no payload)", "kg")
    glow_est_row = r
    # We need eff_dry_row but it's not defined yet — use a forward reference string
    # We'll fill this formula after eff_dry_row is assigned, so store the row and come back
    glow_est_r = r  # store row index; formula filled after subsystem section
    r += 1

    # Stage 1 liftoff TWR
    make_label_calc(ws, r, "Stage 1 Liftoff TWR (est., no payload)", "—")
    twr_est_row = r
    twr_est_r = r
    r += 1

    # Gravity loss = coeff / TWR + trajectory penalty
    make_label_calc(ws, r, "Gravity Loss (= coeff / TWR + trajectory penalty)", "m/s")
    grav_loss_row = r
    ws.cell(r, 2, value=f"=IFERROR(B{grav_coeff_row}/B{twr_est_row}+B{traj_penalty_row},\"\")").font = Font(size=10, italic=True)
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["calc"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    ws.cell(r, 2).number_format = "#,##0"
    r += 1

    # LEO Mission ΔV = orbital velocity + drag + gravity loss
    make_label_calc(ws, r, "LEO Mission ΔV (calculated)", "m/s")
    leo_dv_calc_row = r
    ws.cell(r, 2, value=f"=IFERROR(B{orb_vel_row}+B{drag_loss_row}+B{grav_loss_row},\"\")").font = Font(bold=True, size=10, italic=True)
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["calc"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    ws.cell(r, 2).number_format = "#,##0"
    dv_budget_row = leo_dv_calc_row
    r += 1

    # GTO Mission ΔV = GTO orbital ΔV + drag + gravity loss
    make_label_calc(ws, r, "GTO Mission ΔV (calculated)", "m/s")
    gto_dv_calc_row = r
    ws.cell(r, 2, value=f"=IFERROR(B{gto_orb_dv_row}+B{drag_loss_row}+B{grav_loss_row},\"\")").font = Font(bold=True, size=10, italic=True)
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["calc"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    ws.cell(r, 2).number_format = "#,##0"
    dv_gto_budget_row = gto_dv_calc_row
    r += 1

    # LEO ΔV allocation check
    alloc_sum = "+".join(f"IFERROR({col(5+si*3)}{dv_alloc_row}*1,0)" for si in range(MAX_STAGES))
    ws.cell(r, 1, value="LEO ΔV Allocation Check (should = LEO Mission ΔV above)").font = Font(size=10, italic=True)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value=f"={alloc_sum}").font = Font(bold=True, size=10)
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["calc"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    ws.cell(r, 2).number_format = "#,##0"
    ws.cell(r, 3, value=f'=IF(ABS(B{r}-B{dv_budget_row})<1,"✓ OK","⚠ Adjust stage allocations to sum to LEO Mission ΔV")').font = Font(size=9, italic=True)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4 + MAX_STAGES * 3 - 1)
    r += 1

    # ── Section: Engine Mass (auto-looked-up) ─────────────────────────────────
    ws.cell(r, 1, value="Engine Mass per Engine").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    engine_mass_row = r
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        eng_cell = f"{col(c_val)}{stage_rows['Engine (from Engine DB)']}"
        formula = f'=IFERROR(VLOOKUP({eng_cell},\'Engine DB\'!$A:$K,8,0),"—")'
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "#,##0"
    r += 1

    # ── Section: Subsystem Mass Breakdown ────────────────────────────────────
    r += 1
    style_sec(ws, r, 1, "SUBSYSTEM MASS BREAKDOWN (auto-estimated; override any yellow cell)", span=2 + MAX_STAGES * 3)
    ws.row_dimensions[r].height = 18
    r += 1

    prop_row           = stage_rows["Propellant Combination"]
    base_prop_mass_row = stage_rows["Propellant Mass (nominal)"]  # used for subsystem fractions
    prop_mass_row      = stage_rows["Effective Propellant Mass"]   # used for all performance calcs

    subsys_rows = {}
    for sub in SUBSYSTEMS:
        sub_name, rp1, lh2, ch4, sol, hyp, desc = sub
        ws.cell(r, 1, value=sub_name).font = Font(size=10)
        ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=2)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
        ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
        ws.cell(r, 2).alignment = Alignment(horizontal="center")
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
        subsys_rows[sub_name] = r

        for si in range(MAX_STAGES):
            c_val = 5 + si * 3
            prop_cell  = f"{col(c_val)}{prop_row}"
            pmass_cell = f"{col(c_val)}{base_prop_mass_row}"
            formula = (
                f'=IF({pmass_cell}="","",{pmass_cell}/100*'
                f'IF({prop_cell}="LOX/RP-1",{rp1},'
                f'IF({prop_cell}="LOX/LH2",{lh2},'
                f'IF({prop_cell}="LOX/CH4",{ch4},'
                f'IF(OR({prop_cell}="Solid HTPB",{prop_cell}="Solid PBAN"),{sol},'
                f'{hyp})))))'
            )
            cell = style_input(ws, r, c_val, value=formula)
            cell.fill = PatternFill("solid", fgColor=C["input"])
            cell.number_format = "#,##0"
        r += 1

    # Engine system mass (engine mass × count × 1.15 for feed system hardware)
    ws.cell(r, 1, value="Engine System Total").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=2)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    engine_sys_row = r
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        n_eng_cell    = f"{col(c_val)}{stage_rows['Number of Engines']}"
        eng_mass_cell = f"{col(c_val)}{engine_mass_row}"
        formula = f'=IFERROR({n_eng_cell}*{eng_mass_cell}*1.15, "")'
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "#,##0"
    r += 1

    # Total estimated dry mass
    ws.cell(r, 1, value="TOTAL DRY MASS (Estimated)").font = Font(bold=True, size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["sec"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["sec"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    dry_mass_row = r
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        sub_refs = "+".join(f"{col(c_val)}{subsys_rows[s[0]]}" for s in SUBSYSTEMS)
        eng_ref  = f"{col(c_val)}{engine_sys_row}"
        formula  = f"=IFERROR({sub_refs}+{eng_ref}, \"\")"
        cell = style_calc(ws, r, c_val, value=formula)
        cell.font = Font(bold=True, size=10, italic=True)
        cell.number_format = "#,##0"
    r += 1

    # Dry mass override (enter known value to replace estimate)
    ws.cell(r, 1, value="Dry Mass Override (enter known value)").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    dry_override_row = r
    # Default dry mass overrides for New Glenn Block 2 (from vehicle DB / Blue Origin estimates)
    dry_override_defaults = vd["s_dry_ovr"]
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        style_input(ws, r, c_val, value=dry_override_defaults[si])
        ws.cell(r, c_val).number_format = "#,##0"
    r += 1

    # Effective dry mass: override if present, else estimate
    ws.cell(r, 1, value="Effective Dry Mass").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    eff_dry_row = r
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        ov = f"{col(c_val)}{dry_override_row}"
        es = f"{col(c_val)}{dry_mass_row}"
        formula = f'=IF({ov}<>"",{ov},{es})'
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "#,##0"
    r += 1

    # ── Fill deferred GLOW and TWR formulas (needed eff_dry_row) ─────────────
    # GLOW = sum of all stages (nominal prop + effective dry) — no payload, conservative TWR
    glow_parts = "+".join(
        f"IFERROR({col(5+si*3)}{nom_prop_row}+{col(5+si*3)}{eff_dry_row},0)"
        for si in range(MAX_STAGES)
    )
    glow_cell = ws.cell(glow_est_r, 2, value=f"={glow_parts}")
    glow_cell.font = Font(size=10, italic=True)
    glow_cell.fill = PatternFill("solid", fgColor=C["calc"])
    glow_cell.alignment = Alignment(horizontal="center")
    glow_cell.number_format = "#,##0"

    # TWR = Stage 1 SL thrust (kN→N) × n_engines / (GLOW × g0)
    s1_eng_cell  = f"{col(5)}{stage_rows['Engine (from Engine DB)']}"
    s1_neng_cell = f"{col(5)}{stage_rows['Number of Engines']}"
    sl_thrust_f  = f"IFERROR(VLOOKUP({s1_eng_cell},'Engine DB'!$A:$K,7,0),0)"
    twr_formula  = f"=IFERROR({s1_neng_cell}*{sl_thrust_f}*1000/(B{glow_est_r}*{G0}),\"\")"
    twr_cell = ws.cell(twr_est_r, 2, value=twr_formula)
    twr_cell.font = Font(bold=True, size=10, italic=True)
    twr_cell.fill = PatternFill("solid", fgColor=C["calc"])
    twr_cell.alignment = Alignment(horizontal="center")
    twr_cell.number_format = "0.00"

    # ── Fill deferred ΔV Allocation formulas (needed dv_budget_row / dv_gto_budget_row) ──
    for (row_r, c_val, frac_c, param) in getattr(ws, '_dv_alloc_deferred', []):
        budget = dv_budget_row if param == "LEO ΔV Allocation" else dv_gto_budget_row
        formula = f'=IFERROR({frac_c}*B{budget},"")'
        cell = ws.cell(row_r, c_val, value=formula)
        cell.fill = PatternFill("solid", fgColor=C["calc"])
        cell.font = Font(size=10, italic=True)
        cell.alignment = Alignment(horizontal="center")
        cell.number_format = "#,##0"

    # ── Fill deferred Recovery Propellant Reserved formulas (needed eff_dry_row) ──
    # Formula: eff_dry × (exp(ΔV_rec / (vac_isp × g0)) - 1)
    # Derivation: recovery burns consume prop from a stage of mass (dry+recovery_prop),
    # decelerating to dry mass. Rearranging rocket eq for recovery_prop:
    #   ΔV = vac_isp·g0·ln((dry+recovery_prop)/dry)  →  recovery_prop = dry·(e^(ΔV/isp/g0)-1)
    recovery_prop_row = stage_rows["Recovery Propellant Reserved"]
    for (row_r, c_val) in getattr(ws, '_recovery_prop_deferred', []):
        dv_rec_c  = f"{col(c_val)}{stage_rows['Recovery ΔV Reserve (m/s)']}"
        vac_isp_c = f"{col(c_val)}{stage_rows['Vac Isp (from DB)']}"
        eff_dry_c = f"{col(c_val)}{eff_dry_row}"
        formula = (f'=IFERROR(IF({dv_rec_c}=0,0,'
                   f'{eff_dry_c}*(EXP({dv_rec_c}/({vac_isp_c}*{G0}))-1)),0)')
        cell = ws.cell(row_r, c_val, value=formula)
        cell.fill = PatternFill("solid", fgColor=C["calc"])
        cell.font = Font(size=10, italic=True)
        cell.alignment = Alignment(horizontal="center")
        cell.number_format = "#,##0"

    # ── Section: Performance ─────────────────────────────────────────────────
    r += 1
    style_sec(ws, r, 1, "PERFORMANCE CALCULATIONS", span=2 + MAX_STAGES * 3)
    ws.row_dimensions[r].height = 18
    r += 1

    # ── Max Payload (primary output) ──────────────────────────────────────────
    # Derived analytically from the last active stage using Tsiolkovsky inverted:
    #   m_above = prop / (EXP(dv_alloc / (Isp * g0)) - 1) - dry_mass
    #   max_payload = m_above - adapter_mass
    # The last active stage is determined by Number of Stages input.
    # We compute for each stage and select the appropriate one.
    ws.cell(r, 1, value="MAX PAYLOAD TO ORBIT").font = Font(bold=True, size=12)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["pos"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["pos"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    max_payload_row = r
    adapter_cell = f"B{adapter_row}"

    # Build per-stage payload formula (reusability-aware):
    #   ascent_prop = effective_prop - recovery_reserved
    #   eff_dry_asc = eff_dry + recovery_reserved  (recovery prop is dead weight during ascent)
    #   m_payload = ascent_prop/(e^(ΔV/Isp·g₀)−1) − eff_dry_asc − adapter
    #
    # Payload is solved from the LAST active stage (nested IF on n_stages).
    # Recovery reserve on lower stages reduces their actual ΔV (shown in verification rows).
    # To model the reusability payload penalty: reduce the lower stage ΔV fraction so the
    # upper stage must cover more, which lowers payload through the last-stage formula.
    stage_payload_formulas = []
    for si in range(MAX_STAGES):
        c_val       = 5 + si * 3
        prop_c      = f"{col(c_val)}{prop_mass_row}"
        rec_c       = f"{col(c_val)}{recovery_prop_row}"
        isp_c       = f"{col(c_val)}{stage_rows['Effective Isp (used)']}"
        dv_c        = f"{col(c_val)}{dv_alloc_row}"
        dry_c       = f"{col(c_val)}{eff_dry_row}"
        ascent_prop = f"({prop_c}-{rec_c})"
        eff_dry_asc = f"({dry_c}+{rec_c})"
        f = (f"IFERROR({ascent_prop}/(EXP({dv_c}/({isp_c}*{G0}))-1)"
             f"-{eff_dry_asc}-{adapter_cell}, \"\")")
        stage_payload_formulas.append(f)

    # Select payload from last active stage via nested IF on n_stages
    n_stages_cell = f"B{mission_rows['Number of Stages']}"
    nested = stage_payload_formulas[0]
    for si in range(1, MAX_STAGES):
        nested = f"IF({n_stages_cell}>={si+1},{stage_payload_formulas[si]},{nested})"
    payload_formula = f"={nested}"

    cell = style_calc(ws, r, 2, value=payload_formula)
    cell.font = Font(bold=True, size=12, italic=True)
    cell.number_format = "#,##0"
    cell.fill = PatternFill("solid", fgColor=C["pos"])
    note_cell = ws.cell(r, 3, value="← Solved from last stage: m_payload = (prop−rec)/(e^(ΔV/Isp·g₀)−1) − (dry+rec) − adapter  |  Recovery penalty: lower stage's actual ΔV drops (see verification); reduce that stage's ΔV fraction to propagate penalty to payload")
    note_cell.font = Font(size=8, italic=True, color="444444")
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4 + MAX_STAGES * 3 - 1)
    r += 1

    # ── MAX PAYLOAD TO GTO (analytical — same method as LEO) ─────────────────
    r += 1
    style_sec(ws, r, 1, "GTO PERFORMANCE", span=2 + MAX_STAGES * 3)
    ws.row_dimensions[r].height = 18
    r += 1

    ws.cell(r, 1, value="MAX PAYLOAD TO GTO").font = Font(bold=True, size=12)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["pos"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["pos"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    max_payload_gto_row = r

    gto_dv_alloc_row = stage_rows["GTO ΔV Allocation"]
    gto_stage_payload_formulas = []
    for si in range(MAX_STAGES):
        c_val       = 5 + si * 3
        prop_c      = f"{col(c_val)}{prop_mass_row}"
        rec_c       = f"{col(c_val)}{recovery_prop_row}"
        isp_c       = f"{col(c_val)}{stage_rows['Effective Isp (used)']}"
        dv_c        = f"{col(c_val)}{gto_dv_alloc_row}"
        dry_c       = f"{col(c_val)}{eff_dry_row}"
        ascent_prop = f"({prop_c}-{rec_c})"
        eff_dry_asc = f"({dry_c}+{rec_c})"
        f = (f"IFERROR({ascent_prop}/(EXP({dv_c}/({isp_c}*{G0}))-1)"
             f"-{eff_dry_asc}-{adapter_cell}, \"\")")
        gto_stage_payload_formulas.append(f)

    gto_nested = gto_stage_payload_formulas[0]
    for si in range(1, MAX_STAGES):
        gto_nested = f"IF({n_stages_cell}>={si+1},{gto_stage_payload_formulas[si]},{gto_nested})"
    cell = style_calc(ws, r, 2, value=f"={gto_nested}")
    cell.font = Font(bold=True, size=12, italic=True)
    cell.number_format = "#,##0"
    cell.fill = PatternFill("solid", fgColor=C["pos"])
    note_cell = ws.cell(r, 3, value="← MIN across all stage constraints using GTO ΔV allocations: capacity_i = (prop−rec)/(e^(ΔV/Isp·g₀)−1) − (dry+rec) − gross_above − adapter")
    note_cell.font = Font(size=8, italic=True, color="444444")
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4 + MAX_STAGES * 3 - 1)
    r += 1

    # ── Stage-by-stage verification (Tsiolkovsky forward check) ──────────────
    # Now that we have max_payload, compute actual ΔV per stage as a check
    # gross_mass(i) = prop(i) + dry(i)
    ws.cell(r, 1, value="Gross Stage Mass (Prop + Dry)").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    gross_row = r
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        formula = (f"=IFERROR({col(c_val)}{prop_mass_row}"
                   f"+{col(c_val)}{eff_dry_row}, \"\")")
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "#,##0"
    r += 1

    # Burnout mass (m_f): dry + recovery_reserved (still onboard at MECO) + stages above + payload
    # For expendable stages, recovery_reserved = 0, so formula is unchanged.
    ws.cell(r, 1, value="Burnout Mass m_f").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    burnout_row = r
    payload_calc_cell = f"B{max_payload_row}"
    for si in range(MAX_STAGES):
        c_val    = 5 + si * 3
        dry_cell = f"{col(c_val)}{eff_dry_row}"
        rec_cell = f"{col(c_val)}{recovery_prop_row}"
        above_gross = "+".join(
            f"IFERROR({col(5 + j * 3)}{gross_row}*1,0)"
            for j in range(si + 1, MAX_STAGES)
        )
        if above_gross:
            formula = (f"=IFERROR({dry_cell}+{rec_cell}+{above_gross}"
                       f"+{payload_calc_cell}+{adapter_cell}, \"\")")
        else:
            formula = (f"=IFERROR({dry_cell}+{rec_cell}"
                       f"+{payload_calc_cell}+{adapter_cell}, \"\")")
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "#,##0"
    r += 1

    # Initial mass m_0: burnout + ascent propellant
    # burnout already includes recovery_reserved (still onboard at MECO), so we add only
    # the ascent portion (effective_prop - recovery_reserved) to avoid double-counting.
    ws.cell(r, 1, value="Initial Mass m_0 (at ignition)").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="kg").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    m0_row = r
    for si in range(MAX_STAGES):
        c_val    = 5 + si * 3
        rec_cell = f"{col(c_val)}{recovery_prop_row}"
        formula  = (f"=IFERROR({col(c_val)}{burnout_row}"
                    f"+{col(c_val)}{prop_mass_row}-{rec_cell}, \"\")")
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "#,##0"
    r += 1

    # Mass ratio
    ws.cell(r, 1, value="Mass Ratio (m₀ / m_f)").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="—").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    mr_row = r
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        formula = (f"=IFERROR({col(c_val)}{m0_row}"
                   f"/{col(c_val)}{burnout_row}, \"\")")
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "0.000"
    r += 1

    # Actual Delta-V (verification — should match ΔV Allocation if vehicle is sized correctly)
    ws.cell(r, 1, value="Actual Stage ΔV (verification)").font = Font(bold=True, size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["sec"])
    ws.cell(r, 2, value="m/s").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["sec"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    dv_row = r
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        isp_c = f"{col(c_val)}{stage_rows['Effective Isp (used)']}"
        mr_c  = f"{col(c_val)}{mr_row}"
        formula = f"=IFERROR({isp_c}*{G0}*LN({mr_c}), \"\")"
        cell = style_calc(ws, r, c_val, value=formula)
        cell.font = Font(bold=True, size=10, italic=True)
        cell.number_format = "#,##0"
    r += 1

    # TWR at stage ignition
    ws.cell(r, 1, value="Thrust-to-Weight Ratio (at ignition)").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="—").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    twr_row = r
    for si in range(MAX_STAGES):
        c_val      = 5 + si * 3
        eng_cell   = f"{col(c_val)}{stage_rows['Engine (from Engine DB)']}"
        n_eng_cell = f"{col(c_val)}{stage_rows['Number of Engines']}"
        m0_cell    = f"{col(c_val)}{m0_row}"
        # Stage 1: use sea-level thrust; upper stages: vacuum thrust
        if si == 0:
            thrust_f = f"IFERROR(VLOOKUP({eng_cell},'Engine DB'!$A:$K,7,0),0)"
        else:
            thrust_f = f"IFERROR(VLOOKUP({eng_cell},'Engine DB'!$A:$K,6,0),0)"
        formula = f"=IFERROR({n_eng_cell}*{thrust_f}*1000/({m0_cell}*{G0}), \"\")"
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "0.00"
    r += 1

    # Propellant mass fraction
    ws.cell(r, 1, value="Propellant Mass Fraction").font = Font(size=10)
    ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2, value="—").font = Font(size=9, color="777777")
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["ltgray"])
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    for si in range(MAX_STAGES):
        c_val = 5 + si * 3
        formula = (f"=IFERROR({col(c_val)}{prop_mass_row}"
                   f"/{col(c_val)}{m0_row}, \"\")")
        cell = style_calc(ws, r, c_val, value=formula)
        cell.number_format = "0.000"
    r += 1

    # ── Summary ───────────────────────────────────────────────────────────────
    r += 1
    style_sec(ws, r, 1, "VEHICLE SUMMARY", span=2 + MAX_STAGES * 3)
    ws.row_dimensions[r].height = 18
    r += 1

    dv_actual_sum = "+".join(
        f"IFERROR({col(5 + si * 3)}{dv_row}*1,0)" for si in range(MAX_STAGES)
    )
    total_prop = "+".join(
        f"IFERROR({col(5+si*3)}{prop_mass_row},0)" for si in range(MAX_STAGES)
    )

    summary_rows = [
        ("Max Payload to LEO",          "kg",   f"=B{max_payload_row}"),
        ("Max Payload to GTO",          "kg",   f"=B{max_payload_gto_row}"),
        ("LEO Payload Fraction (GLOW)", "%",    f"=IFERROR(B{max_payload_row}/{col(5)}{m0_row}*100, \"\")"),
        ("Total Vehicle ΔV (actual)",   "m/s",  f"={dv_actual_sum}"),
        ("LEO ΔV Margin vs Mission",    "m/s",  f"={dv_actual_sum}-B{dv_budget_row}"),
        ("Gross Liftoff Mass (GLOW)",   "kg",   f"=IFERROR({col(5)}{m0_row}, \"\")"),
        ("Total Propellant Mass",       "kg",   f"={total_prop}"),
    ]

    for label, unit, formula in summary_rows:
        is_primary = label in ("Max Payload to LEO", "Max Payload to GTO")
        ws.cell(r, 1, value=label).font = Font(bold=True, size=11 if is_primary else 10)
        ws.cell(r, 1).alignment = Alignment(horizontal="left", indent=1)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=C["pos"] if is_primary else C["sec"])
        ws.cell(r, 2, value=formula).font = Font(bold=True, size=11 if is_primary else 10)
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=C["pos"] if is_primary else C["calc"])
        ws.cell(r, 2).alignment = Alignment(horizontal="center")
        ws.cell(r, 2).number_format = "#,##0" if unit in ("m/s", "kg") else "0.00"
        ws.cell(r, 3, value=unit).font = Font(size=9, color="777777")
        ws.cell(r, 3).alignment = Alignment(horizontal="left")
        apply_border_range(ws, r, r, 1, 3)
        r += 1

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# SHEET: Reference Comparison
# ─────────────────────────────────────────────────────────────────────────────
def build_comparison(wb):
    ws = wb.create_sheet("Comparison")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    style_hdr(ws, 1, 1, "REFERENCE VEHICLE QUICK-COMPARE — Key Performance Metrics", span=9)
    headers = ["Vehicle", "Stages", "GLOW (kg)", "Payload LEO (kg)",
               "Payload Fraction (%)", "Payload GTO (kg)", "Propulsion", "Height (m)", "Status"]
    widths   = [26, 8, 16, 16, 16, 16, 22, 10, 16]
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        style_hdr(ws, 2, ci, h, dark=False)
        ws.column_dimensions[col(ci)].width = w
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 30

    for ri, v in enumerate(VEHICLES, 3):
        pf = round(v["payload_leo_kg"] / v["glow_kg"] * 100, 2) if v["glow_kg"] and v["payload_leo_kg"] else None
        row_vals = [
            v["name"], len(v["stages"]), v["glow_kg"], v["payload_leo_kg"],
            pf, v.get("payload_gto_kg", "—"), v.get("propulsion", "—"), v["height_m"], v["status"]
        ]
        for ci, val in enumerate(row_vals, 1):
            cell = ws.cell(ri, ci, value=val if val is not None else "—")
            cell.font = Font(bold=(ci == 1), size=10)
            cell.alignment = Alignment(horizontal="left" if ci in (1, 7, 9) else "center",
                                       vertical="center", indent=1)
            cell.fill = PatternFill("solid", fgColor=C["ltgray"] if ri % 2 == 0 else C["white"])
            if ci in (3, 4, 6):
                cell.number_format = "#,##0"
            if ci == 5 and val is not None:
                cell.number_format = "0.00"
        apply_border_range(ws, ri, ri, 1, len(headers))
        ws.row_dimensions[ri].height = 15

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# SHEET: Instructions
# ─────────────────────────────────────────────────────────────────────────────
def build_readme(wb):
    ws = wb.create_sheet("README", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 100

    style_hdr(ws, 1, 1, "LAUNCH VEHICLE DESIGN TOOL — User Guide")
    ws.row_dimensions[1].height = 28

    lines = [
        ("DISCLAIMER", True),
        ("All analysis in this tool is conducted using only publicly available information or assumptions", False),
        ("based on industry standards and engineering analysis. This includes:", False),
        ("  •  Published payload figures and performance claims from vehicle operators", False),
        ("  •  Announced engine performance data (Isp, thrust) from manufacturer press releases", False),
        ("  •  Official vehicle dimensions from public sources (Wikipedia, manufacturer websites)", False),
        ("  •  Industry-standard models for gravity loss, drag loss, and structural mass fractions", False),
        ("  •  Open-source reference materials and engineering first principles", False),
        ("", False),
        ("No proprietary, export-controlled, or non-public data has been used. Parameters for vehicles", False),
        ("under development (New Glenn 9x4, Vulcan VC6, etc.) are engineering estimates derived from", False),
        ("public announcements and documented assumptions. Where actual data is unavailable, assumptions", False),
        ("are explicitly flagged. Results are for educational and analytical purposes only and should not", False),
        ("be taken as authoritative performance specifications.", False),
        ("", False),
        ("HOW TO USE", True),
        ("1.  Go to 'Vehicle Design' tab — this is your main working sheet.", False),
        ("2.  Fill in Mission Requirements (yellow cells): payload mass, target orbit, delta-v budget.", False),
        ("3.  For each stage, enter the number of stages (1-3) and configure each stage:", False),
        ("      • Engine: Type the exact engine name from the 'Engine DB' tab. Isp will auto-populate via VLOOKUP.", False),
        ("      • Number of Engines: How many of that engine on this stage.", False),
        ("      • Propellant Combination: Must match exactly (e.g. LOX/RP-1, LOX/LH2, LOX/CH4, Solid HTPB, NTO/MMH).", False),
        ("      • Propellant Mass: Total usable propellant mass for this stage in kg.", False),
        ("      • Vac Isp Override: Leave blank to use the Engine DB value. Enter a number to override.", False),
        ("", False),
        ("4.  Subsystem masses auto-calculate from parametric models in the 'Propellants' tab.", False),
        ("      • Yellow cells in the subsystem section can be manually overridden with known values.", False),
        ("      • Use 'Dry Mass Override' if you have a known dry mass (e.g. from a reference vehicle).", False),
        ("", False),
        ("5.  For reusable stages (e.g. New Glenn S1, Falcon 9 S1), enter a Recovery ΔV Reserve:", False),
        ("      • This is the total ΔV budget for recovery burns (boost-back + entry + landing).", False),
        ("      • Typical values: ~750 m/s for RTLS (return-to-launch-site), ~500 m/s for ASDS (drone ship).", False),
        ("      • The tool computes Recovery Propellant Reserved = dry_mass × (e^(ΔV/Isp/g₀) − 1).", False),
        ("      • This propellant is treated as dead weight during ascent (still onboard at MECO).", False),
        ("      • Set to 0 for expendable stages — no penalty applied.", False),
        ("", False),
        ("6.  Performance section calculates automatically:", False),
        ("      • Delta-V per stage via Tsiolkovsky equation: ΔV = Isp × g₀ × ln(m₀/m_f)", False),
        ("      • Burnout mass includes recovery-reserved propellant (still onboard at MECO) for reusable stages.", False),
        ("      • TWR uses sea-level thrust for Stage 1, vacuum thrust for upper stages.", False),
        ("", False),
        ("TIPS & RULES OF THUMB", True),
        ("• LEO missions typically need 9,000–9,500 m/s total ΔV (includes ~1,500 m/s gravity+drag losses).", False),
        ("• GTO missions need ~11,500–12,500 m/s depending on inclination.", False),
        ("• First stage TWR should be 1.2–1.5 at liftoff; too low = slow ascent; too high = unnecessary mass.", False),
        ("• Payload fraction for real vehicles: 1.5–4% to LEO is typical.", False),
        ("• Structural mass fraction (dry/propellant) is typically 5–15%; lower is better engineering.", False),
        ("• LOX/LH2 gives the highest Isp (~450s) but low density means large, heavy tanks.", False),
        ("• LOX/RP-1 gives good density and SL performance; workhorse of most first stages.", False),
        ("• LOX/CH4 is the emerging choice for reusability; good Isp + density balance.", False),
        ("• Reusable first stages carry a payload penalty: ~30–40% less to LEO vs expendable (Falcon 9 example).", False),
        ("• Recovery ΔV budget: RTLS ~750 m/s (boost-back ~300 + entry ~200 + landing ~250);", False),
        ("  ASDS (drone ship) ~450–500 m/s (no boost-back; entry ~200 + landing ~250).", False),
        ("", False),
        ("REFERENCE TABS", True),
        ("• Engine DB:   Real engine performance data for 35+ engines. Add your own rows at the bottom.", False),
        ("• Vehicle DB:  Real vehicle mass breakdowns for 11 vehicles. Use for design grounding.", False),
        ("• Propellants: Propellant properties + subsystem mass fraction model documentation.", False),
        ("• Comparison:  Quick performance comparison table across all reference vehicles.", False),
        ("", False),
        ("LIMITATIONS", True),
        ("• This tool uses simplified mass models. Real designs require detailed MDO and stress analysis.", False),
        ("• Delta-V calculations assume impulsive burns (no trajectory shaping or staging timing).", False),
        ("• Subsystem fractions are statistical — early concept accuracy is ±20–30%.", False),
        ("• Solid stages show N/A for SL Isp in engine DB — use Vac Isp as a conservative estimate.", False),
        ("• Parallel staging (e.g. strap-on boosters firing simultaneously with a core) is not modeled.", False),
        ("  Treat parallel-burn vehicles (Soyuz, Ariane 5) as series stages for delta-v calculation.", False),
        ("", False),
        ("SIMPLIFYING ASSUMPTIONS & CALIBRATION PARAMETERS", True),
        ("", False),
        ("This model combines physics-based equations with several parameters that cannot be derived from", False),
        ("first principles alone. Some defaults are reasonable engineering approximations. Others were chosen", False),
        ("specifically to bring outputs into agreement with Falcon 9 Block 5 published performance figures", False),
        ("and should be treated as guesses until calibrated against a vehicle you trust.", False),
        ("", False),
        ("The table below summarises every tunable parameter, its default value, and an honest assessment", False),
        ("of how much confidence to place in that default.", False),
        ("", False),
        ("  Parameter                        Default        Confidence   Notes", False),
        ("  ─────────────────────────────────────────────────────────────────────────────────────────────", False),
        ("  Gravity Loss Coefficient         1,750          LOW          Empirical. Fitted to approximate typical", False),
        ("                                                               vehicles. Needs recalibration per vehicle.", False),
        ("", False),
        ("  Trajectory Penalty               500 m/s        GUESS        Reverse-engineered from F9 ASDS LEO", False),
        ("                                                               payload figure. No independent physical", False),
        ("                                                               basis. Set to 0 for expendable vehicles.", False),
        ("                                                               Do not apply to other vehicles without", False),
        ("                                                               recalibrating against known performance.", False),
        ("", False),
        ("  Stage ΔV Fractions (LEO)         0.40 / 0.60   GUESS        Reverse-engineered to hit a target F9", False),
        ("  Stage ΔV Fractions (GTO)         0.33 / 0.67   GUESS        payload number. Real optimal splits depend", False),
        ("                                                               on mass ratios and Isp. Re-derive per", False),
        ("                                                               vehicle by iterating until ΔV check = 0.", False),
        ("", False),
        ("  Recovery ΔV Reserve              500 m/s        MODERATE     Plausible range for ASDS. Varies with S1", False),
        ("                                                               mass and landing site distance. Not", False),
        ("                                                               verified against flight telemetry.", False),
        ("", False),
        ("  GTO Orbital ΔV                   10,400 m/s     MODERATE     Varies ±300–600 m/s with launch site,", False),
        ("                                                               inclination, and target apogee. Set per", False),
        ("                                                               mission rather than using the default.", False),
        ("", False),
        ("  Atm Fraction (S1)                0.6            MODERATE     Common rule of thumb. Varies with ascent", False),
        ("                                                               profile and nozzle design. Can be off by", False),
        ("                                                               5–10 s effective Isp.", False),
        ("", False),
        ("  Drag Loss                        120 m/s        MODERATE     Reasonable for typical trajectory.", False),
        ("                                                               Varies with vehicle diameter and max-q", False),
        ("                                                               throttling strategy.", False),
        ("", False),
        ("  Subsystem mass fractions         various        MODERATE     Statistical fits to a small number of", False),
        ("   (Propellants tab)                                           real vehicles. ±20–30% accuracy claimed;", False),
        ("                                                               individual subsystems may be worse.", False),
        ("", False),
        ("── STRUCTURAL SIMPLIFICATIONS (not tunable) ────────────────────────────────────────────────────", True),
        ("", False),
        ("These are inherent to the model architecture, not user inputs:", False),
        ("", False),
        ("• Isp is modeled as a fixed linear blend between SL and vacuum — not an altitude-integrated curve.", False),
        ("• Recovery burns are modeled as a single equivalent burn from dry mass, not multiple discrete burns.", False),
        ("• Fairing jettison is not modeled as a staging event — add fairing mass to Payload Adapter as needed.", False),
        ("• Isp does not vary with throttle level.", False),
        ("• GLOW used for TWR/gravity loss calculation excludes payload mass (slightly optimistic TWR).", False),
        ("• Parallel staging (strap-on boosters burning simultaneously with core) is not supported.", False),
    ]

    for ri, (text, bold) in enumerate(lines, 3):
        cell = ws.cell(ri, 1, value=text)
        cell.font = Font(bold=bold, size=10 if not bold else 11,
                         color="1F3864" if bold else "1F1F1F")
        if bold:
            cell.fill = PatternFill("solid", fgColor=C["sec"])
        ws.row_dimensions[ri].height = 16

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# SHEET: Payload Direct-Injection Orbital Reach Analysis
# ─────────────────────────────────────────────────────────────────────────────
def build_payload_analysis(wb, vehicle_cfg, payload_cfg):
    """
    Generic direct-injection orbital reach analysis for any vehicle + payload.

    The launch vehicle injects the payload into a maximum-apogee elliptical
    orbit at perigee.  No circularization burn is assumed — the spacecraft
    performs its own orbit-raising.  All stage ΔV minus trajectory losses goes
    directly into perigee injection velocity.

    Supports "standard" (series) and "parallel_core" (strap-on / SRB) stages
    via compute_dv_budget().  See vehicle analysis config dicts for examples.
    """
    MU_E = 3.986004418e14
    R_E  = 6_371_000

    budget   = compute_dv_budget(vehicle_cfg, payload_cfg)
    v_inject = budget["v_inject"]
    vname    = vehicle_cfg["vehicle_name"]
    pname    = payload_cfg["name"]
    pmass    = payload_cfg["mass_kg"]

    # ── Perigee sweep: vis-viva apogee for 185–500 km perigee altitudes ───────
    sweep_rows = []
    for h_p in range(185, 525, 25):
        r_p    = R_E + h_p * 1000
        v_circ = math.sqrt(MU_E / r_p)
        v_esc  = math.sqrt(2 * MU_E / r_p)
        denom  = 2 * MU_E / r_p - v_inject ** 2
        if denom <= 0:
            sweep_rows.append((h_p, v_circ, v_esc, None, None, None))
        else:
            a      = MU_E / denom
            r_apo  = 2 * a - r_p
            h_apo  = (r_apo - R_E) / 1000
            period = 2 * math.pi * math.sqrt(a ** 3 / MU_E) / 3600
            sweep_rows.append((h_p, v_circ, v_esc, h_apo, period, a / 1000))

    # ── Build worksheet ───────────────────────────────────────────────────────
    ws = wb.create_sheet(f"{pname} — {vname}"[:31])   # Excel 31-char sheet name limit
    ws.sheet_view.showGridLines = False

    col_widths = [38, 14, 14, 14, 14, 14, 14]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 1
    style_hdr(ws, r, 1,
              f"{pname.upper()}  DIRECT INJECTION ANALYSIS  —  {vname}",
              span=7)
    ws.row_dimensions[r].height = 26
    r += 1

    note = ws.cell(r, 1,
        f"Payload: {pname}, {pmass:,} kg  |  "
        "No circularization — spacecraft performs own orbit raising  |  "
        "All stage ΔV (minus losses) goes into maximising apogee at perigee injection")
    note.font = Font(size=9, italic=True, color="555555")
    ws.merge_cells(f"A{r}:G{r}")
    ws.row_dimensions[r].height = 14
    r += 2

    # ── Helper: write one label/unit/value budget row ─────────────────────────
    def brow(label, unit, value, bold=False):
        nonlocal r
        fg = C["pos"] if bold else C["ltgray"]
        c1 = ws.cell(r, 1, value=label)
        c1.font      = Font(bold=bold, size=10)
        c1.fill      = PatternFill("solid", fgColor=fg)
        c1.alignment = Alignment(horizontal="left", indent=1)
        c2 = ws.cell(r, 2, value=unit)
        c2.font      = Font(size=9, color="777777")
        c2.fill      = PatternFill("solid", fgColor=fg)
        c2.alignment = Alignment(horizontal="center")
        c3 = ws.cell(r, 3, value=round(value, 1) if isinstance(value, float) else value)
        c3.font      = Font(bold=bold, size=10, italic=True)
        c3.fill      = PatternFill("solid", fgColor=C["pos"] if bold else C["calc"])
        c3.alignment = Alignment(horizontal="center")
        c3.number_format = "#,##0" if unit in ("kg", "m/s") else "0.00"
        r += 1

    # ── Section: Stage ΔV Budget ──────────────────────────────────────────────
    style_sec(ws, r, 1,
              f"STAGE ΔV BUDGET  —  {pmass:,} kg PAYLOAD", span=7)
    ws.row_dimensions[r].height = 18
    r += 1

    for si, stage in enumerate(budget["stages"], 1):
        sn = stage["name"]

        if stage["type"] == "standard":
            sc = vehicle_cfg["stages"][si - 1]
            if sc.get("subcooling_pct", 0):
                brow(f"{sn}  Nominal Propellant",                      "kg",  stage["prop_mass_kg"])
                brow(f"{sn}  Eff. Propellant  ({sc['subcooling_pct']}% subcooling)", "kg", stage["eff_prop"])
            else:
                brow(f"{sn}  Propellant",                              "kg",  stage["eff_prop"])
            if stage["rec_prop"]:
                brow(f"{sn}  Recovery Propellant Reserved  ({stage['rec_dv_ms']} m/s ΔV)", "kg", stage["rec_prop"])
            brow(f"{sn}  Effective Dry Mass during ascent",            "kg",  stage["eff_dry_asc"])
            brow(f"{sn}  Effective Isp  ({stage['engine']['name']})",  "s",   stage["isp_eff"])
            brow(f"{sn}  Initial Mass m₀",                             "kg",  stage["m0"])
            brow(f"{sn}  Burnout Mass m_f",                            "kg",  stage["mf"])
            brow(f"{sn}  ΔV  (Tsiolkovsky)",                           "m/s", stage["dv"],  bold=True)

        elif stage["type"] == "parallel_core":
            brow(f"{sn}  Core Propellant",                             "kg",  stage["eff_prop"])
            if stage["rec_prop"]:
                brow(f"{sn}  Core Recovery Propellant Reserved",       "kg",  stage["rec_prop"])
            brow(f"{sn}  Core Effective Dry Mass",                     "kg",  stage["eff_dry_asc"])
            brow(f"{sn}  Core Effective Isp  ({stage['engine']['name']})", "s", stage["isp_eff"])
            brow(f"{sn}  Booster Propellant  ({stage['booster_engine']['name']} ×{stage['booster_n_eng']})", "kg", stage["booster_prop"])
            brow(f"{sn}  Booster Dry Mass",                            "kg",  stage["booster_dry"])
            brow(f"{sn}  Booster Effective Isp",                       "s",   stage["booster_isp_eff"])
            brow(f"{sn}  Blended Isp  (Phase 1, parallel burn)",       "s",   stage["isp_parallel"])
            brow(f"{sn}  Core Prop Burned in Phase 1",                 "kg",  stage["prop_core_p1"])
            brow(f"{sn}  Liftoff Mass m₀",                             "kg",  stage["m0"])
            brow(f"{sn}  Phase 1 ΔV  (boosters + core)",               "m/s", stage["dv_phase1"])
            brow(f"{sn}  Phase 2 ΔV  (core only)",                     "m/s", stage["dv_phase2"])
            brow(f"{sn}  Total ΔV",                                    "m/s", stage["dv"],  bold=True)

        r += 1  # blank separator between stages

    brow("Total Stage ΔV  (all stages)", "m/s", budget["total_dv"], bold=True)
    r += 1

    # ── Section: Losses & Injection Velocity ─────────────────────────────────
    style_sec(ws, r, 1, "TRAJECTORY LOSSES  &  INJECTION VELOCITY", span=7)
    ws.row_dimensions[r].height = 18
    r += 1

    grav_label = (f"Gravity Loss  ({vehicle_cfg['grav_loss_coeff']:,} / TWR"
                  + (f" + {vehicle_cfg['traj_penalty_ms']} m/s penalty)" if vehicle_cfg.get("traj_penalty_ms") else ")"))
    brow("Liftoff TWR  (stage masses only, no payload)",  "—",   budget["twr"])
    brow(grav_label,                                      "m/s", budget["grav_loss"])
    brow("Drag Loss",                                     "m/s", budget["drag_loss"])
    brow("Total Losses",                                  "m/s", budget["losses"])
    brow("INJECTION VELOCITY AT PERIGEE  (total ΔV − losses)", "m/s", v_inject, bold=True)
    r += 1

    r_ref     = R_E + 300_000
    v_esc_ref = math.sqrt(2 * MU_E / r_ref)
    ratio     = v_inject / v_esc_ref
    note2 = ws.cell(r, 1,
        f"Injection velocity {v_inject:,.0f} m/s = {ratio*100:.1f}% of escape velocity "
        f"at 300 km ({v_esc_ref:,.0f} m/s)  —  "
        + ("well-bound elliptical orbit" if ratio < 1 else "EXCEEDS ESCAPE VELOCITY — hyperbolic trajectory"))
    note2.font      = Font(size=9, italic=True, color="444444")
    note2.fill      = PatternFill("solid", fgColor=C["ltgray"])
    note2.alignment = Alignment(horizontal="left", indent=1, wrap_text=True)
    ws.merge_cells(f"A{r}:G{r}")
    ws.row_dimensions[r].height = 16
    r += 2

    # ── Section: Perigee sweep table ─────────────────────────────────────────
    style_hdr(ws, r, 1,
              "APOGEE ALTITUDE vs PERIGEE ALTITUDE  —  Direct Injection, No Circularization",
              span=7)
    ws.row_dimensions[r].height = 20
    r += 1

    for ci, h in enumerate(["Perigee Alt (km)", "v_circ at perigee (m/s)",
                             "v_inject (m/s)", "v_escape at perigee (m/s)",
                             "Apogee Alt (km)", "Semi-major axis (km)", "Orbital Period (hr)"], 1):
        style_hdr(ws, r, ci, h, dark=False)
    ws.row_dimensions[r].height = 28
    r += 1

    for h_p, v_circ, v_esc, h_apo, period, a_km in sweep_rows:
        fg = C["ltgray"] if (h_p // 50) % 2 else C["white"]
        vals = [h_p, round(v_circ), round(v_inject), round(v_esc),
                round(h_apo) if h_apo is not None else "escape",
                round(a_km)  if a_km  is not None else "—",
                round(period, 2) if period is not None else "—"]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(r, ci, value=v)
            cell.fill      = PatternFill("solid", fgColor=fg)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font      = Font(size=10)
            if ci in (1, 2, 3, 4, 5, 6) and isinstance(v, (int, float)):
                cell.number_format = "#,##0"
        apply_border_range(ws, r, r, 1, 7)
        ws.row_dimensions[r].height = 15
        r += 1

    # ── Section: Result summary ───────────────────────────────────────────────
    r += 1
    style_sec(ws, r, 1, "RESULT", span=7)
    ws.row_dimensions[r].height = 18
    r += 1

    r_rep     = R_E + 300_000
    denom_rep = 2 * MU_E / r_rep - v_inject ** 2
    if denom_rep > 0:
        a_rep     = MU_E / denom_rep
        h_apo_rep = (2 * a_rep - r_rep - R_E) / 1000
        T_rep     = 2 * math.pi * math.sqrt(a_rep ** 3 / MU_E) / 3600
        result_headline = (f"At 300 km perigee:  apogee ≈ {h_apo_rep:,.0f} km  "
                           f" |  orbital period ≈ {T_rep:.2f} hr")
    else:
        result_headline = "Injection velocity exceeds escape velocity — hyperbolic trajectory (solar system escape)"

    def fmt_apo(v):
        return f"{v:,} km" if isinstance(v, int) else str(v)

    low_h_apo  = next((round(row[3]) for row in sweep_rows          if row[3] is not None), "escape")
    high_h_apo = next((round(row[3]) for row in reversed(sweep_rows) if row[3] is not None), "escape")

    for text, bold, fg in [
        (result_headline, True, C["pos"]),
        (f"  Apogee range across perigee sweep: {fmt_apo(low_h_apo)} (185 km perigee) "
         f"→ {fmt_apo(high_h_apo)} (500 km perigee)", False, C["calc"]),
        (f"  Spacecraft performs own apogee-raising burns from this orbit to reach TLI.  "
         "Moon mean distance ≈ 378,000 km — beyond direct injection capability.", False, C["ltgray"]),
        (f"  Model: gravity loss coeff {vehicle_cfg['grav_loss_coeff']:,} / TWR"
         + (f" + {vehicle_cfg['traj_penalty_ms']} m/s trajectory penalty" if vehicle_cfg.get('traj_penalty_ms') else "")
         + f".  Drag loss {vehicle_cfg.get('drag_loss_ms', 120)} m/s.", False, C["ltgray"]),
    ]:
        cell = ws.cell(r, 1, value=text)
        cell.font      = Font(bold=bold, size=11 if bold else 9, italic=(not bold))
        cell.fill      = PatternFill("solid", fgColor=fg)
        cell.alignment = Alignment(horizontal="left", indent=1, wrap_text=True)
        ws.merge_cells(f"A{r}:G{r}")
        ws.row_dimensions[r].height = 28 if bold else 22
        r += 1

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    # ── Main LV design workbook ───────────────────────────────────────────────
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    tab_colors = {
        "README":         "1F3864",
        "Vehicle Design": "2E7D32",
        "Engine DB":      "BF360C",
        "Vehicle DB":     "4527A0",
        "Propellants":    "00695C",
        "Comparison":     "1565C0",
    }

    # ── One design workbook per vehicle ───────────────────────────────────────
    for vd in [NG_7X2_DESIGN, NG_9X4_DESIGN, NG_9X4_NORTHSTAR_DESIGN]:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        build_readme(wb)
        build_design(wb, vd)
        build_engine_db(wb)
        build_vehicle_db(wb)
        build_propellants(wb)
        build_comparison(wb)
        for ws in wb.worksheets:
            if ws.title in tab_colors:
                ws.sheet_properties.tabColor = tab_colors[ws.title]
        raw_name  = wb["Vehicle Design"].cell(6, 2).value or "launch_vehicle"
        safe_name = "".join(c if c.isalnum() or c in " _-" else "" for c in str(raw_name)).strip().replace(" ", "_")
        fname = f"lv_design_{safe_name}.xlsx"
        wb.save(fname)
        print(f"✓ Saved: {fname}  (Engines: {len(ENGINES)}, Vehicles: {len(VEHICLES)})")

    # ── Payload analysis workbooks — one per vehicle config ───────────────────
    for vcfg in [ANALYSIS_NG_BLOCK2_7X2, ANALYSIS_VULCAN_VC6,
                 ANALYSIS_SLS_BLOCK1,     ANALYSIS_NG_9X4,
                 ANALYSIS_NG_9X4_NORTHSTAR]:
        wb_a = openpyxl.Workbook()
        wb_a.remove(wb_a.active)
        build_payload_analysis(wb_a, vcfg, PAYLOAD_BLUE_MOON_MK1)
        wb_a.worksheets[0].sheet_properties.tabColor = "B8860B"
        wb_a.save(vcfg["output_file"])
        budget = compute_dv_budget(vcfg, PAYLOAD_BLUE_MOON_MK1)
        print(f"✓ Saved: {vcfg['output_file']}"
              f"  (total ΔV {budget['total_dv']:,.0f} m/s,"
              f"  v_inject {budget['v_inject']:,.0f} m/s)")


if __name__ == "__main__":
    main()
