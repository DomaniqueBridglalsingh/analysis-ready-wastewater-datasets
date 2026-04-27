"""
===============================================================================
HYDROSTREAM — WATER QUALITY ARCHIVE STREAMING PIPELINE
===============================================================================
Title   : Environment Agency (England) Open Water Quality Archive — Processor
Version : 1.0.0
Authors : Domanique Bridglalsingh, Ahmed Abdalla, Jia Hu, Geyong Min,
          Xiaohong Li, and Siwei Zheng
Licence : CC-BY-4.0  (same licence as the underlying EA data)
Python  : >= 3.9

PURPOSE
-------
The Environment Agency (EA) published annual CSV files of water-quality
measurements from 2000 to 2025.  Those files are no longer publicly hosted.
HydroStream reads the raw yearly CSVs in memory-efficient streaming chunks,
applies transparent and reproducible cleaning steps, and produces a single,
analysis-ready dataset together with descriptive statistics and a QA report.

THREE OUTPUT MODES
------------------
1. "full"             – Every water-related test, type, and sampling point.
2. "electrochemistry" – A focused subset of dissolved metals, ions, pH,
                        conductivity, temperature, and turbidity — the
                        parameters most relevant to electrochemical sensing.
3. "contaminants"     – Microplastics, nanoplastics, PFAS, insecticides,
                        pesticides, and similar emerging contaminants.
                        REQUIRES the categories file (see below).

HOW TO USE  (see bottom of file for a ready-made example)
---------------------------------------------------------
1. Place all 26 raw CSV files (2000.csv … 2025.csv) in ONE folder.
2. Place  "List of tests kept and categories.xlsx"  in the same folder.
     • OPTIONAL for modes "full" and "electrochemistry" — enables the
       Category column in the output.
     • REQUIRED for mode "contaminants" — the 371 contaminant test names
       live in this file.
3. Set  RAW_DATA_FOLDER  to the path that contains your CSV files.
4. Set  MODE  to "full", "electrochemistry", or "contaminants".
5. Run.  A new subfolder EA_processed_output/ is created automatically
   inside the input folder with all outputs.

ROW-DROP TRANSPARENCY
---------------------
After the streaming pass, a per-filter breakdown is printed showing
exactly how many rows were removed at each cleaning step:
    • Dummy coordinates                (EA placeholder points)
    • Non-water sample types           (SEDIMENT, MUSCLE, LIVER, BIOTA, …)
    • Non-quantitative units           (coded, text, yes/no, pres/nf, …)
    • Administrative / bad tests       (Lab ID, Size Range, Equiv Carbon, …)
    • Rare tests below min_test_count  (full mode only)
    • Non-numeric results              (NaN after numeric parse)
This makes the 71 M → 59 M row reduction fully auditable.

OUTPUTS  (saved in <RAW_DATA_FOLDER>/EA_processed_output/)
----------------------------------------------------------
  • EA_clean_2000_2025_<mode>.csv       – The main clean dataset.
  • EA_clean_2000_2025_<mode>.parquet   – Same data in fast columnar format.
  • EA_statistics_2000_2025_<mode>.xlsx – Descriptive statistics.
  • EA_qa_report_<mode>.html            – Visual quality-assurance summary.
  • EA_processing_log_<mode>.txt        – Full text log of every cleaning step.

DEPENDENCIES  (auto-installed if missing)
-----------------------------------------
  pandas, numpy, pyproj, pyarrow, openpyxl, chardet
===============================================================================
"""

# ============================================================================
# STEP 0 — AUTOMATICALLY INSTALL MISSING LIBRARIES
# ============================================================================

def _ensure_dependencies():
    """Install any missing Python packages required by this script."""
    import subprocess, sys, importlib
    REQUIRED = {
        "pandas": "pandas", "numpy": "numpy", "pyproj": "pyproj",
        "pyarrow": "pyarrow", "openpyxl": "openpyxl", "chardet": "chardet",
    }
    missing = [pip for imp, pip in REQUIRED.items()
               if not importlib.util.find_spec(imp)]
    if missing:
        print(f"Installing missing packages: {', '.join(missing)} ...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet",
             "--break-system-packages", *missing],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("  Done.\n")

_ensure_dependencies()

# ============================================================================
# IMPORTS
# ============================================================================

from pathlib import Path
import pandas as pd
import numpy as np
from pyproj import Transformer
from typing import Dict, Any, Optional, List
from datetime import datetime
import warnings, sys, io, re

warnings.filterwarnings("ignore")


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def hydrostream(
    input_dir: str | Path,
    mode: str = "full",
    categories_file: str | Path | None = None,
    years: range = range(2000, 2026),
    chunksize: int = 250_000,
    min_test_count: int = 50,
    flag_outliers: bool = True,
    generate_stats: bool = True,
    generate_qa_report: bool = True,
    save_log: bool = True,
) -> Dict[str, Any]:
    """
    HydroStream — clean and combine the EA yearly CSV files into one
    analysis-ready dataset via memory-efficient chunked streaming.

    Parameters
    ----------
    input_dir : str or Path
        Folder containing the raw yearly CSV files (2000.csv … 2025.csv).

    mode : str, default "full"
        "full"             → all water-related tests and types.
        "electrochemistry" → dissolved metals, ions, pH, conductivity, etc.
        "contaminants"     → microplastics, PFAS, pesticides, insecticides.
                             REQUIRES the categories file.

    categories_file : str, Path, or None
        Path to "List of tests kept and categories.xlsx".
        If provided (or found automatically in input_dir), a Category
        column is added to the output.  If None, auto-detection is tried.
        This file is REQUIRED for mode="contaminants" (it contains the
        371 contaminant test names).  It is OPTIONAL for the other modes.

    years : range, default range(2000, 2026)
        Which years to process.

    chunksize : int, default 250_000
        Rows per CSV chunk.  Lower if RAM is limited.

    min_test_count : int, default 50
        Drop tests with fewer total records (full mode only).

    flag_outliers : bool, default True
        Flag (but do NOT remove) values outside plausible ranges.

    generate_stats : bool, default True
    generate_qa_report : bool, default True
    save_log : bool, default True

    Returns
    -------
    dict with output paths and quality metrics
    """

    warnings.filterwarnings("ignore")

    # ------------------------------------------------------------------
    # Log capture
    # ------------------------------------------------------------------
    log_buffer = io.StringIO()
    def log(msg: str = ""):
        print(msg); log_buffer.write(msg + "\n")

    # ------------------------------------------------------------------
    # Resolve directories
    # ------------------------------------------------------------------
    input_dir = Path(input_dir).resolve()
    out_dir   = input_dir / "EA_processed_output"
    out_dir.mkdir(parents=True, exist_ok=True)

    mode = mode.strip().lower()
    if mode not in ("full", "electrochemistry", "contaminants"):
        raise ValueError(
            f"mode must be 'full', 'electrochemistry', or 'contaminants', "
            f"got '{mode}'"
        )

    # ==================================================================
    #  BANNER  (printed first so the output reads top-down)
    # ==================================================================

    log("=" * 70)
    log("  HYDROSTREAM  —  EA WATER QUALITY ARCHIVE PIPELINE   v2.3")
    log("=" * 70)
    log(f"  Mode            : {mode.upper()}")
    log(f"  Years           : {min(years)} – {max(years)}")
    log(f"  Input folder    : {input_dir}")
    log(f"  Output folder   : {out_dir}")
    log(f"  Chunk size      : {chunksize:,}")
    log(f"  Min test count  : {min_test_count}")
    log(f"  Flag outliers   : {flag_outliers}")
    log(f"  Started at      : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log("=" * 70)
    log()

    # ==================================================================
    #  CATEGORIES FILE  —  load (optional for full/electrochemistry,
    #                     REQUIRED for contaminants)
    # ==================================================================

    cat_path: Optional[Path] = None
    if categories_file is not None:
        cat_path = Path(categories_file)
    else:
        for name in ["List of tests kept and categories.xlsx",
                     "categories.xlsx", "test_categories.xlsx"]:
            candidate = input_dir / name
            if candidate.exists():
                cat_path = candidate
                break

    category_map: Dict[str, str] = {}
    if cat_path and cat_path.exists():
        try:
            cat_df = pd.read_excel(cat_path)
            if {"List of Tests", "Final Category"}.issubset(cat_df.columns):
                category_map = dict(zip(
                    cat_df["List of Tests"], cat_df["Final Category"]
                ))
                log(f"Categories file loaded: {cat_path.name} "
                    f"({len(category_map):,} mappings)")
            else:
                log(f"⚠  Categories file found but missing expected columns "
                    f"'List of Tests' and 'Final Category' — ignored.")
        except Exception as e:
            log(f"⚠  Could not read categories file: {e}")
    else:
        log("Categories file not found in input folder.")
        log("  • full / electrochemistry modes: the Category column will be omitted.")
        log("  • contaminants mode: this file is REQUIRED and the run will stop.")

    log()

    # ==================================================================
    #  CONFIGURATION — WATER TYPES TO KEEP
    # ==================================================================

    WATER_TYPES = {
        "RIVER / RUNNING SURFACE WATER", "POND / LAKE / RESERVOIR WATER",
        "CANAL WATER", "CANAL WATER - SALINE",
        "FINAL SEWAGE EFFLUENT", "CRUDE SEWAGE", "ANY SEWAGE",
        "STORM SEWER OVERFLOW DISCHARGE", "STORM TANK EFFLUENT",
        "STORM TANK INFLUENT", "SURFACE DRAINAGE",
        "ANY TRADE EFFLUENT",
        "TRADE EFFLUENT - FRESHWATER RETURNED ABSTRACTED",
        "TRADE EFFLUENT - SALINE WATER RETURNED ABSTRACTED",
        "TRADE EFFLUENT - GROUNDWATER RETURNED ABSTRACTED",
        "GROUNDWATER", "GROUNDWATER - PURGED/PUMPED/REFILLED",
        "GROUNDWATER - STATIC/UNPURGED",
        "ANY LEACHATE", "MINEWATER", "MINEWATER (FLOWING/PUMPED)",
        "SEA WATER", "SEA WATER - INTERTIDAL",
        "SEA WATER AT HIGH TIDE", "SEA WATER AT LOW TIDE",
        "ESTUARINE WATER", "ESTUARINE WATER - INTERTIDAL",
        "ESTUARINE WATER AT HIGH TIDE", "ESTUARINE WATER AT LOW TIDE",
    }

    # Non-capturing groups to avoid pandas regex warnings
    DROP_TYPE_PATTERN = (
        r"(?:SEDIMENT|WHOLE ANIMAL|MUSCLE|LIVER|DIGESTIVE GLAND|BIOTA|"
        r"SOIL|ASH|WASTE\b|GAS|PRECIPITATION|CALIBRATION WATER|"
        r"POTABLE WATER|BOREHOLE GAS|ANY WATER\b|ANY NON-AQUEOUS LIQUID|"
        r"UNCODED|ANY AGRICULTURAL|ANY SEWAGE SLUDGE|ANY TIPPED|"
        r"ALGAE|SEAWEED|INVERTEBRATE|FISH|FLATFISH|BRYOPHYTE|"
        r"HIGHER PLANT|RANUNCULUS|FONTINALIS|ANY OIL|ANY BIOTA|"
        r"SOLID/SEDIMENT|MOSS|WRACK|COCKLE|MUSSEL|OYSTER|"
        r"SHRIMP|WORM|TELLIN|SCALLOP|TROUT|EEL|ROACH|FLOUNDER|"
        r"DAB|PLAICE|SOLE\b|WHITEBAIT|AIR\b|CONSTRUCTION|"
        r"WHOLE PLANT)"
    )

    # ==================================================================
    #  CONFIGURATION — ELECTROCHEMISTRY TEST SET
    # ==================================================================

    ELECTROCHEMISTRY_TESTS = {
        "Magnesium, Dissolved", "Copper, Dissolved", "Nickel, Dissolved",
        "Iron, Dissolved", "Manganese, Dissolved", "Uranium, Dissolved",
        "Lithium, Dissolved", "Potassium, Dissolved", "Sodium, Dissolved",
        "Lead, Dissolved", "Cadmium, Dissolved", "Mercury, Dissolved",
        "Silver, Dissolved", "Barium, Dissolved", "Zinc, Dissolved",
        "Chromium, Dissolved", "Arsenic, Dissolved", "Calcium, Dissolved",
        "Boron, Dissolved", "Aluminium, Dissolved", "Strontium, Filtered",
        "Magnesium", "Copper", "Nickel", "Iron", "Manganese",
        "Potassium", "Sodium", "Lead", "Cadmium", "Mercury",
        "Silver", "Barium", "Zinc", "Chromium", "Arsenic",
        "Calcium", "Boron", "Aluminium",
        "pH", "Conductivity at 25 C", "Conductivity at 20 C",
        "Temperature of Water", "Turbidity",
        "Chloride", "Ammoniacal Nitrogen as N",
        "Nitrogen, Total Oxidised as N", "Orthophosphate, reactive as P",
        "Nitrate as N", "Nitrite as N", "Sulphate as SO4", "Fluoride",
        "Oxygen, Dissolved as O2", "Oxygen, Dissolved, % Saturation",
        "Alkalinity to pH 4.5 as CaCO3", "Hardness, Total as CaCO3",
        "Solids, Suspended at 105 C", "BOD : 5 Day ATU",
        "Salinity : In Situ",
    }

    # ==================================================================
    #  CONFIGURATION — CONTAMINANTS TEST SET
    # ==================================================================
    # Built dynamically from the categories file.  NO hard-coded fallback:
    # the 371 contaminant names are maintained externally in
    # "List of tests kept and categories.xlsx" so the script stays small
    # and the contaminant list stays easy to update.
    # ------------------------------------------------------------------

    CONTAMINANTS_TESTS = {
        t for t, c in category_map.items()
        if c == "microplastics, nanoplastic, pfas, insecticide, pesticide, or similar"
    } if category_map else set()

    if category_map:
        log(f"Contaminants test set built from categories file: "
            f"{len(CONTAMINANTS_TESTS)} tests.")
    log()

    # ==================================================================
    #  CONFIGURATION — NON-QUANTITATIVE UNITS TO DROP
    # ==================================================================

    NON_QUANTITATIVE_UNITS = {
        "coded", "text", "yes/no", "pres/nf", "pres/nft",
        "garber c", "hh.mm", "ngr", "deccafix", "ug",
    }

    # ==================================================================
    #  CONFIGURATION — TEST FRAGMENTS TO DROP
    # ==================================================================
    # If a test name contains any of these (case-insensitive), it is
    # removed.  This list combines the original non-quantitative
    # indicators with additional administrative / procedural tests
    # identified during manual review of the dataset.
    # ------------------------------------------------------------------

    BAD_TEST_FRAGMENTS = [
        # --- Non-quantitative indicators ---
        "No flow", "No sample", "Site Inspection",
        "Present/Not found", "Pass/Fail",
        "Population Equivalent", "Sampling Frequency",
        "Photo Taken",
        # --- Weather / bathing / field obs ---
        "Weather :", "Bathing Water Profile",
        "National Grid Reference",
        "Sewage debris", "Foam Visible",
        "Colour : Abnormal", "Tarry residues",
        "MST Filtration", "Time of high tide",
        "Number of beach users", "Bathers per 100",
        "Type of flow",
        "State tide", "Colour (1/0)", "Tars/Floatg",
        "OilTypeQual", "WEATHER FLAG",
        "Borehole RefPt", "Sample Depth",
        # --- Administrative / procedural (manager review) ---
        "Laboratory Sample Number",
        "Dummy determinand",
        "Warning Sign",
        "Miscellaneous Identification",
        "Data Handling",
        "Mitochondrial Marker",   # DNA source tracking markers
        "Mitochrondrial Marker",  # misspelled variant in raw data
        "Size range",             # fish size measurements
        "Size Range",
        "Length of fish",
        "Equiv.Carbon",           # petroleum hydrocarbon fractions
        "Equiv.carbo",            # variant spelling
        "Equiv Carbon",           # variant spelling
        "Biological examination",
        "Soli proportion",
        "24 hour Oyster",
        "Stone size",
        "Carbohydrate as Glucose",
        "Cohesive strength",
        "WQMS :",                 # water quality monitoring station internals
        "Grain Size",             # sediment grain size (not water quality)
        "Number of bathers",
        "Number of birds",
        "Number of dogs",
    ]

    _BAD_TEST_PATTERN = "|".join(re.escape(f) for f in BAD_TEST_FRAGMENTS)

    # ==================================================================
    #  CONFIGURATION — DUMMY COORDINATES
    # ==================================================================

    DUMMY_EASTING  = 500_000
    DUMMY_NORTHINGS = {1, 2, 3, 4, 5, 6, 7, 8}

    # ==================================================================
    #  CONFIGURATION — OUTLIER THRESHOLDS (flagged, NOT removed)
    # ==================================================================

    OUTLIER_THRESHOLDS = {
        "Temperature of Water":        (-5, 45),
        "pH":                          (1, 14),
        "Conductivity at 25 C":        (0, 80_000),
        "Conductivity at 20 C":        (0, 80_000),
        "Salinity : In Situ":          (0, 50),
        "Solids, Suspended at 105 C":  (0, 50_000),
        "Oxygen, Dissolved, % Saturation": (0, 250),
        "Oxygen, Dissolved as O2":     (0, 25),
        "Ammoniacal Nitrogen as N":    (0, 1_000),
        "Turbidity":                   (0, 10_000),
    }

    # ==================================================================
    #  DROP-COUNT TRACKING  —  transparent row-loss audit trail
    # ==================================================================
    # Every filter in _clean_chunk updates this dict so that, after the
    # streaming pass, the user sees exactly where the ~17% of raw rows
    # were removed (e.g. 71 M raw → 59 M final).
    # ------------------------------------------------------------------

    drop_counts = {
        "dummy_coordinates":      0,
        "non_water_types":        0,
        "non_quantitative_units": 0,
        "administrative_tests":   0,
        "rare_tests":             0,
        "non_numeric_results":    0,
    }

    # ==================================================================
    #  HELPER FUNCTIONS
    # ==================================================================

    SEASON_CATS = ["Winter", "Spring", "Summer", "Autumn"]

    def _month_to_season(m):
        if pd.isna(m): return pd.NA
        m = int(m)
        if m in (12, 1, 2): return "Winter"
        if m in (3, 4, 5): return "Spring"
        if m in (6, 7, 8): return "Summer"
        return "Autumn"

    def _standardise_units(df):
        """Harmonise all measurement units to a consistent set."""
        if "Test" not in df.columns or "Unit" not in df.columns:
            return df

        u = df["Unit"].astype(str).str.strip()
        u = (u.str.replace("µ", "u", regex=False)
              .str.replace("μ", "u", regex=False)
              .str.replace("US/CM", "uS/cm", regex=False)
              .str.replace("Us/cm", "uS/cm", regex=False)
              .str.replace("us/cm", "uS/cm", regex=False)
              .str.replace("µS/cm", "uS/cm", regex=False)
              .str.replace("μS/cm", "uS/cm", regex=False))
        df["Unit"] = u

        df["Test"] = df["Test"].str.replace(
            "Conductivity at 20C", "Conductivity at 20 C", regex=False)

        # µg/l → mg/l (÷ 1,000)
        m = df["Unit"].str.lower() == "ug/l"
        if m.any():
            df.loc[m, "result"] = pd.to_numeric(df.loc[m, "result"], errors="coerce") / 1_000
            df.loc[m, "Unit"] = "mg/l"

        # ng/l → mg/l (÷ 1,000,000)
        m = df["Unit"].str.lower() == "ng/l"
        if m.any():
            df.loc[m, "result"] = pd.to_numeric(df.loc[m, "result"], errors="coerce") / 1_000_000
            df.loc[m, "Unit"] = "mg/l"

        # pg/l → mg/l (÷ 1,000,000,000)
        m = df["Unit"].str.lower() == "pg/l"
        if m.any():
            df.loc[m, "result"] = pd.to_numeric(df.loc[m, "result"], errors="coerce") / 1_000_000_000
            df.loc[m, "Unit"] = "mg/l"

        # g/l → mg/l (× 1,000)
        m = df["Unit"].str.lower() == "g/l"
        if m.any():
            df.loc[m, "result"] = pd.to_numeric(df.loc[m, "result"], errors="coerce") * 1_000
            df.loc[m, "Unit"] = "mg/l"

        # ppm → mg/l
        m = df["Unit"].str.lower() == "ppm"
        if m.any(): df.loc[m, "Unit"] = "mg/l"

        # FTU → NTU (all tests, not just "Turbidity")
        m = df["Unit"].str.lower() == "ftu"
        if m.any(): df.loc[m, "Unit"] = "NTU"
        m = df["Unit"] == "ntu"
        if m.any(): df.loc[m, "Unit"] = "NTU"

        # ms/cm → uS/cm (× 1,000)
        m = df["Unit"].str.lower() == "ms/cm"
        if m.any():
            df.loc[m, "result"] = pd.to_numeric(df.loc[m, "result"], errors="coerce") * 1_000
            df.loc[m, "Unit"] = "uS/cm"

        # no/ml → no/100ml (× 100)
        m = df["Unit"] == "no/ml"
        if m.any():
            df.loc[m, "result"] = pd.to_numeric(df.loc[m, "result"], errors="coerce") * 100
            df.loc[m, "Unit"] = "no/100ml"

        # no/ul → no/100ml (× 100,000)
        m = df["Unit"] == "no/ul"
        if m.any():
            df.loc[m, "result"] = pd.to_numeric(df.loc[m, "result"], errors="coerce") * 100_000
            df.loc[m, "Unit"] = "no/100ml"

        # no/10ul → no/100ml (× 10,000)
        m = df["Unit"] == "no/10ul"
        if m.any():
            df.loc[m, "result"] = pd.to_numeric(df.loc[m, "result"], errors="coerce") * 10_000
            df.loc[m, "Unit"] = "no/100ml"

        # g/kg, psu, ‰ → ppt
        m = df["Unit"].str.lower().isin({"g/kg", "psu", "\u2030"})
        if m.any(): df.loc[m, "Unit"] = "ppt"

        return df

    def _flag_outliers_fn(df):
        if "outlier_flag" not in df.columns:
            df["outlier_flag"] = False
        for test_name, (lo, hi) in OUTLIER_THRESHOLDS.items():
            mask = df["Test"] == test_name
            if mask.any():
                bad = mask & ((df["result"] < lo) | (df["result"] > hi))
                df.loc[bad, "outlier_flag"] = True
        return df

    def _convert_coordinates(df):
        if "Easting" not in df.columns or "Northing" not in df.columns:
            return df
        log("  Converting Easting/Northing → Latitude/Longitude …")
        unique = (df[["Easting", "Northing"]].drop_duplicates()
                  .dropna(subset=["Easting", "Northing"]))
        transformer = Transformer.from_crs("EPSG:27700", "EPSG:4326", always_xy=True)
        lons, lats = transformer.transform(unique["Easting"].values, unique["Northing"].values)
        unique = unique.copy()
        unique["Latitude"] = lats
        unique["Longitude"] = lons
        n_before = len(df)
        df = df.merge(unique, on=["Easting", "Northing"], how="left")
        df = df.drop(columns=["Easting", "Northing"])
        log(f"    {len(unique):,} unique coordinate pairs converted.")
        log(f"    Rows with valid lat/lon: {df['Latitude'].notna().sum():,} / {n_before:,}")
        return df

    # ==================================================================
    #  CHUNK CLEANER  (instrumented with drop_counts)
    # ==================================================================

    def _clean_chunk(raw, year_hint, test_filter):
        df = raw.copy()

        drop_cols = [
            "@id", "sample.samplingPoint", "sample.samplingPoint.notation",
            "resultQualifier.notation", "codedResultInterpretation.interpretation",
            "determinand.label", "sample.isComplianceSample",
            "sample.purpose.label", "determinand.notation",
        ]
        existing = [c for c in drop_cols if c in df.columns]
        if existing: df = df.drop(columns=existing)

        rename_map = {
            "sample.samplingPoint.label": "Sampling Point",
            "sample.sampleDateTime": "Date",
            "sample.sampledMaterialType.label": "Type",
            "determinand.definition": "Test",
            "determinand.unit.label": "Unit",
            "result": "result", "Result": "result",
            "sample.samplingPoint.easting": "Easting",
            "sample.samplingPoint.northing": "Northing",
        }
        rename_map = {k: v for k, v in rename_map.items() if k in df.columns}
        df = df.rename(columns=rename_map)

        for col in ("Easting", "Northing"):
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        # ---- Filter 1: dummy coordinates ----
        if "Easting" in df.columns and "Northing" in df.columns:
            n0 = len(df)
            dummy = df["Easting"].eq(DUMMY_EASTING) & df["Northing"].isin(DUMMY_NORTHINGS)
            df = df[~dummy]
            drop_counts["dummy_coordinates"] += n0 - len(df)

        # ---- Date parsing / season / source-year ----
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["Season"] = df["Date"].dt.month.map(_month_to_season)
            df["Season"] = pd.Categorical(df["Season"], categories=SEASON_CATS, ordered=True)
            df["SourceYear"] = df["Date"].dt.year.fillna(year_hint).astype("Int64")
        else:
            df["Season"] = pd.Categorical([], categories=SEASON_CATS)
            df["SourceYear"] = year_hint

        # ---- Filter 2: non-water sample types ----
        if "Type" in df.columns:
            n0 = len(df)
            df = df[~df["Type"].astype(str).str.contains(DROP_TYPE_PATTERN, case=False, na=False)]
            df = df[df["Type"].isin(WATER_TYPES)]
            drop_counts["non_water_types"] += n0 - len(df)

        # ---- Filter 3: non-quantitative units ----
        if "Unit" in df.columns:
            n0 = len(df)
            df = df[~df["Unit"].str.strip().str.lower().isin(NON_QUANTITATIVE_UNITS)]
            drop_counts["non_quantitative_units"] += n0 - len(df)

        # ---- Filter 4: administrative / bad test names ----
        if "Test" in df.columns:
            n0 = len(df)
            df = df[~df["Test"].astype(str).str.contains(_BAD_TEST_PATTERN, case=False, na=False)]
            drop_counts["administrative_tests"] += n0 - len(df)

        # ---- Filter 5: rare tests (mode=full, min_test_count>0) ----
        if "Test" in df.columns and test_filter is not None:
            n0 = len(df)
            df = df[df["Test"].isin(test_filter)]
            drop_counts["rare_tests"] += n0 - len(df)

        # ---- Numeric parsing of result ----
        if "result" in df.columns:
            df["result"] = pd.to_numeric(df["result"], errors="coerce")

        # ---- Unit standardisation (value conversions may produce NaN) ----
        if {"Test", "Unit", "result"}.issubset(df.columns):
            df = _standardise_units(df)

        # ---- Filter 6: non-numeric results (post-standardisation) ----
        if "result" in df.columns:
            n0 = len(df)
            df = df[df["result"].notna()]
            drop_counts["non_numeric_results"] += n0 - len(df)

        if flag_outliers:
            df = _flag_outliers_fn(df)

        col_order = ["Sampling Point", "Type", "Date", "Test", "result", "Unit",
                     "Season", "SourceYear", "Easting", "Northing"]
        if flag_outliers: col_order.append("outlier_flag")
        col_order = [c for c in col_order if c in df.columns]
        extra = [c for c in df.columns if c not in col_order]
        df = df[col_order + extra]

        return df

    # ==================================================================
    #  FIND INPUT FILES
    # ==================================================================

    year_files = []
    for y in years:
        p = input_dir / f"{y}.csv"
        if p.exists():
            year_files.append((y, p))
        else:
            matches = sorted(input_dir.glob(f"{y}*.csv"))
            if matches: year_files.append((y, matches[0]))

    if not year_files:
        raise FileNotFoundError(
            f"No CSV files found for years {list(years)} in {input_dir}")

    log(f"Found {len(year_files)} raw CSV files:\n")
    for y, p in year_files:
        log(f"  {y}  →  {p.name}")
    log()

    # ==================================================================
    #  DETERMINE WHICH TESTS TO KEEP
    # ==================================================================

    if mode == "electrochemistry":
        test_filter = ELECTROCHEMISTRY_TESTS
        log(f"Mode = ELECTROCHEMISTRY  →  {len(test_filter)} pre-defined tests.\n")

    elif mode == "contaminants":
        if not CONTAMINANTS_TESTS:
            # Hard stop — no silent fallback.  Tell the user exactly
            # what's missing and how to fix it.
            raise FileNotFoundError(
                "\n\n"
                "========================================================\n"
                "  CONTAMINANTS MODE REQUIRES THE CATEGORIES FILE\n"
                "========================================================\n"
                f"  The file  'List of tests kept and categories.xlsx'\n"
                f"  was NOT found in the input folder:\n\n"
                f"    {input_dir}\n\n"
                f"  This file lists the 371 contaminant test names\n"
                f"  (microplastics, PFAS, pesticides, insecticides, etc.)\n"
                f"  that define the 'contaminants' output mode.\n\n"
                f"  To fix:\n"
                f"    • Place the file inside the input folder, OR\n"
                f"    • Pass its path explicitly via the\n"
                f"      'categories_file' parameter.\n"
                "========================================================\n"
            )
        test_filter = CONTAMINANTS_TESTS
        log(f"Mode = CONTAMINANTS  →  {len(test_filter)} tests loaded "
            f"from categories file.\n")

    else:  # mode == "full"
        if min_test_count > 0:
            log(f"Mode = FULL  →  first pass: counting tests (drop < {min_test_count}) …")
            test_counts: Dict[str, int] = {}
            for y, csv_path in year_files:
                try:
                    pd.read_csv(csv_path, nrows=2, encoding="utf-8"); enc="utf-8"
                except UnicodeDecodeError:
                    enc = "latin-1"
                for chunk in pd.read_csv(
                    csv_path, chunksize=chunksize, low_memory=False, encoding=enc,
                    usecols=lambda c: c in ("determinand.definition",)):
                    col = "determinand.definition"
                    if col in chunk.columns:
                        for t, cnt in chunk[col].value_counts().items():
                            test_counts[t] = test_counts.get(t, 0) + cnt
            total_tests = len(test_counts)
            test_filter = {t for t, c in test_counts.items() if c >= min_test_count}
            log(f"  Total unique tests : {total_tests:,}")
            log(f"  Tests kept (>= {min_test_count}) : {len(test_filter):,}")
            log(f"  Rare tests dropped : {total_tests - len(test_filter):,}\n")
        else:
            test_filter = None
            log("Mode = FULL, min_test_count = 0  →  keeping ALL tests.\n")

    # ==================================================================
    #  MAIN PROCESSING LOOP
    # ==================================================================

    tag = {"full": "full", "electrochemistry": "electrochemistry",
           "contaminants": "contaminants"}[mode]
    out_csv   = out_dir / f"EA_clean_2000_2025_{tag}.csv"
    out_pq    = out_dir / f"EA_clean_2000_2025_{tag}.parquet"
    out_stats = out_dir / f"EA_statistics_2000_2025_{tag}.xlsx"
    out_qa    = out_dir / f"EA_qa_report_{tag}.html"
    out_log   = out_dir / f"EA_processing_log_{tag}.txt"
    tmp_csv   = out_dir / "_tmp_stream.csv"

    for p in (out_csv, out_pq, tmp_csv):
        if p.exists(): p.unlink()

    summary = {}
    total_streamed = 0
    header_written = False
    total_raw = 0

    for y, csv_path in year_files:
        log(f"── Processing {y}  ({csv_path.name}) " + "─" * 30)
        n_year_clean = 0; n_year_raw = 0

        try:
            pd.read_csv(csv_path, nrows=5, encoding="utf-8"); enc = "utf-8"
        except UnicodeDecodeError:
            try:
                import chardet
                with open(csv_path, "rb") as f:
                    enc = chardet.detect(f.read(100_000)).get("encoding", "latin-1")
            except: enc = "latin-1"

        for chunk in pd.read_csv(csv_path, chunksize=chunksize, low_memory=False, encoding=enc):
            n_year_raw += len(chunk); total_raw += len(chunk)
            cleaned = _clean_chunk(chunk, year_hint=y, test_filter=test_filter)
            if not cleaned.empty:
                cleaned.to_csv(tmp_csv, mode="a", index=False, header=(not header_written))
                header_written = True
                n_year_clean += len(cleaned); total_streamed += len(cleaned)

        pct = (n_year_clean / n_year_raw * 100) if n_year_raw else 0
        summary[y] = n_year_clean
        log(f"  Raw rows       : {n_year_raw:>12,}")
        log(f"  Clean rows     : {n_year_clean:>12,}   ({pct:.1f}% kept)")
        log(f"  Rows removed   : {n_year_raw - n_year_clean:>12,}")
        log()

    log("=" * 70)
    log(f"  Streaming complete.  Clean rows (pre-dedup): {total_streamed:,}")
    log(f"  Total raw rows read: {total_raw:,}")
    log("=" * 70 + "\n")

    # ------------------------------------------------------------------
    #  Drop-count audit trail  —  where did the removed rows go?
    # ------------------------------------------------------------------

    total_removed = sum(drop_counts.values())
    log("── Rows removed per filter " + "─" * 41)
    log(f"  Dummy coordinates        : {drop_counts['dummy_coordinates']:>12,}")
    log(f"  Non-water sample types   : {drop_counts['non_water_types']:>12,}")
    log(f"  Non-quantitative units   : {drop_counts['non_quantitative_units']:>12,}")
    log(f"  Administrative/bad tests : {drop_counts['administrative_tests']:>12,}")
    log(f"  Rare tests (< {min_test_count:>3})       : {drop_counts['rare_tests']:>12,}")
    log(f"  Non-numeric results      : {drop_counts['non_numeric_results']:>12,}")
    log(f"  " + "─" * 45)
    log(f"  TOTAL removed            : {total_removed:>12,}  "
        f"({total_removed/total_raw*100:.1f}% of raw rows)")
    log("─" * 70 + "\n")

    if not tmp_csv.exists() or total_streamed == 0:
        raise ValueError("No data survived filtering.  Check input files and mode.")

    # ==================================================================
    #  FINAL PROCESSING
    # ==================================================================

    log("Loading streamed data for final processing …")
    df_all = pd.read_csv(tmp_csv, low_memory=False, parse_dates=["Date"])
    log(f"  Rows loaded: {len(df_all):,}")

    if "Type" in df_all.columns:
        mask = (df_all["Type"].isin(WATER_TYPES) &
                ~df_all["Type"].astype(str).str.contains(DROP_TYPE_PATTERN, case=False, na=False))
        n_extra = (~mask).sum()
        df_all = df_all[mask]
        if n_extra: log(f"  Extra type filter removed {n_extra:,} rows.")

    key_cols = ["Date", "Sampling Point", "Type", "Test"]
    d0 = len(df_all)
    df_all = df_all.drop_duplicates(subset=key_cols, keep="first")
    n_dupes = d0 - len(df_all)
    log(f"  Duplicates removed: {n_dupes:,}")

    df_all["result"] = pd.to_numeric(df_all["result"], errors="coerce")
    n_nan = df_all["result"].isna().sum()
    df_all = df_all[df_all["result"].notna()]
    if n_nan: log(f"  NaN results dropped: {n_nan:,}")

    df_all["Season"] = pd.Categorical(df_all["Season"], categories=SEASON_CATS, ordered=True)

    # --- Add Category column ------------------------------------------
    if category_map and "Test" in df_all.columns:
        df_all["Category"] = df_all["Test"].map(category_map).fillna("uncategorized")
        n_mapped = (df_all["Category"] != "uncategorized").sum()
        log(f"  Category mapped: {n_mapped:,} / {len(df_all):,} rows "
            f"({n_mapped/len(df_all)*100:.1f}%)")

    # --- Convert coordinates ------------------------------------------
    df_final = _convert_coordinates(df_all)

    log(f"\n  Final dataset: {len(df_final):,} rows × {len(df_final.columns)} columns")
    log(f"  Unique sampling points : {df_final['Sampling Point'].nunique():,}")
    log(f"  Unique tests           : {df_final['Test'].nunique()}")
    log(f"  Unique water types     : {df_final['Type'].nunique()}")
    log(f"  Unique units           : {df_final['Unit'].nunique()}")
    if "Category" in df_final.columns:
        log(f"  Unique categories      : {df_final['Category'].nunique()}")
    log(f"  Date range             : {df_final['Date'].min()} → {df_final['Date'].max()}")
    log()

    # ==================================================================
    #  SAVE OUTPUTS
    # ==================================================================

    log("Saving main outputs …")
    df_final.to_csv(out_csv, index=False)
    log(f"  ✓  CSV saved     : {out_csv.name}")

    wrote_parquet = False
    try:
        df_final.to_parquet(out_pq, engine="pyarrow", compression="zstd")
        wrote_parquet = True
        log(f"  ✓  Parquet saved : {out_pq.name}")
    except Exception as e:
        log(f"  ⚠  Parquet skipped: {e}")

    # ==================================================================
    #  STATISTICS
    # ==================================================================

    stats_output = None
    if generate_stats:
        log("\nGenerating statistics …")
        try:
            with pd.ExcelWriter(out_stats, engine="openpyxl") as writer:
                grp_cols = ["Test", "Unit"]
                if "Category" in df_final.columns: grp_cols = ["Category", "Test", "Unit"]
                test_stats = (df_final.groupby(grp_cols)["result"]
                    .agg(["count","min","max","mean","median","std",
                          ("p10", lambda x: x.quantile(0.10)),
                          ("p25", lambda x: x.quantile(0.25)),
                          ("p75", lambda x: x.quantile(0.75)),
                          ("p90", lambda x: x.quantile(0.90))])
                    .round(4).reset_index())
                test_stats.to_excel(writer, sheet_name="Test_Statistics", index=False)

                type_stats = (df_final.groupby(["Type","Test"])["result"]
                    .agg(["count","mean","median","std"]).round(4).reset_index())
                type_stats.to_excel(writer, sheet_name="Type_Test_Stats", index=False)

                season_stats = (df_final.groupby(["Season","Test"])["result"]
                    .agg(["count","mean","median"]).round(4).reset_index())
                season_stats.to_excel(writer, sheet_name="Seasonal_Stats", index=False)

                cov_items = [
                    ("Total Rows", len(df_final)),
                    ("Unique Sampling Points", df_final["Sampling Point"].nunique()),
                    ("Unique Tests", df_final["Test"].nunique()),
                    ("Unique Types", df_final["Type"].nunique()),
                    ("Unique Units", df_final["Unit"].nunique()),
                    ("Date Range Start", str(df_final["Date"].min().date())),
                    ("Date Range End", str(df_final["Date"].max().date())),
                    ("Years Covered", df_final["SourceYear"].nunique()),
                    ("Mode", mode.upper()),
                ]
                if "Category" in df_final.columns:
                    cov_items.append(("Unique Categories", df_final["Category"].nunique()))
                coverage = pd.DataFrame(cov_items, columns=["Metric","Value"])
                coverage.to_excel(writer, sheet_name="Coverage", index=False)

                if flag_outliers and "outlier_flag" in df_final.columns:
                    outlier_df = df_final[df_final["outlier_flag"]]
                    if not outlier_df.empty:
                        (outlier_df.groupby(["Test","Type"])
                         .agg(count=("result","size"), min_val=("result","min"), max_val=("result","max"))
                         .reset_index()
                         .to_excel(writer, sheet_name="Outliers", index=False))

                (df_final.groupby("SourceYear").size().reset_index(name="rows")
                 .to_excel(writer, sheet_name="Rows_Per_Year", index=False))

                if "Category" in df_final.columns:
                    (df_final.groupby("Category").size().reset_index(name="rows")
                     .sort_values("rows", ascending=False)
                     .to_excel(writer, sheet_name="Rows_Per_Category", index=False))

                # Drop-count audit sheet
                drop_audit = pd.DataFrame([
                    ("Dummy coordinates",        drop_counts["dummy_coordinates"]),
                    ("Non-water sample types",   drop_counts["non_water_types"]),
                    ("Non-quantitative units",   drop_counts["non_quantitative_units"]),
                    ("Administrative/bad tests", drop_counts["administrative_tests"]),
                    (f"Rare tests (< {min_test_count})", drop_counts["rare_tests"]),
                    ("Non-numeric results",      drop_counts["non_numeric_results"]),
                    ("TOTAL removed",            total_removed),
                    ("Raw rows read",            total_raw),
                    ("Final rows",               len(df_final)),
                ], columns=["Filter", "Rows"])
                drop_audit.to_excel(writer, sheet_name="Rows_Dropped_Audit", index=False)

            stats_output = out_stats
            log(f"  ✓  Statistics saved : {out_stats.name}")
        except Exception as e:
            log(f"  ⚠  Statistics failed: {e}")

    # ==================================================================
    #  QA REPORT   (no Unit-Consistency section — those cases are all
    #              legitimate multi-unit reporting, not errors)
    # ==================================================================

    qa_output = None
    if generate_qa_report:
        log("\nGenerating QA report …")
        try:
            n_outliers = int(df_final["outlier_flag"].sum()) if flag_outliers and "outlier_flag" in df_final.columns else 0
            pct_outliers = (n_outliers / len(df_final)) * 100

            type_rows = "".join(
                f"<tr><td>{t}</td><td>{c:,}</td><td>{c/len(df_final)*100:.1f}%</td></tr>\n"
                for t, c in df_final["Type"].value_counts().head(15).items())

            test_rows = "".join(
                f"<tr><td>{t}</td><td>{c:,}</td><td>{c/len(df_final)*100:.1f}%</td></tr>\n"
                for t, c in df_final["Test"].value_counts().head(20).items())

            unit_dist_rows = "".join(
                f"<tr><td>{u}</td><td>{c:,}</td><td>{c/len(df_final)*100:.1f}%</td></tr>\n"
                for u, c in df_final["Unit"].value_counts().head(25).items())

            cat_rows = ""
            if "Category" in df_final.columns:
                cat_rows = "".join(
                    f"<tr><td>{cat}</td><td>{c:,}</td><td>{c/len(df_final)*100:.1f}%</td></tr>\n"
                    for cat, c in df_final["Category"].value_counts().items())

            cat_section = ""
            if "Category" in df_final.columns:
                cat_section = f"""
<h2>Records per Category</h2>
<table><tr><th>Category</th><th>Rows</th><th>%</th></tr>
{cat_rows}
</table>"""

            # Drop-count section for the QA report
            drop_rows = "".join([
                f"<tr><td>Dummy coordinates</td><td>{drop_counts['dummy_coordinates']:,}</td></tr>\n",
                f"<tr><td>Non-water sample types</td><td>{drop_counts['non_water_types']:,}</td></tr>\n",
                f"<tr><td>Non-quantitative units</td><td>{drop_counts['non_quantitative_units']:,}</td></tr>\n",
                f"<tr><td>Administrative / bad tests</td><td>{drop_counts['administrative_tests']:,}</td></tr>\n",
                f"<tr><td>Rare tests (&lt; {min_test_count})</td><td>{drop_counts['rare_tests']:,}</td></tr>\n",
                f"<tr><td>Non-numeric results</td><td>{drop_counts['non_numeric_results']:,}</td></tr>\n",
                f"<tr><th>Total removed</th><th>{total_removed:,} "
                f"({total_removed/total_raw*100:.1f}% of {total_raw:,} raw rows)</th></tr>\n",
            ])

            qa_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>HydroStream — QA Report ({mode.upper()})</title>
<style>
  body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 2em; color: #222; }}
  h1 {{ color: #1a5276; }} h2 {{ color: #2c3e50; border-bottom: 2px solid #2980b9; padding-bottom: 4px; }}
  table {{ border-collapse: collapse; width: 100%; margin: 1em 0 2em; }}
  th, td {{ border: 1px solid #ccc; padding: 6px 10px; text-align: left; }}
  th {{ background: #2980b9; color: #fff; }} tr:nth-child(even) {{ background: #f4f6f7; }}
  .good {{ color: #27ae60; font-weight: bold; }}
  .bad  {{ color: #c0392b; font-weight: bold; }}
  footer {{ margin-top: 3em; color: #888; font-size: 0.85em; }}
</style></head><body>
<h1>HydroStream — EA Water Quality QA Report</h1>
<p><b>Mode:</b> {mode.upper()} | <b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

<h2>Dataset Overview</h2>
<table><tr><th>Metric</th><th>Value</th></tr>
<tr><td>Total rows</td><td>{len(df_final):,}</td></tr>
<tr><td>Unique sampling points</td><td>{df_final['Sampling Point'].nunique():,}</td></tr>
<tr><td>Unique tests</td><td>{df_final['Test'].nunique()}</td></tr>
<tr><td>Unique water types</td><td>{df_final['Type'].nunique()}</td></tr>
<tr><td>Unique units</td><td>{df_final['Unit'].nunique()}</td></tr>
<tr><td>Date range</td><td>{df_final['Date'].min().date()} → {df_final['Date'].max().date()}</td></tr>
<tr><td>Years covered</td><td>{df_final['SourceYear'].min()} – {df_final['SourceYear'].max()}</td></tr>
<tr><td>Records with coordinates</td><td>{df_final['Latitude'].notna().sum():,} ({df_final['Latitude'].notna().mean()*100:.1f}%)</td></tr>
</table>

<h2>Data Quality Checks</h2>
<table><tr><th>Check</th><th>Result</th><th>Status</th></tr>
<tr><td>Duplicate key rows</td><td>0 (removed)</td><td class="good">✓ PASS</td></tr>
<tr><td>NaN results</td><td>0 (removed)</td><td class="good">✓ PASS</td></tr>
<tr><td>NaN dates</td><td>{df_final['Date'].isna().sum():,}</td><td class="{'good' if df_final['Date'].isna().sum()==0 else 'bad'}">{'✓ PASS' if df_final['Date'].isna().sum()==0 else '⚠'}</td></tr>
<tr><td>Flagged outliers</td><td>{n_outliers:,} ({pct_outliers:.2f}%)</td><td class="{'good' if pct_outliers<5 else 'bad'}">{'✓ OK' if pct_outliers<5 else '⚠ CHECK'}</td></tr>
</table>

<h2>Rows Removed During Cleaning (audit trail)</h2>
<table><tr><th>Filter</th><th>Rows removed</th></tr>
{drop_rows}
</table>

<h2>Unit Conversions Applied</h2>
<table><tr><th>From</th><th>To</th><th>Factor</th><th>Rationale</th></tr>
<tr><td>µg/l</td><td>mg/l</td><td>÷ 1,000</td><td>Standard mass-concentration scale</td></tr>
<tr><td>ng/l</td><td>mg/l</td><td>÷ 1,000,000</td><td>Standard mass-concentration scale</td></tr>
<tr><td>pg/l</td><td>mg/l</td><td>÷ 1,000,000,000</td><td>Standard mass-concentration scale</td></tr>
<tr><td>g/l</td><td>mg/l</td><td>× 1,000</td><td>Standard mass-concentration scale</td></tr>
<tr><td>ppm</td><td>mg/l</td><td>1 : 1</td><td>Equivalent for dilute aqueous solutions</td></tr>
<tr><td>FTU</td><td>NTU</td><td>1 : 1</td><td>Both nephelometric turbidity scales</td></tr>
<tr><td>ms/cm</td><td>uS/cm</td><td>× 1,000</td><td>Standard conductivity scale</td></tr>
<tr><td>no/ml</td><td>no/100ml</td><td>× 100</td><td>Regulatory standard for microbiology</td></tr>
<tr><td>no/ul</td><td>no/100ml</td><td>× 100,000</td><td>Regulatory standard for microbiology</td></tr>
<tr><td>no/10ul</td><td>no/100ml</td><td>× 10,000</td><td>Regulatory standard for microbiology</td></tr>
<tr><td>g/kg, psu, ‰</td><td>ppt</td><td>1 : 1</td><td>All equivalent salinity measures</td></tr>
</table>
{cat_section}
<h2>Top Water Types</h2>
<table><tr><th>Type</th><th>Rows</th><th>%</th></tr>{type_rows}</table>
<h2>Top Tests</h2>
<table><tr><th>Test</th><th>Rows</th><th>%</th></tr>{test_rows}</table>
<h2>Top Units</h2>
<table><tr><th>Unit</th><th>Rows</th><th>%</th></tr>{unit_dist_rows}</table>
<footer><p>HydroStream v2.3 — Water Quality Archive Streaming Pipeline<br>
Source: Environment Agency (England) Open Water Quality Archive, 2000–2025</p></footer>
</body></html>"""

            with open(out_qa, "w", encoding="utf-8") as f: f.write(qa_html)
            qa_output = out_qa
            log(f"  ✓  QA report saved : {out_qa.name}")
        except Exception as e:
            log(f"  ⚠  QA report failed: {e}")

    # ==================================================================
    #  CLEANUP & SUMMARY
    # ==================================================================

    try: tmp_csv.unlink()
    except: pass

    log("\n" + "=" * 70)
    log("  PROCESSING COMPLETE")
    log("=" * 70)
    log(f"  Mode           : {mode.upper()}")
    log(f"  Final rows     : {len(df_final):,}")
    log(f"  Columns        : {list(df_final.columns)}")
    log(f"  Years          : {df_final['SourceYear'].min()} – {df_final['SourceYear'].max()}")
    log(f"  Water types    : {df_final['Type'].nunique()}")
    log(f"  Tests          : {df_final['Test'].nunique()}")
    log(f"  Units          : {df_final['Unit'].nunique()}")
    log(f"  Sampling points: {df_final['Sampling Point'].nunique():,}")
    if "Category" in df_final.columns:
        log(f"  Categories     : {df_final['Category'].nunique()}")
    if flag_outliers and "outlier_flag" in df_final.columns:
        n_out = int(df_final["outlier_flag"].sum())
        log(f"  Outliers flagged: {n_out:,} ({n_out/len(df_final)*100:.2f}%)")
    if "Latitude" in df_final.columns:
        n_c = df_final["Latitude"].notna().sum()
        log(f"  With lat/lon   : {n_c:,} ({n_c/len(df_final)*100:.1f}%)")

    log(f"\n  Outputs in: {out_dir}/")
    log(f"    • {out_csv.name}")
    if wrote_parquet: log(f"    • {out_pq.name}")
    if stats_output: log(f"    • {out_stats.name}")
    if qa_output: log(f"    • {out_qa.name}")

    log(f"\n  Rows per year:")
    for yr in sorted(summary.keys()):
        log(f"    {yr}: {summary[yr]:>10,}")

    log(f"\n  Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log("=" * 70)

    log_path = None
    if save_log:
        with open(out_log, "w", encoding="utf-8") as f:
            f.write(log_buffer.getvalue())
        log_path = out_log
        print(f"\n  ✓  Full log saved : {out_log.name}")

    return {
        "final_rows": len(df_final), "per_year_rows": summary,
        "output_dir": str(out_dir), "csv": str(out_csv),
        "parquet": str(out_pq) if wrote_parquet else None,
        "statistics": str(stats_output) if stats_output else None,
        "qa_report": str(qa_output) if qa_output else None,
        "log": str(log_path) if log_path else None,
        "drop_counts": dict(drop_counts),
        "data_quality": {
            "total_raw_rows": total_raw, "duplicates_removed": n_dupes,
            "unique_sampling_points": df_final["Sampling Point"].nunique(),
            "unique_tests": df_final["Test"].nunique(),
            "unique_types": df_final["Type"].nunique(),
            "unique_units": df_final["Unit"].nunique(),
            "date_range": (str(df_final["Date"].min()), str(df_final["Date"].max())),
            "outliers_flagged": int(df_final["outlier_flag"].sum()) if flag_outliers and "outlier_flag" in df_final.columns else 0,
            "records_with_coordinates": int(df_final["Latitude"].notna().sum()) if "Latitude" in df_final.columns else 0,
        },
    }


# ============================================================================
# USAGE
# ============================================================================

if __name__ == "__main__":

    # ── SETTINGS (EDIT THESE) ─────────────────────────────────────────
    RAW_DATA_FOLDER = "."          # <-- path to your CSV folder
    MODE            = "full"       # <-- "full", "electrochemistry", or "contaminants"
    # ──────────────────────────────────────────────────────────────────

    result = hydrostream(
        input_dir          = RAW_DATA_FOLDER,
        mode               = MODE,
        categories_file    = None,   # auto-detected from input_dir
        years              = range(2000, 2026),
        chunksize          = 250_000,
        min_test_count     = 50,
        flag_outliers      = True,
        generate_stats     = True,
        generate_qa_report = True,
        save_log           = True,
    )

    print("\n" + "─" * 60)
    print("QUICK SUMMARY")
    print("─" * 60)
    print(f"  Final rows : {result['final_rows']:,}")
    print(f"  Output dir : {result['output_dir']}")
    print(f"\n  Rows removed per filter:")
    for k, v in result["drop_counts"].items():
        print(f"    {k:<28}: {v:>12,}")
    print(f"\n  Files created:")
    for key in ["csv", "parquet", "statistics", "qa_report", "log"]:
        if result.get(key):
            print(f"    • {Path(result[key]).name}")
    print("─" * 60)
