<div align="center">
  <img src="logo.png" height="90" align="middle" alt="HydroStream Logo">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <img src="exeter_uni.png" height="80" align="middle" alt="University of Exeter Logo">

  # HydroStream V2.1

  **Environment Agency legacy and current-API water-quality pipeline**

  ![Python](https://img.shields.io/badge/Python-%3E%3D3.11-blue)
  ![Version](https://img.shields.io/badge/version-2.1.0-green)
  ![License](https://img.shields.io/badge/license-CC--BY--4.0-lightgrey)
  [![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.21935180.svg)](https://doi.org/10.5281/zenodo.21935180)
</div>

> **Authors:** Domanique Bridglalsingh, Ahmed Abdalla, Jia Hu, Geyong Min, Xiaohong Li, and Siwei Zheng  
> **Websites:** [HydroStar](http://www.hydrostar-eu.com) | [University of Exeter](https://www.exeter.ac.uk)

---

## Purpose

HydroStream V2.1 constructs reproducible, analysis-ready water-quality datasets from two Environment Agency (EA) source structures:

1. Legacy annual CSV exports covering the historical archive.
2. Current Water Quality Explorer/API CSV exports used from 13 October 2025 onward.

The pipeline validates every input, adapts both schemas to one internal structure, applies the same documented cleaning rules, and writes a concise scientific dataset with optional statistics, QA, and processing-log outputs. Raw files are read only and are never rewritten.

HydroStream does not replace the EA Water Quality Explorer. The Explorer provides the source observations; HydroStream provides a transparent and repeatable workflow for integrating, filtering, harmonising, and auditing them.

## Version 2.1.0

V2.1.0 is a reliability and data-quality release. It:

- adds the boolean, unit-aware, qualifier-aware `PlausibilityFlag` while retaining every observation;
- avoids the Jupyter `NameError: __file__ is not defined` failure that could occur after a long notebook run;
- writes the optional processing log continuously, so a partial log remains inside the run workspace if processing fails;
- restores existing published outputs transactionally and retains the failed run workspace for diagnosis instead of deleting hours of staged work;
- prints the conservative peak-disk estimate and explains the staging, output, and DuckDB temporary-space components;
- retains the simplified V2 output contract: clean CSV and Parquet, plus only the requested statistics workbook, QA report, and processing log;
- keeps the validated source cutover, scientific cleaning rules, source-identity deduplication, unit policy, dependency bounds, and `_v2` output suffixes.

The `v1.0.0` and `v2.0.0` Git tags remain fixed releases. V2.0.0 is archived at [Zenodo DOI 10.5281/zenodo.21992969](https://doi.org/10.5281/zenodo.21992969). The all-versions DOI is [10.5281/zenodo.21935180](https://doi.org/10.5281/zenodo.21935180).

## Validated source boundary

| Source | Dates used |
| :--- | :--- |
| Legacy annual exports | Strictly before `2025-10-13` |
| Current API-format exports | `2025-10-13` onward |

The transition day was checked independently: all 490 legacy observations dated 13 October 2025 were recoverable in the current source using sampling-point code, sample ID, and determinand code. HydroStream therefore uses the current source for the complete cutover day.

The frozen validation snapshot ends on **31 July 2026**. Fresh EA downloads may differ because the official archive is updated and corrected over time.

## Project layout

Create one project folder and keep the two source formats separate:

```text
HydroStream_V2/
├── legacy_raw/
│   ├── 2000.csv
│   ├── 2001.csv
│   ├── ...
│   └── 2025.csv
├── api_raw/
│   ├── EA_2025-10-13_2025-12-30.csv
│   └── EA_2026-01-01_to_2026-07-31.csv
├── List of tests kept and categories.xlsx
├── hydrostream_v2.1.py
├── hydrostream_v2.1.ipynb
├── CITATION.cff
└── requirements-hydrostream-v2.txt
```

`EA_processed_output_v2/` is created automatically after preflight succeeds.

### Input rules

- Legacy files must use their calendar-year names, such as `2000.csv` and `2025.csv`.
- `api_raw/` may contain one or more `.csv` files. Filenames may describe any sensible date or geographic split.
- Every file must satisfy one complete, unambiguous source-schema contract. Unknown, incomplete, hybrid, or duplicate-column schemas fail before processing.
- Avoid overlapping API downloads where possible. Equivalent copies of the same EA source record are removed deterministically; conflicting copies stop the run.
- Do not mix an unvalidated all-history API download into the frozen legacy-plus-current release snapshot.

## Obtaining source data

The current EA archive is available through:

- [Water Quality Explorer](https://environment.data.gov.uk/water-quality)
- [Public API documentation](https://environment.data.gov.uk/water-quality/api-docs)
- [EA current-API guidance](https://environment.data.gov.uk/support/faqs/275879249/1156874241)

HydroStream does not download source data automatically. Keep unchanged copies of every downloaded CSV and record the download date. To reproduce a published release exactly, use its frozen raw snapshot rather than a later live download.

## Installation

Python 3.11 or 3.12 is recommended:

```bash
python -m pip install -r requirements-hydrostream-v2.txt
```

HydroStream never installs or changes packages while it is running. V2.1.0 uses the same dependency ranges as V2.0.0.

## Run the notebook

Open `hydrostream_v2.1.ipynb` and run its two code cells in order:

1. The first cell defines the complete production pipeline.
2. The second cell contains the editable project path and run settings.

The release notebook contains no personal paths and no one-off recovery cell. Its default project root is the notebook directory (`.`).

## Run from the command line

```bash
python hydrostream_v2.1.py \
  --input-dir /path/to/HydroStream_V2 \
  --mode full \
  --start-year 2000 \
  --end-year 2026 \
  --finalizer duckdb \
  --duckdb-memory-limit 6GB
```

Optional switches include `--categories-file`, `--chunksize`, `--min-test-count`, `--temp-dir`, `--no-stats`, `--no-qa`, and `--no-log`. Run `python hydrostream_v2.1.py --help` for the complete interface.

## Call the function from Python

The requested release filename contains a full stop before `1`, so it is not a valid name in a normal `from ... import ...` statement. Load it safely with the standard library:

```python
from runpy import run_path

hydrostream = run_path(
    "hydrostream_v2.1.py",
    run_name="hydrostream_v2_1",
)["hydrostream"]

result = hydrostream(
    input_dir="/path/to/HydroStream_V2",
    mode="full",
    categories_file=None,
    years=range(2000, 2027),
    chunksize=250_000,
    min_test_count=50,
    generate_stats=True,
    generate_qa_report=True,
    save_log=True,
    cutover_date="2025-10-13",
    duckdb_memory_limit="6GB",
    finalizer="duckdb",
    temp_dir=None,
)

print(result["csv"])
print(f'{result["final_rows"]:,} rows')
print(
    f'{result["data_quality"]["records_with_plausibility_flag"]:,} '
    "plausibility flags"
)
```

## Output modes

| Mode | Content |
| :--- | :--- |
| `full` | Retained quantitative water-matrix tests. Rare tests can be removed with `min_test_count`. |
| `electrochemistry` | Selected dissolved metals, ions, pH, conductivity, temperature, turbidity, and related tests. |
| `contaminants` | Tests assigned to the reviewed contaminants category. A valid categories workbook is required. |

## Primary dataset schema

The main CSV and Parquet files contain, in this order:

```text
Sampling Point
Type
Date
Test
result
ResultQualifier
Unit
Season
SourceYear
Latitude
Longitude
PlausibilityFlag
Category (when a categories workbook is used)
```

`RecordID`, `SamplingPointCode`, and `DeterminandCode` are used internally for validation, ordering, and source-identity checks, then omitted from the user-facing dataset.

### Qualified results

Current API results can contain reporting qualifiers such as `<3` or `>10`. HydroStream separates the qualifier while retaining a numeric result:

| Source value | `result` | `ResultQualifier` |
| :---: | ---: | :---: |
| `<3` | `3.0` | `<` |
| `>=10` | `10.0` | `>=` |
| `7.2` | `7.2` | blank |

Always interpret `result` together with `ResultQualifier`. A qualified value is a reporting bound, not an exact measurement. The optional statistics workbook summarises stored numeric values and does not apply a censored-data model.

## Plausibility flag

`PlausibilityFlag` is produced in every mode. It is `True` only when the exact test and canonical unit match a reviewed rule and the result is definitively outside the valid interval.

| Test | Required unit | `True` when |
| :--- | :---: | :--- |
| `Temperature of Water` | `cel` | value is `< -5` or `> 50` |
| `pH` | `phunits` | value is `< 1` or `> 14` |
| `Conductivity at 20 C` | `uS/cm` | value is `< 0` |
| `Conductivity at 25 C` | `uS/cm` | value is `< 0` |
| `Conductivity : In Situ` / `Conductivity: In Situ` | `uS/cm` | value is `< 0` |

Rules:

- boundaries are valid (`-5`, `50`, `1`, `14`, and `0` are not flagged when exact);
- conductivity has no upper limit;
- qualifier semantics are conservative: for example, conductivity `<0` is flagged but `<=0` is not, and temperature `>50` is flagged but `>=50` is not;
- an indeterminate reporting bound is not flagged;
- a test with the wrong unit is not flagged;
- flagged rows remain in the dataset unchanged.

The optional statistics workbook includes a `Plausibility_Flags` sheet, and the QA report and processing log include flag counts.

## Unit policy

Reviewed dimensional conversions adjust both the numeric value and unit:

| From | To | Value operation |
| :--- | :--- | :--- |
| `ug/l` | `mg/l` | divide by 1,000 |
| `ng/l` | `mg/l` | divide by 1,000,000 |
| `pg/l` | `mg/l` | divide by 1,000,000,000 |
| `g/l` | `mg/l` | multiply by 1,000 |
| `mS/cm` | `uS/cm` | multiply by 1,000 |
| `no/ml` | `no/100ml` | multiply by 100 |
| `no/ul` | `no/100ml` | multiply by 100,000 |
| `no/10ul` | `no/100ml` | multiply by 10,000 |

The following units are deliberately not treated as interchangeable:

- `ppm` is not converted to `mg/l`;
- `FTU` is not converted to `NTU`;
- `g/kg`, `PSU`, `‰`, and `ppt` remain distinct.

Current API verbose labels may be abbreviated without changing physical meaning, for example `NEPHELOMETRIC TURBIDITY UNITS` to `NTU`. Current-API `UNITLESS VALUE` records are rejected unless their `(DeterminandCode, Test)` context is explicitly reviewed and supplied through `unitless_quantitative_allowlist`.

## Cleaning and validation

HydroStream applies these stages:

1. Hash and schema-check every source file.
2. Apply the validated legacy/API cutover.
3. Parse dates and restrict observations to requested years.
4. Remove known dummy legacy coordinates and convert valid British National Grid coordinates to WGS 84.
5. Keep the documented water and wastewater material types.
6. Remove non-quantitative units and unreviewed current-API `UNITLESS VALUE` records.
7. Remove administrative or procedural tests.
8. Apply the selected mode and optional rare-test threshold.
9. Parse numeric and qualified results.
10. Apply reviewed unit-label and numeric conversions.
11. Add `PlausibilityFlag` without removing observations.
12. Verify source identities and remove only equivalent copies of the same EA source record.
13. Confirm raw inputs remained unchanged and publish outputs transactionally.

## Generated outputs

Outputs are written to `EA_processed_output_v2/`.

| Output | Purpose |
| :--- | :--- |
| `EA_clean_..._v2.csv` | Primary analysis-ready dataset. |
| `EA_clean_..._v2.parquet` | Compressed equivalent of the primary dataset. |
| `EA_statistics_..._v2.xlsx` | Optional descriptive statistics and audit workbook. |
| `EA_qa_report_..._v2.html` | Optional human-readable QA report. |
| `EA_processing_log_..._v2.txt` | Optional continuously written processing log. |

CSV and Parquet are always produced. The other three files are controlled by `generate_stats`, `generate_qa_report`, and `save_log`. Setting all three options to `False` leaves only the two clean datasets.

Schema, unit, source-file, hash, raw-integrity, and deduplication checks still run. Their concise audit tables are included in the optional statistics workbook and processing log rather than published as separate sidecar files.

## Memory and disk requirements

`chunksize` limits the number of raw rows held in memory. It does **not** reduce the total temporary disk needed by a full archive run.

At peak, an out-of-core run may hold these files concurrently:

- the cleaned streaming staging CSV;
- the DuckDB working database and spill files;
- the final public CSV;
- the compressed Parquet output;
- the optional workbook, QA report, and log;
- short-lived publication backups when replacing an earlier output set.

That is why required free space can be several times larger than the raw CSV download size. The preflight estimate is deliberately conservative and runs before archive processing starts. The exact estimate and available capacity are printed in the run header and returned in `result["capacity_preflight"]`.

For a 16 GB computer, start with a 6 GB DuckDB memory limit. Use a fast local drive. If another volume has more capacity, pass its existing directory through `temp_dir`; this moves DuckDB temporary spill files, while the staging and final publication files remain with the project so publication can stay atomic.

## Failed runs and retained workspaces

The active workspace is removed automatically only after every requested output is published successfully. If a run fails, look inside:

```text
EA_processed_output_v2/.hydrostream-v2-run-*/
```

When `save_log=True`, that directory contains the processing log written up to the failure. V2.1 also warns about older incomplete workspaces at the next start because they can consume substantial storage.

Do not delete an incomplete workspace until you have inspected the log and confirmed that no staged work is needed. If you intentionally start again from raw inputs, removing an obsolete failed workspace first can recover considerable disk capacity.

## Reproducibility

The EA archive is live, so later downloads can contain additions or corrections. A reproducible release should preserve:

- frozen raw files and download dates;
- SHA-256 hashes recorded in the optional statistics workbook or processing log;
- the categories workbook and its hash;
- the tagged `hydrostream_v2.1.py` version;
- the run environment recorded in the optional statistics workbook or processing log;
- the exact output files used in analysis or publication.

## Citation, data licence, and software licence

When using HydroStream, cite the tagged GitHub release, its version-specific Zenodo DOI, the [all-versions Zenodo DOI](https://doi.org/10.5281/zenodo.21935180), and the associated publication. Do not cite a moving `main` branch as the only software reference.

`CITATION.cff` records the V2.1.0 title, release date, authors, repository, licence, and keywords for GitHub and Zenodo. The version-specific V2.1.0 DOI is assigned when the new Zenodo version is published; the all-versions DOI remains unchanged.

EA source observations are provided under the [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/) and should retain the Environment Agency attribution supplied with the source dataset.

HydroStream repository materials are released under CC BY 4.0, as stated for the existing project release.
