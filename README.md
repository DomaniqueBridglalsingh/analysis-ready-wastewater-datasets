<div align="center">
  <img src="logo.png" height="90" align="middle" alt="HydroStream Logo">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <img src="exeter_uni.png" height="80" align="middle" alt="University of Exeter Logo">

  # HydroStream V2

  **Environment Agency legacy and current-API water-quality pipeline**

  ![Python](https://img.shields.io/badge/Python-%3E%3D3.11-blue)
  ![Version](https://img.shields.io/badge/version-2.0.0-green)
  ![License](https://img.shields.io/badge/license-CC--BY--4.0-lightgrey)
</div>

> **Authors:** Domanique Bridglalsingh, Ahmed Abdalla, Jia Hu, Geyong Min, Xiaohong Li, and Siwei Zheng  
> **Websites:** [HydroStar](http://www.hydrostar-eu.com) | [University of Exeter](https://www.exeter.ac.uk)

---

## Purpose

HydroStream V2 constructs reproducible, analysis-ready water-quality datasets from two Environment Agency (EA) source structures:

1. Legacy annual CSV exports covering the historical archive.
2. Current Water Quality Explorer/API CSV exports used from 13 October 2025 onward.

The pipeline validates every input, adapts both schemas to one internal structure, applies the same documented cleaning rules, and writes a concise scientific dataset together with separate QA and provenance outputs. Raw files are read only and are never rewritten.

HydroStream does not replace the EA Water Quality Explorer. The Explorer provides the source observations; HydroStream provides a transparent and repeatable workflow for integrating, filtering, harmonising and auditing them.

## Versions in this repository

| File | Version | Source coverage |
| :--- | :---: | :--- |
| `hydrostream.py` / `hydrostream.ipynb` | V1.0.0 | Legacy annual EA CSV files |
| `hydrostream_v2.py` / `hydrostream_v2.ipynb` | V2.0.0 | Legacy annual files plus current API-format CSV files |

The `v1.0.0` Git tag remains the fixed V1 release. V2 development does not alter that tag.

## Validated V2 source boundary

The default transition is:

| Source | Dates used |
| :--- | :--- |
| Legacy annual exports | Strictly before `2025-10-13` |
| Current API-format exports | `2025-10-13` onward |

The transition day was checked independently: all 490 legacy observations dated 13 October 2025 were recoverable in the current source using sampling-point code, sample ID and determinand code. V2 therefore uses the current source for the complete cutover day.

The frozen V2.0.0 data snapshot used during validation ends on **31 July 2026**. Fresh EA downloads may differ because the official archive is updated and corrected over time.

## Project layout

Create one project folder and place the raw data in two separate directories:

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
├── hydrostream_v2.py
├── hydrostream_v2.ipynb
└── requirements-hydrostream-v2.txt
```

`EA_processed_output_v2/` is created automatically after a successful run.

### Input rules

- Legacy files must use their calendar-year names, such as `2000.csv` and `2025.csv`.
- `api_raw/` may contain one or more `.csv` files. Their filenames can describe any sensible date or geographic split.
- Every file must satisfy one complete, unambiguous source-schema contract. Unknown, incomplete, hybrid or duplicate-column schemas fail before processing.
- Avoid overlapping API downloads where possible. Equivalent repeated EA source records are removed deterministically; conflicting copies cause the run to stop.
- For a full historical run, use the frozen legacy annual files plus API-format data from the cutover date onward. Do not mix an unvalidated all-history API download into the V2.0.0 release snapshot.

## Obtaining source data

The current EA archive is available through:

- [Water Quality Explorer](https://environment.data.gov.uk/water-quality)
- [Public API documentation](https://environment.data.gov.uk/water-quality/api-docs)
- [EA guidance for the current API](https://environment.data.gov.uk/support/faqs/275879249/1156874241)

HydroStream V2 does not download source data automatically. Download the required CSV files, keep an unchanged copy, and record the download date. To reproduce the published release exactly, use the frozen raw snapshot associated with that release rather than a later live download.

If you already have the V1 `RAW_DATA_FOLDER.zip`, extract its annual `2000.csv` to `2025.csv` files into `legacy_raw/`, then place current-format post-cutover downloads in `api_raw/`.

## Installation

Python 3.11 or 3.12 is recommended. From the repository or project directory, run:

```bash
python -m pip install -r requirements-hydrostream-v2.txt
```

HydroStream never installs or changes packages while it is running.

## Running from Python

```python
from hydrostream_v2 import hydrostream

result = hydrostream(
    input_dir="/path/to/HydroStream_V2",
    mode="full",
    categories_file=None,          # auto-detect the standard workbook name
    years=range(2000, 2027),       # 2000 through 2026 inclusive
    chunksize=250_000,
    min_test_count=50,
    generate_stats=True,
    generate_qa_report=True,
    save_log=True,
    cutover_date="2025-10-13",
    duckdb_memory_limit="6GB",
    finalizer="duckdb",
)

print(result["csv"])
print(f'{result["final_rows"]:,} rows')
```

The accompanying notebook follows the V1 two-cell format: the complete V2 implementation is in the first cell, and the editable settings and run block are in the second.

## Output modes

| Mode | Content |
| :--- | :--- |
| `full` | Retained quantitative water-matrix tests. Rare tests can be removed using `min_test_count`. |
| `electrochemistry` | Selected dissolved metals, ions, pH, conductivity, temperature, turbidity and related tests. |
| `contaminants` | Tests assigned to the reviewed contaminants category. A valid categories workbook is required. |

## Primary dataset schema

The main CSV and Parquet files contain only:

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
Category (when a categories workbook is used)
```

`RecordID`, `SamplingPointCode` and `DeterminandCode` remain available in the standalone provenance output but are not included in the primary scientific dataset.

### Qualified results

Current API results can contain reporting qualifiers such as `<3` or `>10`. V2 separates the qualifier while retaining a numeric result:

| Source value | `result` | `ResultQualifier` |
| :---: | ---: | :---: |
| `<3` | `3.0` | `<` |
| `>=10` | `10.0` | `>=` |
| `7.2` | `7.2` | blank |

Always interpret `result` together with `ResultQualifier`. A qualified value is a reporting bound, not an exact measurement. Statistics produced by the optional workbook summarise the stored numeric values and do not apply a censored-data model.

## Unit policy

Reviewed dimensional conversions continue to adjust both the value and unit:

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

The following units are deliberately **not** treated as interchangeable:

- `ppm` is not converted to `mg/l`.
- `FTU` is not converted to `NTU`.
- `g/kg`, `PSU`, `‰` and `ppt` remain distinct.

Current API verbose labels may still be abbreviated without changing their physical meaning, for example `NEPHELOMETRIC TURBIDITY UNITS` to `NTU`. The exact source label is retained as `raw_unit` in the provenance output.

## Cleaning and validation

V2 applies the following stages:

1. Hash and schema-check every source file.
2. Apply the validated legacy/API cutover.
3. Parse dates and restrict observations to the requested years.
4. Remove known dummy legacy coordinates and convert valid British National Grid coordinates to WGS 84.
5. Keep the documented water and wastewater material types.
6. Remove non-quantitative units and unreviewed current-API `UNITLESS VALUE` records.
7. Remove administrative or procedural tests.
8. Apply the selected mode and optional rare-test threshold.
9. Parse numeric and qualified results.
10. Apply reviewed unit-label and numeric conversions.
11. Verify source identities and remove only equivalent copies of the same EA source record.
12. Write all requested outputs transactionally and confirm that the raw inputs remained unchanged.

V2.0.0 performs no outlier flagging or outlier removal and produces no outlier columns.

## Generated outputs

All outputs are written to `EA_processed_output_v2/`.

| Output | Purpose |
| :--- | :--- |
| `EA_clean_..._v2.csv` | Primary analysis-ready dataset. |
| `EA_clean_..._v2.parquet` | Compressed equivalent of the primary dataset. |
| `EA_provenance_..._v2.csv` | Standalone scientific rows plus EA/source identifiers and unit provenance. |
| `EA_stations_v2.csv` | Deterministically selected station metadata. |
| `EA_metadata_v2.csv` | Primary output data dictionary. |
| `EA_unit_crosswalk_v2.csv` | Raw and canonical unit labels with review status. |
| `EA_schema_crosswalk_v2.csv` | Legacy/API field mapping. |
| `EA_source_summary_v2.csv` | Per-source row totals. |
| `EA_source_manifest_v2.csv` | File inventory, hashes, schema and date coverage. |
| `EA_source_identity_dedup_v2.csv` | Repeated source-identity audit. |
| `EA_raw_integrity_v2.csv` | Before/after raw-file integrity check. |
| `EA_runtime_provenance_v2.json` | Function, runtime, package and argument provenance. |
| `EA_statistics_..._v2.xlsx` | Optional descriptive statistics workbook. |
| `EA_qa_report_..._v2.html` | Optional human-readable QA report. |
| `EA_processing_log_..._v2.txt` | Optional processing log. |

## Performance

- DuckDB is the production finaliser for archive-scale runs and can spill intermediate work to disk.
- The pandas finaliser is restricted to bounded runs below the configured raw-row safety limit.
- The default chunk size is 250,000 rows.
- A full run requires substantial free disk space for staging, final outputs and DuckDB temporary files. V2 performs a conservative capacity check before processing.
- On a 16 GB computer, start with the default 6 GB DuckDB limit. Use a fast local drive and ensure that the temporary volume has ample free space.

## Reproducibility

The EA archive is a live dataset, so later downloads can include additions or corrections. A reproducible release should preserve:

- the frozen raw files;
- download dates;
- SHA-256 hashes from `EA_source_manifest_v2.csv`;
- the categories workbook and its hash;
- the tagged `hydrostream_v2.py` version;
- `EA_runtime_provenance_v2.json`;
- the exact output files used in the analysis or publication.

## Data licence and citation

EA source observations are provided under the [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/) and should retain the Environment Agency attribution supplied with the source dataset.

When using HydroStream, cite the tagged GitHub release, its Zenodo DOI, and the associated publication. Do not cite a moving `main` branch as the only software reference.
