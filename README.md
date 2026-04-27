<div align="center">
  <img src="logo.png" width="120" alt="HydroStream Logo">

  # HydroStream
  **EA Water Quality Archive Pipeline**

  ![Python](https://img.shields.io/badge/Python-%3E%3D3.9-blue)
  ![Version](https://img.shields.io/badge/version-1.0.0-green)
  ![License](https://img.shields.io/badge/license-CC--BY--4.0-lightgrey)

  *Environment Agency (England) Open Water Quality Archive — Processor*
</div>

> **Authors:** Domanique Bridglalsingh, Ahmed Abdalla, Jia Hu, Geyong Min, Xiaohong Li, and Siwei Zheng  
> **Website:** [www.hydrostar-eu.com](http://www.hydrostar-eu.com)

---

## Purpose

The Environment Agency (EA) published annual CSV files of water-quality measurements from 2000 to 2025. These files are no longer publicly hosted. 

HydroStream reads the raw yearly CSVs, applies **transparent and reproducible cleaning**, and produces a single **analysis-ready dataset** ready for analysis, modelling, and industrial use.

---
## Repository Layout

HydroStream is kept intentionally simple: the main code and support files live in the repository root, and the raw EA input files are packaged separately inside `RAW_DATA_FOLDER.zip`.

```text
Repository Root/
├── hydrostream.py
├── hydrostream.ipynb
├── README.md
├── logo.png
├── List of tests kept and categories.xlsx
└── RAW_DATA_FOLDER.zip
    └── RAW_DATA_FOLDER/
        ├── 2000.csv
        ├── 2001.csv
        ├── ...
        ├── 2025.csv
        └── List of tests kept and categories.xlsx
```

### What this means

- `hydrostream.py` contains the main HydroStream processing function.
- `hydrostream.ipynb` provides an interactive notebook version for running the workflow.
- `RAW_DATA_FOLDER.zip` packages the required raw input folder expected by the pipeline.
- After extraction, `RAW_DATA_FOLDER/` should contain the annual EA CSV files from 2000 to 2025 plus `List of tests kept and categories.xlsx`.

---

## Dependencies

The script automatically installs missing Python packages. Core libraries include:

<p>
  <img src="https://img.shields.io/badge/pandas-150458?logo=pandas&logoColor=white" alt="pandas">
  <img src="https://img.shields.io/badge/numpy-013243?logo=numpy&logoColor=white" alt="numpy">
  <img src="https://img.shields.io/badge/pyarrow-2C3E50" alt="pyarrow">
  <img src="https://img.shields.io/badge/pyproj-4B8BBE" alt="pyproj">
  <img src="https://img.shields.io/badge/openpyxl-217346" alt="openpyxl">
  <img src="https://img.shields.io/badge/chardet-666666" alt="chardet">
</p>

---

## Output Modes

HydroStream supports three built-in processing modes:

* **`full`**: All water-related tests, types, and sampling points.
* **`electrochemistry`**: Dissolved metals, ions, pH, conductivity, temperature, and turbidity.
* **`contaminants`**: Microplastics, PFAS, pesticides, and insecticides *(requires categories file)*.

---

## How to Use

**1. Directory Setup** Prepare your working directory with the raw data folder:

    Working Directory/
    └── RAW_DATA_FOLDER/
        ├── 2000.csv
        ├── 2001.csv
        ├── ...
        ├── 2025.csv
        └── List of tests kept and categories.xlsx

**2. Execution** Create a Jupyter Notebook or download `hydrostream.ipynb` in the same directory as the `RAW_DATA_FOLDER` (which contains both the raw CSV files and the Excel categories file), then run:

    # ============================================================================
    # USAGE
    # ============================================================================
    
    from pathlib import Path
    
    if __name__ == "__main__":
    
        # ── SETTINGS (EDIT THESE) ─────────────────────────────────────────
        RAW_DATA_FOLDER = "./RAW_DATA_FOLDER"   # folder with CSVs + Excel file
        MODE = "full"                           # "full", "electrochemistry", or "contaminants"
        # ──────────────────────────────────────────────────────────────────
    
        result = hydrostream(
            input_dir=RAW_DATA_FOLDER,
            mode=MODE,
            categories_file=None,   # auto-detected from input_dir
            years=range(2000, 2026),
            chunksize=250_000,
            min_test_count=50,
            flag_outliers=True,
            generate_stats=True,
            generate_qa_report=True,
            save_log=True,
        )
    
        print("\n" + "─" * 60)
        print("QUICK SUMMARY")
        print("─" * 60)
        print(f"  Final rows : {result['final_rows']:,}")
        print(f"  Output dir : {result['output_dir']}")
    
        print("\n  Rows removed per filter:")
        for k, v in result["drop_counts"].items():
            print(f"    {k:<28}: {v:>12,}")
    
        print("\n  Files created:")
        for key in ["csv", "parquet", "statistics", "qa_report", "log"]:
            if result.get(key):
                print(f"    • {Path(result[key]).name}")
    
        print("─" * 60)
**3. Results** After running the function, a new folder called `EA_processed_output/` will be created automatically in the same working directory containing your processed files.

---

## Outputs

All outputs are saved in the `EA_processed_output/` directory.

| Output Type | Format | Description |
| :--- | :---: | :--- |
| **Clean Dataset** | `.csv`, `.parquet` | Final analysis-ready dataset. |
| **Statistics** | `.xlsx` | Summary processing statistics. |
| **QA Report** | `.html` | Quality Assurance report. |
| **Processing Log**| `.txt` | Detailed execution log. |

---

## Cleaning Logic

Rows are processed systematically. Removed rows are tracked for full auditability.

* Dummy coordinate removal
* Water-type filtering
* Unit filtering *(non-quantitative removed)*
* Administrative test removal
* Rare test filtering *(optional)*
* Numeric conversion
* Unit standardisation
* Outlier flagging *(flagged, not removed)*

---

## Performance

* **Scalable:** Processes ~70M rows using chunked streaming.
* **Memory Efficient:** Avoids full memory load during cleaning.
* **Optimized:** Requires only a single, final aggregation step.

---

## Output Dataset Schema

**Columns:**
Sampling Point · Type · Date · Test · result · Unit · Season · SourceYear · Latitude · Longitude · outlier_flag · Category *(optional)*
