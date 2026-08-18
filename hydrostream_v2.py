"""
===============================================================================
HYDROSTREAM V2 — EA WATER QUALITY LEGACY + CURRENT API PIPELINE
===============================================================================
Version : 2.0.0
Python  : 3.12.x  (focused release suite tested on 3.12.3)

HydroStream V2 supports BOTH Environment Agency source structures:

1) Legacy annual CSV exports in ``legacy_raw/`` (2000.csv ... 2025.csv)
2) Current EA API CSV exports in ``api_raw/`` (late-2025 and 2026 onward)

The raw files are never rewritten. Each source format is adapted in memory to
one canonical structure, then the same cleaning/QA pipeline is applied.

Validated default cutover
-------------------------
Legacy source : strictly before 2025-10-13
Current API   : 2025-10-13 onward

The overlap date was checked separately: all 490 legacy observations on
2025-10-13 were recoverable in the new API using sampling-point code + sample
ID + determinand code. Therefore V2 uses the new API for the whole cutover day.

Qualified results
-----------------
The new API may encode detection-limit qualifiers inside ``result`` (e.g. <3).
V2 separates these into:

    result          = 3
    ResultQualifier = <

Legacy ``resultQualifier.notation`` is preserved in the same field.

Expected project layout
-----------------------
HydroStream_V2/
├── legacy_raw/
│   ├── 2000.csv
│   ├── ...
│   └── 2025.csv
├── api_raw/
│   ├── EA_2025-10-13_2025-12-30.csv
│   └── EA_2026-01-01_to_2026-07-31.csv
├── List of tests kept and categories.xlsx
└── hydrostream_v2.py

Notes
-----
* DuckDB is the out-of-core production finaliser. A pandas finaliser remains
  available only for explicitly bounded small runs and focused tests.
* Dependencies are never installed or changed by this module. Install the
  release requirements from ``requirements-hydrostream-v2.txt`` before running.
* ``mode="full"`` means the complete retained HydroStream water-matrix scope;
  it deliberately excludes biological, sediment and other non-water matrices.
===============================================================================
"""

from __future__ import annotations

# ============================================================================
# DEPENDENCIES / IMPORTS
# ============================================================================

def _try_optional_package(import_name: str, pip_name: str):
    """Import an optional package without mutating the Python environment."""
    import importlib
    import importlib.util

    if importlib.util.find_spec(import_name) is None:
        return None
    return importlib.import_module(import_name)

from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Collection, Dict, Iterable, Optional, Tuple, Union
import csv
import hashlib
import html
import io
import json
import os
import platform
import re
import shutil
import sys
import tempfile
import unicodedata

try:
    import chardet
    import numpy as np
    import openpyxl
    import pandas as pd
    import pyarrow
    from pyproj import Transformer
except ModuleNotFoundError as exc:
    raise ImportError(
        "HydroStream V2 is missing a required dependency "
        f"({exc.name!r}). Install the release requirements with "
        "`python -m pip install -r requirements-hydrostream-v2.txt`. "
        "HydroStream never installs packages at import or run time."
    ) from exc


# ============================================================================
# FAIL-CLOSED CONTRACTS AND ORCHESTRATION HELPERS
# ============================================================================

LEGACY_REQUIRED_FIELDS = frozenset({
    "@id",
    "sample.samplingPoint.notation",
    "sample.samplingPoint.label",
    "sample.sampleDateTime",
    "sample.sampledMaterialType.label",
    "determinand.notation",
    "determinand.definition",
    "result",
    "resultQualifier.notation",
    "determinand.unit.label",
    "sample.samplingPoint.easting",
    "sample.samplingPoint.northing",
})

LEGACY_OPTIONAL_FIELDS = frozenset({
    "sample.samplingPoint",
    "determinand.label",
    "codedResultInterpretation.interpretation",
    "sample.isComplianceSample",
    "sample.purpose.label",
})

API_REQUIRED_FIELDS = frozenset({
    "id",
    "samplingPoint.notation",
    "samplingPoint.prefLabel",
    "samplingPoint.longitude",
    "samplingPoint.latitude",
    "phenomenonTime",
    "sampleMaterialType",
    "determinand.notation",
    "determinand.prefLabel",
    "result",
    "unit",
})

API_OPTIONAL_FIELDS = frozenset({
    "samplingPoint.region",
    "samplingPoint.area",
    "samplingPoint.subArea",
    "samplingPoint.samplingPointStatus",
    "samplingPoint.samplingPointType",
    "samplingPurpose",
})

HYDROSTREAM_VERSION = "2.0.0"
UNITLESS_QUANTITATIVE_ALLOWLIST: frozenset[Tuple[str, str]] = frozenset()
PANDAS_FALLBACK_MAX_ROWS = 2_000_000
LEGACY_ARCHIVE_LAST_YEAR = 2025

FINAL_ORDER_COLUMNS = [
    "Date", "SamplingPointCode", "SampleID", "DeterminandCode",
    "SourceFormat", "SourceRecordID",
]

IDENTITY_EQUIVALENCE_FIELDS = [
    "SamplingPointCode", "Sampling Point", "Type", "Date",
    "DeterminandCode", "Test", "RawResult", "result", "ResultQualifier",
    "RawUnit", "Unit", "Easting", "Northing", "Latitude", "Longitude",
    "Region", "Area", "SubArea", "SamplingPointStatus",
    "SamplingPointType", "SamplingPurpose", "SampleID",
    "LegacySamplingPointURI", "LegacyDeterminandLabel",
    "LegacyCodedResultInterpretation", "LegacyIsComplianceSample",
]


class SchemaValidationError(ValueError):
    """Raised when a source header does not meet one unambiguous contract."""


class SourceValidationError(ValueError):
    """Raised when the requested source inventory is incomplete or malformed."""

    def __init__(self, message: str, manifest: Optional[pd.DataFrame] = None):
        super().__init__(message)
        self.manifest = manifest


class SourceIdentityConflictError(ValueError):
    """Raised when one EA source identity has conflicting canonical payloads."""


def _detect_schema(columns: Iterable[str]) -> str:
    """Return the one complete schema contract satisfied by ``columns``."""
    cols_list = [str(column) for column in columns]
    cols = set(cols_list)
    duplicate_columns = sorted({c for c in cols_list if cols_list.count(c) > 1})
    if duplicate_columns:
        raise SchemaValidationError(
            "Duplicate CSV columns are not allowed: " + ", ".join(duplicate_columns)
        )

    legacy_ok = LEGACY_REQUIRED_FIELDS.issubset(cols)
    api_ok = API_REQUIRED_FIELDS.issubset(cols)
    if legacy_ok and api_ok:
        raise SchemaValidationError(
            "Ambiguous hybrid EA schema satisfies both complete legacy and API contracts; "
            "conflicting marker columns include @id and id"
        )
    if legacy_ok:
        return "legacy"
    if api_ok:
        return "api"

    legacy_missing = sorted(LEGACY_REQUIRED_FIELDS - cols)
    api_missing = sorted(API_REQUIRED_FIELDS - cols)
    legacy_marked = bool(cols & (LEGACY_REQUIRED_FIELDS - {"result", "determinand.notation"}))
    api_marked = bool(cols & (API_REQUIRED_FIELDS - {"result", "determinand.notation"}))
    if legacy_marked and not api_marked:
        detail = "expected legacy; missing required columns: " + ", ".join(legacy_missing)
    elif api_marked and not legacy_marked:
        detail = "expected api; missing required columns: " + ", ".join(api_missing)
    else:
        detail = (
            "unknown or conflicting schema; missing legacy columns: "
            + ", ".join(legacy_missing)
            + "; missing API columns: "
            + ", ".join(api_missing)
            + "; actual columns: "
            + ", ".join(sorted(cols))
        )
    raise SchemaValidationError(detail)


def _validate_schema(columns: Iterable[str], expected: str) -> str:
    """Validate a complete schema and its expected source-directory format."""
    schema = _detect_schema(columns)
    if schema != expected:
        raise SchemaValidationError(
            f"expected {expected} schema but complete {schema} schema was detected"
        )
    return schema


def _detect_encoding(path: Path) -> str:
    try:
        pd.read_csv(
            path,
            nrows=3,
            # ``utf-8-sig`` is identical for ordinary UTF-8 and also strips a
            # leading BOM before the physical csv.reader header validation.
            encoding="utf-8-sig",
            keep_default_na=False,
            na_values=[""],
        )
        return "utf-8-sig"
    except UnicodeDecodeError:
        with path.open("rb") as handle:
            guess = chardet.detect(handle.read(100_000)).get("encoding")
        return guess or "latin-1"


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(8 * 1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _inspect_source_file(path: Path, expected_schema: str,
                         chunksize: int) -> Dict[str, Any]:
    """Hash, validate and date-profile one raw CSV without modifying it."""
    record: Dict[str, Any] = {
        "source_format": expected_schema,
        "source_filename": path.name,
        "source_path": str(path),
        "file_size_bytes": int(path.stat().st_size),
        "sha256": _sha256_file(path),
        # The directory contract is known even if header detection fails; the
        # rejection reason below records the detected conflict/missing fields.
        "schema": expected_schema,
        "observation_min_date": None,
        "observation_max_date": None,
        "observed_years": "",
        "raw_rows": 0,
        "accepted": False,
        "selected": False,
        "duplicate_sha256_of": None,
        "reason": "",
        "encoding": None,
    }
    try:
        encoding = _detect_encoding(path)
        record["encoding"] = encoding
        # Read the physical header before pandas can mangle duplicate names
        # (for example ``result`` and ``result.1``). Duplicate required fields
        # are ambiguous source data and must fail the schema preflight.
        with path.open("r", encoding=encoding, newline="") as handle:
            columns = next(csv.reader(handle))
        try:
            record["schema"] = _validate_schema(columns, expected_schema)
        except Exception:
            # A rejected manifest entry still reports its physical data-record
            # count; scientific/date profiling remains forbidden because the
            # schema is not trustworthy.
            with path.open("r", encoding=encoding, newline="") as handle:
                reader = csv.reader(handle)
                next(reader, None)
                record["raw_rows"] = sum(1 for _ in reader)
            raise
        date_column = (
            "sample.sampleDateTime" if expected_schema == "legacy"
            else "phenomenonTime"
        )
        minimum: Optional[pd.Timestamp] = None
        maximum: Optional[pd.Timestamp] = None
        observed_years: set[int] = set()
        for chunk in pd.read_csv(
            path,
            usecols=[date_column],
            chunksize=max(1, int(chunksize)),
            dtype=str,
            low_memory=False,
            encoding=encoding,
            keep_default_na=False,
            na_values=[""],
        ):
            record["raw_rows"] += int(len(chunk))
            dates = pd.to_datetime(
                chunk[date_column], errors="coerce", format="mixed"
            )
            valid = dates.dropna()
            if valid.empty:
                continue
            chunk_min = valid.min()
            chunk_max = valid.max()
            minimum = chunk_min if minimum is None else min(minimum, chunk_min)
            maximum = chunk_max if maximum is None else max(maximum, chunk_max)
            observed_years.update(int(y) for y in valid.dt.year.unique())
        if record["raw_rows"] and minimum is None:
            raise SourceValidationError(
                "required observation date column contains no parseable dates"
            )
        record["observation_min_date"] = str(minimum) if minimum is not None else None
        record["observation_max_date"] = str(maximum) if maximum is not None else None
        record["observed_years"] = ";".join(str(y) for y in sorted(observed_years))
        record["_observed_year_set"] = observed_years
        record["accepted"] = True
        record["reason"] = "accepted"
    except Exception as exc:
        record["reason"] = f"{type(exc).__name__}: {exc}"
    return record


def _build_source_manifest(root: Path, years: Collection[int],
                           chunksize: int, cutover: pd.Timestamp
                           ) -> Tuple[list, pd.DataFrame]:
    """Build the authoritative, fail-closed source inventory for a run."""
    legacy_dir = root / "legacy_raw"
    api_dir = root / "api_raw"
    if not root.is_dir():
        raise FileNotFoundError(f"Project root is not a directory: {root}")
    if not legacy_dir.is_dir():
        raise FileNotFoundError(f"Missing required source directory: {legacy_dir}")
    if not api_dir.is_dir():
        raise FileNotFoundError(f"Missing required source directory: {api_dir}")

    year_set = {int(year) for year in years}
    records: list = []
    expected_legacy_years = sorted(
        year for year in year_set if year <= LEGACY_ARCHIVE_LAST_YEAR
    )
    for year in expected_legacy_years:
        path = legacy_dir / f"{year}.csv"
        if not path.is_file():
            records.append({
                "source_format": "legacy",
                "source_filename": path.name,
                "source_path": str(path),
                "file_size_bytes": None,
                "sha256": None,
                "schema": "legacy",
                "observation_min_date": None,
                "observation_max_date": None,
                "observed_years": "",
                "raw_rows": None,
                "accepted": False,
                "selected": False,
                "duplicate_sha256_of": None,
                "reason": f"missing requested legacy source file for {year}",
                "encoding": None,
            })
            continue
        record = _inspect_source_file(path, "legacy", chunksize)
        observed = record.pop("_observed_year_set", set())
        record["selected"] = bool(record["accepted"] and (observed & year_set))
        if record["accepted"] and not record["selected"]:
            record["accepted"] = False
            record["reason"] = (
                f"requested legacy file {path.name} has no observations in "
                f"requested years {sorted(year_set)}"
            )
        records.append(record)

    for path in sorted(api_dir.glob("*.csv"), key=lambda p: p.name.encode("utf-8")):
        record = _inspect_source_file(path, "api", chunksize)
        observed = record.pop("_observed_year_set", set())
        record["selected"] = bool(record["accepted"] and (observed & year_set))
        records.append(record)

    hashes: Dict[str, str] = {}
    for record in records:
        digest = record.get("sha256")
        if not record.get("accepted") or not digest:
            continue
        if digest in hashes:
            record["duplicate_sha256_of"] = hashes[digest]
        else:
            hashes[digest] = str(record["source_filename"])

    public_columns = [
        "source_format", "source_filename", "source_path", "file_size_bytes",
        "sha256", "schema", "observation_min_date", "observation_max_date",
        "observed_years", "raw_rows", "accepted", "selected",
        "duplicate_sha256_of", "reason", "encoding",
    ]
    manifest = pd.DataFrame(records, columns=public_columns)
    rejected = manifest.loc[~manifest["accepted"].fillna(False)]
    if not rejected.empty:
        details = "; ".join(
            f"{row.source_filename} [{row.source_format}]: {row.reason}"
            for row in rejected.itertuples()
        )
        raise SourceValidationError(
            "Source preflight rejected one or more inputs: " + details,
            manifest=manifest,
        )

    need_api = any(year >= cutover.year for year in year_set)
    selected_api = manifest[
        manifest["source_format"].eq("api") & manifest["selected"].fillna(False)
    ]
    if need_api and selected_api.empty:
        raise SourceValidationError(
            "No validated API source contains observations in the requested years",
            manifest=manifest,
        )

    sources = []
    for row in manifest.loc[manifest["selected"].fillna(False)].itertuples():
        observed = [int(value) for value in str(row.observed_years).split(";") if value]
        participating = sorted(set(observed) & year_set)
        sources.append({
            "year": participating[0] if participating else min(year_set),
            "path": Path(row.source_path),
            "format": row.source_format,
            "encoding": row.encoding,
        })
    if not sources:
        raise SourceValidationError(
            "No validated source contains observations in the requested years",
            manifest=manifest,
        )
    return sources, manifest


def _normalise_category_key(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value).strip()).casefold()


def _load_category_workbook(path: Path) -> Tuple[Dict[str, str], Dict[str, Any]]:
    """Validate category decisions, allowing only equivalent duplicate keys."""
    frame = pd.read_excel(
        path,
        keep_default_na=False,
        na_values=[""],
    )
    required = {"List of Tests", "Final Category"}
    missing = sorted(required - set(frame.columns))
    if missing:
        raise ValueError(
            f"Category workbook {path.name} is missing required columns: "
            + ", ".join(missing)
        )
    names = frame["List of Tests"].astype("string").str.strip()
    categories = frame["Final Category"].astype("string").str.strip()
    blank = names.isna() | names.eq("") | categories.isna() | categories.eq("")
    if blank.any():
        raise ValueError(
            f"Category workbook {path.name} has {int(blank.sum())} blank key/category rows"
        )
    validated = pd.DataFrame({
        "key": names.map(_normalise_category_key),
        "category": categories,
    })
    conflicts = (
        validated.groupby("key", dropna=False)["category"].nunique(dropna=False)
    )
    conflicts = conflicts[conflicts > 1]
    if not conflicts.empty:
        raise ValueError(
            f"Category workbook {path.name} has conflicting normalized keys: "
            + ", ".join(conflicts.index.astype(str).tolist()[:20])
        )
    duplicate_rows = int(validated["key"].duplicated(keep=False).sum())
    collapsed = validated.drop_duplicates("key", keep="first")
    mapping = dict(zip(collapsed["key"], collapsed["category"]))
    provenance = {
        "path": str(path),
        "sha256": _sha256_file(path),
        "file_size_bytes": int(path.stat().st_size),
        "rows": int(len(frame)),
        "normalized_keys": int(len(mapping)),
        "equivalent_duplicate_rows": duplicate_rows,
        "conflicting_keys": 0,
    }
    return mapping, provenance


def _parse_memory_limit(value: str) -> int:
    match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*([KMGT]B)\s*", str(value), re.I)
    if not match:
        raise ValueError(
            "duckdb_memory_limit must be a positive value such as 4GB or 750MB"
        )
    amount = float(match.group(1))
    if amount <= 0:
        raise ValueError("duckdb_memory_limit must be positive")
    factors = {"KB": 1024, "MB": 1024 ** 2, "GB": 1024 ** 3, "TB": 1024 ** 4}
    return int(amount * factors[match.group(2).upper()])


def _capacity_preflight(output_parent: Path, temp_parent: Path,
                        manifest: pd.DataFrame, using_duckdb: bool,
                        duckdb_memory_limit: str,
                        pandas_fallback_max_rows: int) -> Dict[str, Any]:
    """Apply a conservative, source-size-derived disk and fallback policy."""
    selected = manifest.loc[manifest["selected"].fillna(False)]
    raw_rows = int(pd.to_numeric(selected["raw_rows"], errors="coerce").fillna(0).sum())
    raw_bytes = int(
        pd.to_numeric(selected["file_size_bytes"], errors="coerce").fillna(0).sum()
    )
    if not using_duckdb and raw_rows > int(pandas_fallback_max_rows):
        raise RuntimeError(
            "DuckDB is required for this archive-scale run: "
            f"{raw_rows:,} raw rows exceed the configured pandas safety cap of "
            f"{int(pandas_fallback_max_rows):,}. Install the release DuckDB dependency "
            "or explicitly narrow the requested source coverage."
        )

    configured_memory = _parse_memory_limit(duckdb_memory_limit)
    estimated_staging = max(16 * 1024 ** 2, int(max(raw_bytes * 1.25, raw_rows * 256)))
    estimated_outputs = int(estimated_staging * 1.50)
    estimated_duck_temp = int(estimated_staging * 1.50) if using_duckdb else 0
    output_usage = shutil.disk_usage(output_parent)
    temp_usage = shutil.disk_usage(temp_parent)
    same_device = os.stat(output_parent).st_dev == os.stat(temp_parent).st_dev
    if same_device:
        required_output = estimated_staging + estimated_outputs + estimated_duck_temp
        required_temp = required_output
        available = min(output_usage.free, temp_usage.free)
        if available < required_output:
            raise RuntimeError(
                "Insufficient disk capacity for conservative HydroStream staging: "
                f"estimated {required_output:,} bytes required, {available:,} available "
                f"on {output_parent}. Configure a larger temporary volume or free space."
            )
    else:
        required_output = estimated_staging + estimated_outputs
        required_temp = estimated_duck_temp
        if output_usage.free < required_output or temp_usage.free < required_temp:
            raise RuntimeError(
                "Insufficient output/temporary capacity: "
                f"output requires {required_output:,} bytes ({output_usage.free:,} free); "
                f"DuckDB temp requires {required_temp:,} bytes ({temp_usage.free:,} free)."
            )
    return {
        "raw_rows": raw_rows,
        "raw_source_bytes": raw_bytes,
        "estimated_staging_bytes": estimated_staging,
        "estimated_output_bytes": estimated_outputs,
        "estimated_duckdb_temp_bytes": estimated_duck_temp,
        "required_output_volume_bytes": required_output,
        "required_temp_volume_bytes": required_temp,
        "output_free_bytes": int(output_usage.free),
        "temp_free_bytes": int(temp_usage.free),
        "same_filesystem": bool(same_device),
        "duckdb_memory_limit": str(duckdb_memory_limit),
        "duckdb_memory_limit_bytes": configured_memory,
        "pandas_fallback_max_rows": int(pandas_fallback_max_rows),
    }


def _module_record(name: str, module: Any) -> Dict[str, Any]:
    return {
        "name": name,
        "imported_version": str(getattr(module, "__version__", "UNKNOWN")),
        "module_path": str(getattr(module, "__file__", "UNKNOWN")),
    }


def _runtime_provenance(finalizer: str, duckdb_module: Any,
                        category_provenance: Optional[Dict[str, Any]],
                        capacity: Dict[str, Any]) -> Dict[str, Any]:
    modules = [
        _module_record("pandas", pd),
        _module_record("numpy", np),
        _module_record("pyproj", sys.modules[Transformer.__module__.split(".")[0]]),
        _module_record("openpyxl", openpyxl),
        _module_record("chardet", chardet),
        _module_record("pyarrow", pyarrow),
    ]
    if duckdb_module is not None:
        modules.append(_module_record("duckdb", duckdb_module))
    else:
        modules.append({
            "name": "duckdb",
            "imported_version": None,
            "module_path": None,
            "status": "NOT_INSTALLED",
        })
    return {
        "hydrostream_version": HYDROSTREAM_VERSION,
        "python": {
            "version": platform.python_version(),
            "implementation": platform.python_implementation(),
            "executable": sys.executable,
            "platform": sys.platform,
            "machine": os.uname().machine if hasattr(os, "uname") else "UNKNOWN",
        },
        "packages": modules,
        "finalizer": finalizer,
        "parquet_engine": "duckdb" if finalizer == "duckdb" else "pyarrow",
        "production_script": {
            "path": str(Path(__file__).resolve()),
            "sha256": _sha256_file(Path(__file__).resolve()),
        },
        "category_workbook": category_provenance,
        "capacity_preflight": capacity,
    }


def _publish_artifact(stage_path: Path, final_path: Path) -> None:
    """Monkeypatchable single-artifact atomic publication seam."""
    os.replace(stage_path, final_path)


def _publish_artifacts_transactionally(artifacts: Collection[Tuple[Path, Path]],
                                       run_dir: Path) -> None:
    """Publish a complete artifact set, rolling back on any failure."""
    pairs = list(artifacts)
    missing = [str(stage) for stage, _ in pairs if not stage.is_file()]
    if missing:
        raise RuntimeError("Mandatory staged outputs are missing: " + ", ".join(missing))
    backup_dir = run_dir / "publication_backups"
    backup_dir.mkdir(exist_ok=False)
    backups: list = []
    published: list = []
    try:
        for _, final in pairs:
            if final.exists():
                backup = backup_dir / final.name
                os.replace(final, backup)
                backups.append((backup, final))
        for stage, final in pairs:
            _publish_artifact(stage, final)
            published.append((final, stage))
    except Exception:
        for final, stage in reversed(published):
            if final.exists():
                os.replace(final, stage)
        for backup, final in reversed(backups):
            if backup.exists():
                os.replace(backup, final)
        if backup_dir.exists():
            shutil.rmtree(backup_dir)
        if run_dir.exists():
            shutil.rmtree(run_dir)
        raise
    shutil.rmtree(backup_dir)

# ============================================================================
# MAIN FUNCTION
# ============================================================================


def hydrostream(
    input_dir: Union[str, Path],
    mode: str = "full",
    categories_file: Optional[Union[str, Path]] = None,
    years: Iterable[int] = range(2000, 2027),
    chunksize: int = 250_000,
    min_test_count: int = 50,
    generate_stats: bool = True,
    generate_qa_report: bool = True,
    save_log: bool = True,
    cutover_date: str = "2025-10-13",
    duckdb_memory_limit: str = "6GB",
    *,
    finalizer: str = "auto",
    unitless_quantitative_allowlist: Collection[Tuple[str, str]] = (),
    temp_dir: Optional[Union[str, Path]] = None,
    pandas_fallback_max_rows: int = PANDAS_FALLBACK_MAX_ROWS,
) -> Dict[str, Any]:
    """
    Clean and harmonise EA legacy + current API water-quality records.

    Parameters
    ----------
    input_dir
        Project root containing ``legacy_raw/`` and ``api_raw/``.

    mode
        ``full``, ``electrochemistry`` or ``contaminants``.

    categories_file
        Optional path to ``List of tests kept and categories.xlsx``.

    years
        Calendar years to include. Default includes 2000 through 2026.

    chunksize
        Number of rows read per CSV chunk.

    min_test_count
        In full mode, tests occurring fewer than this many times across the
        selected V2 source coverage are removed. Set 0 to disable.

    generate_stats
        Write descriptive/processing statistics workbook.

    generate_qa_report
        Write HTML QA report.

    save_log
        Save text processing log.

    cutover_date
        First date sourced from the current API. Legacy records on/after this
        date are excluded. Default ``2025-10-13``.

    duckdb_memory_limit
        DuckDB memory cap used for out-of-core finalisation.

    finalizer
        ``auto`` selects DuckDB when importable and otherwise permits pandas
        only below ``pandas_fallback_max_rows``. ``duckdb`` and ``pandas``
        force the named engine subject to the same safety checks.

    unitless_quantitative_allowlist
        Explicitly reviewed ``(DeterminandCode, Test)`` contexts allowed to
        retain API ``UNITLESS VALUE`` records. The production default is empty.

    temp_dir
        Optional existing parent directory for the unique DuckDB temporary
        workspace. Run staging remains on the output filesystem for atomic
        publication.

    pandas_fallback_max_rows
        Hard raw-row ceiling for the small-run pandas finaliser.
    """

    # ------------------------------------------------------------------
    # Validate intent and all inputs before creating output state.
    # ------------------------------------------------------------------
    root = Path(input_dir).expanduser().resolve()
    legacy_dir = root / "legacy_raw"
    api_dir = root / "api_raw"
    out_dir = root / "EA_processed_output_v2"

    mode = mode.strip().lower()
    if mode not in {"full", "electrochemistry", "contaminants"}:
        raise ValueError("mode must be full, electrochemistry, or contaminants")
    if int(chunksize) <= 0:
        raise ValueError("chunksize must be a positive integer")
    if int(pandas_fallback_max_rows) <= 0:
        raise ValueError("pandas_fallback_max_rows must be positive")

    year_set = {int(y) for y in years}
    if not year_set:
        raise ValueError("years cannot be empty")

    cutover = pd.Timestamp(cutover_date)
    if pd.isna(cutover):
        raise ValueError(f"Invalid cutover_date: {cutover_date!r}")

    requested_finalizer = str(finalizer).strip().lower()
    if requested_finalizer not in {"auto", "duckdb", "pandas"}:
        raise ValueError("finalizer must be auto, duckdb, or pandas")
    duckdb = _try_optional_package("duckdb", "duckdb")
    if requested_finalizer == "duckdb" and duckdb is None:
        raise ImportError(
            "The requested DuckDB finaliser is unavailable. Install the release "
            "dependencies from requirements-hydrostream-v2.txt; HydroStream "
            "will not install or silently replace the finaliser."
        )
    using_duckdb = requested_finalizer == "duckdb" or (
        requested_finalizer == "auto" and duckdb is not None
    )
    selected_finalizer = "duckdb" if using_duckdb else "pandas"

    sources, source_manifest = _build_source_manifest(
        root, year_set, int(chunksize), cutover
    )

    # Category decisions are optional for full/electrochemistry and mandatory
    # for contaminants, but any workbook selected for use must validate.
    cat_path: Optional[Path] = None
    if categories_file is not None:
        candidate = Path(categories_file).expanduser()
        cat_path = (
            (root / candidate).resolve() if not candidate.is_absolute()
            else candidate.resolve()
        )
        if not cat_path.is_file():
            raise FileNotFoundError(f"Category workbook does not exist: {cat_path}")
    else:
        candidates = [
            root / "List of tests kept and categories.xlsx",
            root / "categories.xlsx",
            root / "test_categories.xlsx",
            legacy_dir / "List of tests kept and categories.xlsx",
        ]
        cat_path = next((path for path in candidates if path.is_file()), None)

    category_map: Dict[str, str] = {}
    category_provenance: Optional[Dict[str, Any]] = None
    if cat_path is not None:
        category_map, category_provenance = _load_category_workbook(cat_path)
    elif mode == "contaminants":
        raise FileNotFoundError(
            "contaminants mode requires a validated category workbook"
        )

    temp_parent = (
        Path(temp_dir).expanduser().resolve() if temp_dir is not None else root
    )
    if not temp_parent.is_dir():
        raise FileNotFoundError(f"Temporary parent directory does not exist: {temp_parent}")
    capacity = _capacity_preflight(
        root,
        temp_parent,
        source_manifest,
        using_duckdb,
        duckdb_memory_limit,
        int(pandas_fallback_max_rows),
    )

    # All destructive/transient work begins only after the complete preflight.
    out_dir.mkdir(parents=True, exist_ok=True)
    run_dir = Path(tempfile.mkdtemp(prefix=".hydrostream-v2-run-", dir=out_dir))
    tmp_csv = run_dir / "hydrostream_v2_stream.csv"
    identity_index_csv = run_dir / "source_identity_index.csv"
    if temp_dir is None:
        duck_tmp_dir = run_dir / "duckdb_tmp"
        duck_tmp_dir.mkdir()
    else:
        duck_tmp_dir = Path(tempfile.mkdtemp(
            prefix="hydrostream-v2-duckdb-", dir=temp_parent
        ))

    started_utc = datetime.now(timezone.utc)
    log_buffer = io.StringIO()

    def log(msg: str = "") -> None:
        print(msg)
        log_buffer.write(msg + "\n")

    log("=" * 78)
    log("  HYDROSTREAM V2 — EA LEGACY + CURRENT API WATER QUALITY PIPELINE")
    log("=" * 78)
    log(f"  Version          : {HYDROSTREAM_VERSION}")
    log(f"  Mode             : {mode.upper()}")
    log(f"  Years requested  : {min(year_set)} – {max(year_set)}")
    log(f"  Legacy directory : {legacy_dir}")
    log(f"  API directory    : {api_dir}")
    log(f"  Source cutover   : legacy < {cutover.date()} | API >= {cutover.date()}")
    log(f"  Chunk size       : {chunksize:,}")
    log(f"  Min test count   : {min_test_count}")
    log(f"  Finaliser        : {selected_finalizer}")
    log(f"  Run workspace    : {run_dir}")
    log(f"  Started UTC      : {started_utc.isoformat()}")
    log("=" * 78)
    log()

    if cat_path is not None:
        assert category_provenance is not None
        log(
            f"Categories validated: {cat_path.name} "
            f"({len(category_map):,} normalized decisions; "
            f"SHA256 {category_provenance['sha256']})"
        )
    else:
        log("Categories file not supplied; Category output is disabled.")
    log()

    # ==================================================================
    # CONFIGURATION — SAME CORE SCIENTIFIC FILTERS AS V1
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

    DROP_TYPE_PATTERN = (
        r"(?:SEDIMENT|WHOLE ANIMAL|MUSCLE|LIVER|DIGESTIVE GLAND|BIOTA|"
        r"SOIL|ASH|WASTE\b|GAS|PRECIPITATION|CALIBRATION WATER|"
        r"POTABLE WATER|BOREHOLE GAS|ANY WATER\b|ANY NON-AQUEOUS LIQUID|"
        r"UNCODED|ANY AGRICULTURAL|ANY SEWAGE SLUDGE|ANY TIPPED|"
        r"ALGAE|SEAWEED|INVERTEBRATE|FISH|FLATFISH|BRYOPHYTE|"
        r"HIGHER PLANT|RANUNCULUS|FONTINALIS|ANY OIL|ANY BIOTA|"
        r"SOLID/SEDIMENT|MOSS|WRACK|COCKLE|MUSSEL|OYSTER|"
        r"SHRIMP|WORM|TELLIN|SCALLOP|TROUT|EEL|ROACH|FLOUNDER|"
        r"DAB|PLAICE|SOLE\b|WHITEBAIT|AIR\b|CONSTRUCTION|WHOLE PLANT)"
    )

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

    CONTAMINANTS_TEST_KEYS = {
        key for key, category in category_map.items()
        if category == "microplastics, nanoplastic, pfas, insecticide, pesticide, or similar"
    } if category_map else set()

    NON_QUANTITATIVE_UNITS = {
        "coded", "text", "yes/no", "pres/nf", "pres/nft",
        "garber c", "hh.mm", "ngr", "deccafix", "ug", "no/year",
    }

    BAD_TEST_FRAGMENTS = [
        "No flow", "No sample", "Site Inspection", "Present/Not found",
        "Pass/Fail", "Population Equivalent", "Sampling Frequency",
        "Photo Taken", "Weather :", "Bathing Water Profile",
        "National Grid Reference", "Sewage debris", "Foam Visible",
        "Colour : Abnormal", "Tarry residues", "MST Filtration",
        "Time of high tide", "Number of beach users", "Bathers per 100",
        "Type of flow", "State tide", "Colour (1/0)", "Tars/Floatg",
        "OilTypeQual", "WEATHER FLAG", "Borehole RefPt", "Sample Depth",
        "Laboratory Sample Number", "Dummy determinand", "Warning Sign",
        "Miscellaneous Identification", "Data Handling",
        "Mitochondrial Marker", "Mitochrondrial Marker", "Size range",
        "Size Range", "Length of fish", "Equiv.Carbon", "Equiv.carbo",
        "Equiv Carbon", "Biological examination", "Soli proportion",
        "24 hour Oyster", "Stone size", "Carbohydrate as Glucose",
        "Cohesive strength", "WQMS :", "Grain Size", "Number of bathers",
        "Number of birds", "Number of dogs",
    ]
    BAD_TEST_PATTERN = "|".join(re.escape(x) for x in BAD_TEST_FRAGMENTS)

    DUMMY_EASTING = 500_000
    DUMMY_NORTHINGS = {1, 2, 3, 4, 5, 6, 7, 8}

    # ==================================================================
    # QA COUNTERS
    # ==================================================================

    drop_counts = {
        "dummy_coordinates": 0,
        "non_water_types": 0,
        "non_quantitative_units": 0,
        "administrative_tests": 0,
        "test_filter": 0,
        "non_numeric_results": 0,
        "invalid_dates": 0,
        "outside_requested_year": 0,
        "unitless_value_not_allowlisted": 0,
    }

    integration_counts = {
        "legacy_rows_excluded_at_cutover": 0,
        "api_rows_before_cutover": 0,
        "qualified_numeric_results": 0,
        "less_than_results": 0,
        "greater_than_results": 0,
    }

    unit_crosswalk_counts: Counter = Counter()
    source_summary: list = []
    unreviewed_category_counts: Counter = Counter()

    # ==================================================================
    # HELPERS
    # ==================================================================

    transformer = Transformer.from_crs("EPSG:27700", "EPSG:4326", always_xy=True)

    def _optional_series(raw: pd.DataFrame, name: str) -> pd.Series:
        """Return an optional source field; required fields use direct access."""
        if name in raw.columns:
            return raw[name]
        return pd.Series(pd.NA, index=raw.index, dtype="object")

    def _text(series: pd.Series) -> pd.Series:
        return series.astype("string").str.strip()

    def _normalise_code(series: pd.Series) -> pd.Series:
        s = (_text(series).str.replace(r"\.0$", "", regex=True))
        out = s.copy()
        numeric = s.str.fullmatch(r"\d+", na=False)
        stripped = s.loc[numeric].str.lstrip("0")
        stripped = stripped.mask(stripped.eq(""), "0")
        out.loc[numeric] = stripped
        return out

    def _normalise_code_value(value: Any) -> str:
        return str(_normalise_code(pd.Series([value], dtype="string")).iloc[0])

    unitless_allowlist_keys = set()
    for decision in unitless_quantitative_allowlist:
        if not isinstance(decision, tuple) or len(decision) != 2:
            raise ValueError(
                "unitless_quantitative_allowlist entries must be "
                "(DeterminandCode, Test) tuples"
            )
        code, test = decision
        unitless_allowlist_keys.add(
            (_normalise_code_value(code), _normalise_category_key(test))
        )

    def _month_to_season(month: Any) -> Any:
        if pd.isna(month):
            return pd.NA
        month = int(month)
        if month in (12, 1, 2):
            return "Winter"
        if month in (3, 4, 5):
            return "Spring"
        if month in (6, 7, 8):
            return "Summer"
        return "Autumn"

    # ------------------------------------------------------------------
    # Current API verbose unit labels -> legacy-style canonical labels.
    # Unknown labels are retained and surfaced in the crosswalk output.
    # ------------------------------------------------------------------

    API_UNIT_MAP = {
        "MILLIGRAM PER LITRE": "mg/l",
        "MICROGRAM PER LITRE": "ug/l",
        "NANOGRAM PER LITRE": "ng/l",
        "PICOGRAM PER LITRE": "pg/l",
        "GRAM PER LITRE": "g/l",
        "MILLIGRAM PER KILOGRAM": "mg/kg",
        "MICROGRAM PER KILOGRAM": "ug/kg",
        "NANOGRAM PER KILOGRAM": "ng/kg",
        "PICOGRAM PER KILOGRAM": "pg/kg",
        "GRAM PER KILOGRAM": "g/kg",
        "PERCENTAGE": "%",
        "CELSIUS": "cel",
        "DEGREES CELSIUS": "cel",
        "PH UNITS": "phunits",
        "MICROSIEMENS PER CENTIMETRE": "uS/cm",
        "MICROSIEMENS PER CENTIMETER": "uS/cm",
        "MILLISIEMENS PER CENTIMETRE": "mS/cm",
        "MILLISIEMENS PER CENTIMETER": "mS/cm",
        "SIEMENS PER METRE": "S/m",
        "METRE": "m",
        "METER": "m",
        "CENTIMETRE": "cm",
        "CENTIMETER": "cm",
        "MILLIMETRE": "mm",
        "MILLIMETER": "mm",
        "MICROMETRE": "um",
        "MICROMETER": "um",
        "PARTS PER THOUSAND": "ppt",
        "PRACTICAL SALINITY UNIT": "psu",
        "NEPHELOMETRIC TURBIDITY UNIT": "NTU",
        "NEPHELOMETRIC TURBIDITY UNITS": "NTU",
        "FORMAZIN TURBIDITY UNIT": "FTU",
        "FORMAZIN TURBIDITY UNITS": "FTU",
        "NUMBER PER 100 MILLILITRE": "no/100ml",
        "NUMBER PER 100 MILLILITRES": "no/100ml",
        "NUMBER PER MILLILITRE": "no/ml",
        "NUMBER PER MILLILITRES": "no/ml",
        "NUMBER PER MICROLITRE": "no/ul",
        "NUMBER PER MICROLITRES": "no/ul",
        "NUMBER PER 10 MICROLITRE": "no/10ul",
        "NUMBER PER 10 MICROLITRES": "no/10ul",
        "COLONY FORMING UNITS PER 100 MILLILITRE": "cfu/100ml",
        "COLONY FORMING UNITS / 100ML": "cfu/100ml",
        "NUMBER PER HUNDRED MILLILITRES": "no/100ml",
        "MOST PROBABLE NUMBER PER 100 MILLILITRE": "mpn/100ml",
        "CODED RESULT": "coded",
        "TEXT": "text",
        "TEXT RESULT": "text",
        "YES(1) OR NO(0)": "yes/no",
        "YES / NO": "yes/no",
        "PRESENT/NOT FOUND": "pres/nf",
        "PRESENT / NOT FOUND": "pres/nf",
        "PRESENT/NOT FOUND/TRACE": "pres/nft",
        "PRESENT/NOT FOUND IN TEST": "pres/nft",
        "GARBER SURVEY ASSESSMENT SCALE CODE": "garber c",
        "NATIONAL GRID REFERENCE (METRIC)": "ngr",
        "TIME (HH.MM)": "hh.mm",
        "TIME HH.MM": "hh.mm",
        "HOUR (HH.MM)": "hh.mm",
        "NUMBER PER YEAR": "no/year",
        "UNITLESS VALUE": "unitless",
        "MICROGRAM": "ug",
        "MILLIGRAM": "mg",
        "GRAM": "g",
        "LITRE PER SECOND": "l/s",
        "LITRES PER SECOND": "l/s",
        "CUBIC METRE PER SECOND": "m3/s",
        "CUBIC METRES PER SECOND": "m3/s",
    }

    def _canonicalise_api_units(series: pd.Series) -> pd.Series:
        raw = _text(series)
        upper = raw.str.upper()
        mapped = upper.map(API_UNIT_MAP)

        # Generic SI mass-concentration names.
        generic_patterns = {
            r"^MILLIGRAMS? PER LITRE$": "mg/l",
            r"^MICROGRAMS? PER LITRE$": "ug/l",
            r"^NANOGRAMS? PER LITRE$": "ng/l",
            r"^PICOGRAMS? PER LITRE$": "pg/l",
            r"^GRAMS? PER LITRE$": "g/l",
            r"^MILLIGRAMS? PER KILOGRAM$": "mg/kg",
            r"^MICROGRAMS? PER KILOGRAM$": "ug/kg",
            r"^NANOGRAMS? PER KILOGRAM$": "ng/kg",
            r"^PICOGRAMS? PER KILOGRAM$": "pg/kg",
            r"^GRAMS? PER KILOGRAM$": "g/kg",
        }
        for pattern, replacement in generic_patterns.items():
            mask = mapped.isna() & upper.str.match(pattern, na=False)
            mapped.loc[mask] = replacement

        return mapped.fillna(raw)

    def _unit_policy_status(raw_units: pd.Series,
                            canonical_units: pd.Series,
                            source_format: str) -> pd.Series:
        if source_format == "legacy":
            return pd.Series(
                "SOURCE_NATIVE_LEGACY", index=raw_units.index, dtype="string"
            )
        upper = _text(raw_units).str.upper()
        status = pd.Series("UNREVIEWED_RETAINED", index=raw_units.index, dtype="string")
        mapped = canonical_units.astype("string").ne(raw_units.astype("string"))
        status.loc[mapped] = "REVIEWED_MAPPING"
        status.loc[upper.eq("UNITLESS VALUE")] = "UNITLESS_REVIEW_REQUIRED"
        status.loc[canonical_units.astype("string").str.lower().isin(
            NON_QUANTITATIVE_UNITS
        )] = "NON_QUANTITATIVE"
        return status

    def _record_unit_crosswalk(source_format: str,
                               raw_units: pd.Series,
                               canonical_units: pd.Series,
                               statuses: pd.Series) -> None:
        pairs = pd.DataFrame({
            "raw": raw_units.astype("string").fillna("<NA>"),
            "canonical": canonical_units.astype("string").fillna("<NA>"),
            "status": statuses.astype("string").fillna("<NA>"),
        }).value_counts(dropna=False)
        for (raw_u, can_u, status), n in pairs.items():
            unit_crosswalk_counts[
                (source_format, str(raw_u), str(can_u), str(status))
            ] += int(n)

    def _legacy_sample_id(record_id: pd.Series) -> pd.Series:
        return record_id.astype("string").str.extract(
            r"-([0-9]+)-[^-/?#]+$", expand=False
        )

    def _api_sample_id(record_id: pd.Series) -> pd.Series:
        return record_id.astype("string").str.extract(
            r"/sample/([^/]+)/observation/", expand=False
        )

    def _legacy_to_canonical(raw: pd.DataFrame,
                             source_file: str,
                             year_hint: int,
                             track_qa: bool = True) -> pd.DataFrame:
        _validate_schema(raw.columns, "legacy")
        df = pd.DataFrame(index=raw.index)
        df["SourceRecordID"] = _text(raw["@id"])
        df["SamplingPointCode"] = _text(raw["sample.samplingPoint.notation"])
        df["Sampling Point"] = _text(raw["sample.samplingPoint.label"])
        df["Type"] = _text(raw["sample.sampledMaterialType.label"])
        df["Date"] = pd.to_datetime(
            raw["sample.sampleDateTime"], errors="coerce", format="mixed"
        )
        df["DeterminandCode"] = _normalise_code(raw["determinand.notation"])
        df["Test"] = _text(raw["determinand.definition"])

        raw_result = _text(raw["result"])
        df["RawResult"] = raw_result
        df["result"] = pd.to_numeric(raw_result, errors="coerce").astype("float64")
        df["ResultQualifier"] = _text(raw["resultQualifier.notation"])

        raw_unit = _text(raw["determinand.unit.label"])
        df["RawUnit"] = raw_unit
        df["Unit"] = raw_unit
        df["UnitPolicyStatus"] = _unit_policy_status(raw_unit, raw_unit, "legacy")
        df["UnitConversionCode"] = "NONE"
        if track_qa:
            _record_unit_crosswalk(
                "legacy", raw_unit, raw_unit, df["UnitPolicyStatus"]
            )

        df["Easting"] = pd.to_numeric(
            raw["sample.samplingPoint.easting"], errors="coerce"
        )
        df["Northing"] = pd.to_numeric(
            raw["sample.samplingPoint.northing"], errors="coerce"
        )
        df["Latitude"] = np.nan
        df["Longitude"] = np.nan

        df["Region"] = pd.NA
        df["Area"] = pd.NA
        df["SubArea"] = pd.NA
        df["SamplingPointStatus"] = pd.NA
        df["SamplingPointType"] = pd.NA
        df["SamplingPurpose"] = _text(_optional_series(raw, "sample.purpose.label"))
        # Retain source-level optional provenance for exact repeated-identity
        # comparison.  These fields are comparison-only and are not added to
        # the cleaned/public scientific schema.
        df["LegacySamplingPointURI"] = _text(
            _optional_series(raw, "sample.samplingPoint")
        )
        df["LegacyDeterminandLabel"] = _text(
            _optional_series(raw, "determinand.label")
        )
        df["LegacyCodedResultInterpretation"] = _text(
            _optional_series(raw, "codedResultInterpretation.interpretation")
        )
        df["LegacyIsComplianceSample"] = _text(
            _optional_series(raw, "sample.isComplianceSample")
        )

        df["SampleID"] = _legacy_sample_id(df["SourceRecordID"])
        df["SourceFormat"] = "legacy"
        df["SourceFile"] = source_file
        df["YearHint"] = year_hint
        return df

    def _api_to_canonical(raw: pd.DataFrame,
                          source_file: str,
                          year_hint: int,
                          track_qa: bool = True) -> pd.DataFrame:
        _validate_schema(raw.columns, "api")
        df = pd.DataFrame(index=raw.index)
        df["SourceRecordID"] = _text(raw["id"])
        df["SamplingPointCode"] = _text(raw["samplingPoint.notation"])
        df["Sampling Point"] = _text(raw["samplingPoint.prefLabel"])
        df["Type"] = _text(raw["sampleMaterialType"])
        df["Date"] = pd.to_datetime(
            raw["phenomenonTime"], errors="coerce", format="mixed"
        )
        df["DeterminandCode"] = _normalise_code(raw["determinand.notation"])
        df["Test"] = _text(raw["determinand.prefLabel"])

        raw_result = _text(raw["result"])
        qualifier = raw_result.str.extract(r"^\s*(<=|>=|<|>)", expand=False).astype("string")
        numeric_text = raw_result.str.replace(
            r"^\s*(?:<=|>=|<|>)\s*", "", regex=True
        )
        numeric = pd.to_numeric(numeric_text, errors="coerce").astype("float64")
        df["RawResult"] = raw_result
        df["ResultQualifier"] = qualifier
        df["result"] = numeric

        if track_qa:
            qmask = qualifier.notna() & numeric.notna()
            integration_counts["qualified_numeric_results"] += int(qmask.sum())
            integration_counts["less_than_results"] += int(
                (qualifier.isin(["<", "<="]) & numeric.notna()).sum()
            )
            integration_counts["greater_than_results"] += int(
                (qualifier.isin([">", ">="]) & numeric.notna()).sum()
            )

        raw_unit = _text(raw["unit"])
        canonical_unit = _canonicalise_api_units(raw_unit)
        df["RawUnit"] = raw_unit
        df["Unit"] = canonical_unit
        df["UnitPolicyStatus"] = _unit_policy_status(
            raw_unit, canonical_unit, "api"
        )
        df["UnitConversionCode"] = np.where(
            canonical_unit.astype("string").ne(raw_unit.astype("string")),
            "LABEL_ONLY",
            "NONE",
        )
        if track_qa:
            _record_unit_crosswalk(
                "api", raw_unit, canonical_unit, df["UnitPolicyStatus"]
            )

        df["Easting"] = np.nan
        df["Northing"] = np.nan
        df["Longitude"] = pd.to_numeric(
            raw["samplingPoint.longitude"], errors="coerce"
        )
        df["Latitude"] = pd.to_numeric(
            raw["samplingPoint.latitude"], errors="coerce"
        )

        df["Region"] = _text(_optional_series(raw, "samplingPoint.region"))
        df["Area"] = _text(_optional_series(raw, "samplingPoint.area"))
        df["SubArea"] = _text(_optional_series(raw, "samplingPoint.subArea"))
        df["SamplingPointStatus"] = _text(
            _optional_series(raw, "samplingPoint.samplingPointStatus")
        )
        df["SamplingPointType"] = _text(
            _optional_series(raw, "samplingPoint.samplingPointType")
        )
        df["SamplingPurpose"] = _text(_optional_series(raw, "samplingPurpose"))
        df["LegacySamplingPointURI"] = pd.NA
        df["LegacyDeterminandLabel"] = pd.NA
        df["LegacyCodedResultInterpretation"] = pd.NA
        df["LegacyIsComplianceSample"] = pd.NA

        df["SampleID"] = _api_sample_id(df["SourceRecordID"])
        df["SourceFormat"] = "api"
        df["SourceFile"] = source_file
        df["YearHint"] = year_hint
        return df

    def _apply_cutover(df: pd.DataFrame,
                       source_format: str,
                       track_qa: bool = True) -> pd.DataFrame:
        if source_format == "legacy":
            keep = df["Date"].isna() | (df["Date"] < cutover)
            if track_qa:
                integration_counts["legacy_rows_excluded_at_cutover"] += int((~keep).sum())
        else:
            keep = df["Date"].isna() | (df["Date"] >= cutover)
            if track_qa:
                integration_counts["api_rows_before_cutover"] += int((~keep).sum())
        return df.loc[keep].copy()

    def _convert_legacy_coordinates(df: pd.DataFrame) -> pd.DataFrame:
        mask = (
            df["SourceFormat"].eq("legacy") &
            df["Easting"].notna() &
            df["Northing"].notna()
        )
        if mask.any():
            lon, lat = transformer.transform(
                df.loc[mask, "Easting"].to_numpy(),
                df.loc[mask, "Northing"].to_numpy(),
            )
            df.loc[mask, "Longitude"] = lon
            df.loc[mask, "Latitude"] = lat
        return df

    def _standardise_units(df: pd.DataFrame) -> pd.DataFrame:
        # pandas 3 can infer an integer extension dtype for one-row chunks.
        # Force stable floating semantics before any fractional conversion.
        df["result"] = pd.to_numeric(df["result"], errors="coerce").astype("float64")
        u = (_text(df["Unit"])
             .str.replace("µ", "u", regex=False)
             .str.replace("μ", "u", regex=False))
        df["Unit"] = u
        df["Test"] = df["Test"].str.replace(
            "Conductivity at 20C", "Conductivity at 20 C", regex=False
        )

        lower = df["Unit"].str.lower()
        conversions = [
            ("ug/l", 1 / 1_000, "mg/l", "UG_L_TO_MG_L_DIV_1000"),
            ("ng/l", 1 / 1_000_000, "mg/l", "NG_L_TO_MG_L_DIV_1000000"),
            ("pg/l", 1 / 1_000_000_000, "mg/l", "PG_L_TO_MG_L_DIV_1000000000"),
            ("g/l", 1_000, "mg/l", "G_L_TO_MG_L_MUL_1000"),
            ("ms/cm", 1_000, "uS/cm", "MS_CM_TO_US_CM_MUL_1000"),
            ("no/ml", 100, "no/100ml", "NO_ML_TO_NO_100ML_MUL_100"),
            ("no/ul", 100_000, "no/100ml", "NO_UL_TO_NO_100ML_MUL_100000"),
            ("no/10ul", 10_000, "no/100ml", "NO_10UL_TO_NO_100ML_MUL_10000"),
        ]
        for source_unit, factor, target_unit, conversion_code in conversions:
            mask = lower.eq(source_unit)
            if mask.any():
                df.loc[mask, "result"] = df.loc[mask, "result"] * factor
                df.loc[mask, "Unit"] = target_unit
                df.loc[mask, "UnitConversionCode"] = conversion_code

        mask = df["Unit"].astype("string").str.lower().eq("us/cm")
        if mask.any():
            df.loc[mask, "Unit"] = "uS/cm"

        return df

    def _source_identity(df: pd.DataFrame) -> pd.Series:
        record_id = df["SourceRecordID"].astype("string").str.strip()
        missing = record_id.isna() | record_id.eq("")
        if missing.any():
            files = sorted(df.loc[missing, "SourceFile"].astype(str).unique())
            raise SourceValidationError(
                "Required SourceRecordID is blank in " + ", ".join(files[:10])
            )
        return df["SourceFormat"].astype("string") + "|" + record_id

    def _exact_payload_frame(df: pd.DataFrame) -> pd.DataFrame:
        """Normalize canonical identity fields for null-aware exact comparison."""
        out = pd.DataFrame(index=df.index)
        out["SourceIdentity"] = _source_identity(df).astype("string")
        out["SourceFile"] = df["SourceFile"].astype("string")
        numeric_fields = {
            "result", "Easting", "Northing", "Latitude", "Longitude",
        }
        for column in IDENTITY_EQUIVALENCE_FIELDS:
            values = df[column]
            if column == "Date":
                out[column] = values.map(
                    lambda value: pd.NA if pd.isna(value)
                    else pd.Timestamp(value).isoformat()
                ).astype("string")
            elif column in numeric_fields:
                out[column] = values.map(
                    lambda value: pd.NA if pd.isna(value)
                    else float(value).hex()
                ).astype("string")
            else:
                out[column] = values.map(
                    lambda value: pd.NA if pd.isna(value) else (
                        unicodedata.normalize("NFC", str(value).strip()) or pd.NA
                    )
                ).astype("string")
        return out.reset_index(drop=True)

    def _read_exact_candidate_payloads(
        candidate_sources: Collection[Tuple[str, str]],
        candidate_identities: Optional[Collection[str]] = None,
    ) -> Iterable[pd.DataFrame]:
        """Re-read only files containing repeated identities and yield matches."""
        source_keys = set(candidate_sources)
        identity_set = (
            set(candidate_identities) if candidate_identities is not None else None
        )
        for src in sources:
            source_key = (str(src["format"]), str(src["path"].name))
            if source_key not in source_keys:
                continue
            for raw in pd.read_csv(
                src["path"], chunksize=chunksize, low_memory=False,
                encoding=src["encoding"], dtype=str,
                keep_default_na=False, na_values=[""],
            ):
                if src["format"] == "legacy":
                    canonical = _legacy_to_canonical(
                        raw, src["path"].name, src["year"], track_qa=False
                    )
                else:
                    canonical = _api_to_canonical(
                        raw, src["path"].name, src["year"], track_qa=False
                    )
                exact = _exact_payload_frame(canonical)
                if identity_set is not None:
                    exact = exact.loc[
                        exact["SourceIdentity"].isin(identity_set)
                    ].copy()
                if not exact.empty:
                    yield exact

    def _make_observation_key(df: pd.DataFrame) -> pd.Series:
        site = df["SamplingPointCode"].astype("string")
        sample = df["SampleID"].astype("string")
        det = df["DeterminandCode"].astype("string")
        date = df["Date"].astype("string")

        primary_ok = site.notna() & sample.notna() & det.notna() & date.notna()
        key = pd.Series(pd.NA, index=df.index, dtype="string")
        key.loc[primary_ok] = (
            "EA|" + site.loc[primary_ok] + "|" +
            sample.loc[primary_ok] + "|" + det.loc[primary_ok] + "|" +
            date.loc[primary_ok]
        )

        fallback_idx = key.index[~primary_ok]
        if len(fallback_idx):
            cols = [
                "Date", "SamplingPointCode", "Sampling Point", "Type",
                "DeterminandCode", "Test", "result", "ResultQualifier", "Unit",
            ]
            tmp = df.loc[fallback_idx, cols].copy()
            for c in cols:
                tmp[c] = tmp[c].astype("string").fillna("")
            joined = tmp.agg("|".join, axis=1)
            key.loc[fallback_idx] = joined.map(
                lambda x: "FB|" + hashlib.sha1(x.encode("utf-8")).hexdigest()
            )
        return key

    def _clean_chunk(df: pd.DataFrame,
                     source_format: str,
                     test_filter: Optional[set]) -> pd.DataFrame:
        df = _apply_cutover(df, source_format, track_qa=True)
        if df.empty:
            return df

        # Valid dates only.
        n0 = len(df)
        df = df[df["Date"].notna()].copy()
        drop_counts["invalid_dates"] += n0 - len(df)
        if df.empty:
            return df

        df["SourceYear"] = df["Date"].dt.year.astype("Int64")
        df["Season"] = df["Date"].dt.month.map(_month_to_season).astype("string")

        # Parsed observation dates, never filenames, control requested scope.
        n0 = len(df)
        df = df.loc[df["SourceYear"].isin(year_set)].copy()
        drop_counts["outside_requested_year"] += n0 - len(df)
        if df.empty:
            return df

        # Legacy placeholder coordinates.
        dummy = (
            df["SourceFormat"].eq("legacy") &
            df["Easting"].eq(DUMMY_EASTING) &
            df["Northing"].isin(DUMMY_NORTHINGS)
        )
        if dummy.any():
            drop_counts["dummy_coordinates"] += int(dummy.sum())
            df = df.loc[~dummy].copy()
        if df.empty:
            return df

        # Coordinate harmonisation.
        df = _convert_legacy_coordinates(df)
        bad_lat = df["Latitude"].notna() & ~df["Latitude"].between(-90, 90)
        bad_lon = df["Longitude"].notna() & ~df["Longitude"].between(-180, 180)
        df.loc[bad_lat, "Latitude"] = np.nan
        df.loc[bad_lon, "Longitude"] = np.nan

        # Water matrices only (same policy as V1).
        n0 = len(df)
        type_text = df["Type"].astype("string")
        keep = (
            type_text.isin(WATER_TYPES) &
            ~type_text.str.contains(DROP_TYPE_PATTERN, case=False, na=False)
        )
        df = df.loc[keep].copy()
        drop_counts["non_water_types"] += n0 - len(df)
        if df.empty:
            return df

        # Fail-safe UNITLESS policy, then the remaining non-quantitative units.
        unit_lower = _text(df["Unit"]).str.lower()
        # The release finding concerns the current API's literal
        # ``UNITLESS VALUE`` policy. Legacy-native ``unitless`` is established
        # V1-compatible quantitative behaviour and must remain unchanged.
        unitless = df["SourceFormat"].eq("api") & unit_lower.eq("unitless")
        if unitless.any():
            context = pd.MultiIndex.from_arrays([
                df["DeterminandCode"].astype("string"),
                df["Test"].map(_normalise_category_key),
            ])
            explicitly_allowed = pd.Series(
                context.isin(unitless_allowlist_keys), index=df.index
            )
            rejected_unitless = unitless & ~explicitly_allowed
            drop_counts["unitless_value_not_allowlisted"] += int(
                rejected_unitless.sum()
            )
            df.loc[unitless & explicitly_allowed, "UnitPolicyStatus"] = (
                "EXPLICIT_UNITLESS_ALLOWLIST"
            )
            df = df.loc[~rejected_unitless].copy()
        if df.empty:
            return df

        unit_lower = _text(df["Unit"]).str.lower()
        non_quantitative = unit_lower.isin(NON_QUANTITATIVE_UNITS)
        drop_counts["non_quantitative_units"] += int(non_quantitative.sum())
        df = df.loc[~non_quantitative].copy()
        if df.empty:
            return df

        # Administrative/procedural tests.
        n0 = len(df)
        bad = df["Test"].astype("string").str.contains(
            BAD_TEST_PATTERN, case=False, na=False
        )
        df = df.loc[~bad].copy()
        drop_counts["administrative_tests"] += n0 - len(df)
        if df.empty:
            return df

        normalized_test = df["Test"].map(_normalise_category_key)
        if category_map:
            df["Category"] = normalized_test.map(category_map).fillna("uncategorized")

        # Contaminants decisions fail closed: first inventory every retained
        # analyte without a reviewed workbook decision, then select only the
        # explicitly reviewed contaminant category.
        if mode == "contaminants":
            missing_category = ~normalized_test.isin(category_map)
            if missing_category.any():
                grouped = df.loc[missing_category].groupby(
                    ["DeterminandCode", "Test"], dropna=False
                ).size()
                for key, count in grouped.items():
                    unreviewed_category_counts[key] += int(count)
            keep_contaminants = normalized_test.isin(CONTAMINANTS_TEST_KEYS)
            drop_counts["test_filter"] += int((~keep_contaminants).sum())
            df = df.loc[keep_contaminants].copy()
            if df.empty:
                return df
        elif test_filter is not None:
            n0 = len(df)
            df = df.loc[df["Test"].isin(test_filter)].copy()
            drop_counts["test_filter"] += n0 - len(df)
            if df.empty:
                return df

        # Unit conversions work on already-parsed numeric values.
        df = _standardise_units(df)

        # Remove genuinely non-numeric/categorical results after the other
        # documented filters. Qualified numeric API values survive because the
        # qualifier was separated before this point.
        n0 = len(df)
        df = df.loc[df["result"].notna()].copy()
        drop_counts["non_numeric_results"] += n0 - len(df)
        if df.empty:
            return df

        df["ObservationKey"] = _make_observation_key(df)
        df["SourceIdentity"] = _source_identity(df)

        # Stable streaming schema. These additional provenance/station fields
        # are used internally and for stations/QA outputs; the main public CSV
        # is narrower.
        cols = [
            "SamplingPointCode", "Sampling Point", "Type", "Date",
            "DeterminandCode", "Test", "result", "ResultQualifier", "Unit",
            "Season", "SourceYear", "Latitude", "Longitude",
            "Region", "Area", "SubArea", "SamplingPointStatus",
            "SamplingPointType", "SamplingPurpose", "SourceFormat",
            "SourceFile", "SourceRecordID", "SampleID", "ObservationKey",
            "SourceIdentity", "RawUnit", "UnitPolicyStatus",
            "UnitConversionCode",
        ]
        if category_map:
            cols.append("Category")
        return df[cols]

    # ==================================================================
    # VALIDATED SOURCES
    # ==================================================================

    log(f"Found {sum(s['format']=='legacy' for s in sources)} legacy files and "
        f"{sum(s['format']=='api' for s in sources)} API files:\n")
    for src in sources:
        log(f"  {src['format']:<6} {src['year']} -> {src['path'].name}")
    log()

    # ==================================================================
    # DETERMINE TEST FILTER
    # ==================================================================

    if mode == "electrochemistry":
        test_filter: Optional[set[str]] = set(ELECTROCHEMISTRY_TESTS)
        log(f"Mode ELECTROCHEMISTRY: {len(ELECTROCHEMISTRY_TESTS)} tests.\n")

    elif mode == "contaminants":
        if not CONTAMINANTS_TEST_KEYS:
            raise ValueError(
                "Validated category workbook contains no reviewed contaminants decisions"
            )
        test_filter = None
        log(
            f"Mode CONTAMINANTS: {len(CONTAMINANTS_TEST_KEYS)} reviewed "
            "normalized test decisions; unreviewed analytes will fail closed.\n"
        )

    else:
        if min_test_count <= 0:
            test_filter = None
            log("Mode FULL: rare-test filter disabled.\n")
        else:
            log(f"Mode FULL: counting canonical test names (threshold {min_test_count}) ...")
            test_counts: Counter = Counter()

            for src in sources:
                for raw in pd.read_csv(
                    src["path"], chunksize=chunksize, low_memory=False,
                    encoding=src["encoding"], dtype=str,
                    keep_default_na=False, na_values=[""],
                ):
                    if src["format"] == "legacy":
                        canonical = _legacy_to_canonical(
                            raw, src["path"].name, src["year"], track_qa=False
                        )
                    else:
                        canonical = _api_to_canonical(
                            raw, src["path"].name, src["year"], track_qa=False
                        )
                    canonical = _apply_cutover(
                        canonical, src["format"], track_qa=False
                    )
                    canonical = canonical[canonical["Date"].notna()]
                    canonical = canonical.loc[
                        canonical["Date"].dt.year.isin(year_set)
                    ]
                    counts = canonical["Test"].dropna().value_counts()
                    test_counts.update({str(k): int(v) for k, v in counts.items()})

            test_filter = {
                name for name, count in test_counts.items()
                if count >= min_test_count
            }
            log(f"  Unique canonical tests : {len(test_counts):,}")
            log(f"  Tests retained         : {len(test_filter):,}")
            log(f"  Tests below threshold  : {len(test_counts)-len(test_filter):,}\n")

    # ==================================================================
    # MAIN STREAMING PASS
    # ==================================================================

    total_raw = 0
    total_clean_pre_dedup = 0
    header_written = False
    con: Any = None
    identity_parts: list[pd.DataFrame] = []

    if using_duckdb:
        duck_tmp_sql = str(duck_tmp_dir).replace("'", "''")
        con = duckdb.connect(database=str(run_dir / "hydrostream_work.duckdb"))
        con.execute(f"SET temp_directory='{duck_tmp_sql}'")
        con.execute(f"SET memory_limit='{duckdb_memory_limit}'")
        con.execute("SET threads=2")
        con.execute("SET preserve_insertion_order=false")
        con.execute("""
            CREATE TABLE source_identity_index (
                SourceIdentity VARCHAR,
                SourceFile VARCHAR,
                IdentityHash UBIGINT
            )
        """)

    for src in sources:
        path = src["path"]
        enc = src["encoding"]
        raw_count = 0
        clean_count = 0

        log(f"-- Processing {src['format'].upper():<6} {path.name} " + "-" * 25)

        for raw in pd.read_csv(
            path, chunksize=chunksize, low_memory=False,
            encoding=enc, dtype=str,
            keep_default_na=False, na_values=[""],
        ):
            raw_count += len(raw)
            total_raw += len(raw)

            if src["format"] == "legacy":
                canonical = _legacy_to_canonical(
                    raw, path.name, src["year"], track_qa=True
                )
            else:
                canonical = _api_to_canonical(
                    raw, path.name, src["year"], track_qa=True
                )

            # Index identity before any scientific filter so a conflicting
            # copy cannot disappear and evade the source-record check. A
            # fixed-width 64-bit hash is used only to find candidate repeated
            # identity buckets at archive scale. Hash collisions can only add
            # candidates; exact SourceIdentity grouping below decides whether
            # a duplicate identity truly exists.
            canonical["SourceIdentity"] = _source_identity(canonical)
            identity_hash = pd.util.hash_pandas_object(
                canonical["SourceIdentity"].astype("string"),
                index=False,
                hash_key="hydrostream-key1",
            ).to_numpy(dtype="uint64")
            identity_chunk = pd.DataFrame({
                "SourceIdentity": canonical["SourceIdentity"].astype("string"),
                "SourceFile": canonical["SourceFile"].astype("string"),
                "IdentityHash": identity_hash,
            })
            if using_duckdb:
                con.register("identity_chunk", identity_chunk)
                con.execute("INSERT INTO source_identity_index SELECT * FROM identity_chunk")
                con.unregister("identity_chunk")
            else:
                identity_parts.append(identity_chunk)

            cleaned = _clean_chunk(canonical, src["format"], test_filter)
            if cleaned.empty:
                continue

            cleaned.to_csv(
                tmp_csv, mode="a", index=False,
                header=(not header_written),
            )
            header_written = True
            clean_count += len(cleaned)
            total_clean_pre_dedup += len(cleaned)

        source_summary.append({
            "source_format": src["format"],
            "year_hint": src["year"],
            "source_file": path.name,
            "raw_rows_read": raw_count,
            "clean_rows_pre_dedup": clean_count,
        })

        pct = clean_count / raw_count * 100 if raw_count else 0
        log(f"  Raw rows read        : {raw_count:>12,}")
        log(f"  Clean pre-dedup rows : {clean_count:>12,} ({pct:.1f}%)")
        log()

    expected_raw = int(capacity["raw_rows"])
    if total_raw != expected_raw:
        raise RuntimeError(
            "Source manifest/main-pass row count mismatch: "
            f"manifest={expected_raw:,}, read={total_raw:,}"
        )

    # Exact source identity governs deletion. The fixed-width identity hash
    # locates candidates only; every repeated identity is re-read and compared
    # field-for-field before it can be classified as an equivalent copy.
    # ObservationKey remains only a linkage/replicate diagnostic.
    exact_conflict_payloads = pd.DataFrame()
    if using_duckdb:
        # Production-scale identity scan. First group only the fixed-width
        # candidate hash. DuckDB can spill GROUP BY to disk; this avoids an
        # archive-wide hash table containing tens of millions of long identity
        # strings. Equal SourceIdentity values always share the same hash, so
        # no true repeated identity can be missed. A hash collision merely
        # sends extra rows to the exact SourceIdentity grouping below.
        log("Checking source identities for repeated EA records ...")
        con.execute("""
            CREATE TABLE repeated_identity_hashes AS
            SELECT IdentityHash, count(*) AS hash_copies
            FROM source_identity_index
            GROUP BY IdentityHash
            HAVING count(*) > 1
        """)

        # Exact SourceIdentity is authoritative. Only hash buckets with more
        # than one row reach this second grouping step.
        con.execute("""
            CREATE TABLE source_identity_groups AS
            SELECT
                i.SourceIdentity,
                count(*) AS source_copies,
                CAST(0 AS BIGINT) AS payload_variants,
                min(i.SourceFile) AS chosen_source_file,
                CAST(NULL AS VARCHAR) AS equivalent_source_files
            FROM source_identity_index i
            JOIN repeated_identity_hashes h USING (IdentityHash)
            GROUP BY i.SourceIdentity
            HAVING count(*) > 1
        """)
        duplicate_identity_groups = int(con.execute(
            "SELECT count(*) FROM source_identity_groups"
        ).fetchone()[0])
        log(f"  Repeated SourceIdentity groups: {duplicate_identity_groups:,}")

        if duplicate_identity_groups:
            # string_agg is intentionally restricted to the already-reduced
            # repeated-identity subset. DuckDB does not spill string_agg's
            # complex aggregate state, so it must never run over the complete
            # archive.
            con.execute("""
                CREATE TABLE repeated_identity_files AS
                SELECT SourceIdentity,
                       string_agg(SourceFile, ';' ORDER BY SourceFile)
                           AS equivalent_source_files
                FROM (
                    SELECT DISTINCT i.SourceIdentity, i.SourceFile
                    FROM source_identity_index i
                    JOIN source_identity_groups g USING (SourceIdentity)
                ) repeated_files
                GROUP BY SourceIdentity
            """)
            con.execute("""
                UPDATE source_identity_groups AS g
                SET equivalent_source_files = r.equivalent_source_files
                FROM repeated_identity_files AS r
                WHERE g.SourceIdentity = r.SourceIdentity
            """)

            candidate_sources = {
                (str(identity).split("|", 1)[0], str(source_file))
                for identity, source_file in con.execute("""
                    SELECT DISTINCT i.SourceIdentity, i.SourceFile
                    FROM source_identity_index i
                    JOIN source_identity_groups g USING (SourceIdentity)
                """).fetchall()
            }
            exact_columns = [
                "SourceIdentity", "SourceFile", *IDENTITY_EQUIVALENCE_FIELDS,
            ]
            exact_definitions = ", ".join(
                f'"{column}" VARCHAR' for column in exact_columns
            )
            con.execute(
                f"CREATE TABLE source_identity_exact ({exact_definitions})"
            )
            for exact_chunk in _read_exact_candidate_payloads(candidate_sources):
                con.register("exact_candidate_chunk", exact_chunk[exact_columns])
                con.execute("""
                    INSERT INTO source_identity_exact
                    SELECT c.*
                    FROM exact_candidate_chunk c
                    JOIN source_identity_groups g USING (SourceIdentity)
                """)
                con.unregister("exact_candidate_chunk")

            expected_exact_rows = int(con.execute(
                "SELECT coalesce(sum(source_copies), 0) "
                "FROM source_identity_groups"
            ).fetchone()[0])
            observed_exact_rows = int(con.execute(
                "SELECT count(*) FROM source_identity_exact"
            ).fetchone()[0])
            if observed_exact_rows != expected_exact_rows:
                raise RuntimeError(
                    "Repeated-source exact comparison was incomplete: "
                    f"expected {expected_exact_rows:,} candidate rows, "
                    f"re-read {observed_exact_rows:,}"
                )

            exact_projection = ", ".join(
                f'"{column}"' for column in IDENTITY_EQUIVALENCE_FIELDS
            )
            con.execute(f"""
                CREATE TABLE source_identity_exact_variant_counts AS
                SELECT SourceIdentity, count(*) AS payload_variants
                FROM (
                    SELECT DISTINCT SourceIdentity, {exact_projection}
                    FROM source_identity_exact
                )
                GROUP BY SourceIdentity
            """)
            con.execute("""
                UPDATE source_identity_groups AS groups
                SET payload_variants = exact_counts.payload_variants
                FROM source_identity_exact_variant_counts AS exact_counts
                WHERE groups.SourceIdentity = exact_counts.SourceIdentity
            """)
        conflict_rows = con.execute("""
            SELECT SourceIdentity, source_copies, payload_variants,
                   equivalent_source_files
            FROM source_identity_groups
            WHERE payload_variants > 1
            ORDER BY SourceIdentity
            LIMIT 20
        """).fetchdf()
        if not conflict_rows.empty:
            con.register(
                "conflicting_identity_keys",
                conflict_rows[["SourceIdentity"]],
            )
            exact_conflict_payloads = con.execute("""
                SELECT exact_rows.*
                FROM source_identity_exact exact_rows
                JOIN conflicting_identity_keys keys USING (SourceIdentity)
                ORDER BY exact_rows.SourceIdentity, exact_rows.SourceFile
            """).fetchdf()
            con.unregister("conflicting_identity_keys")
        equivalent_duplicate_rows = int(con.execute(
            "SELECT coalesce(sum(source_copies - 1), 0) "
            "FROM source_identity_groups WHERE payload_variants = 1"
        ).fetchone()[0])
    else:
        identity_index = pd.concat(identity_parts, ignore_index=True)
        copies = identity_index.groupby("SourceIdentity", sort=False).size()
        repeated = copies[copies > 1]
        group_rows = []
        for source_identity, source_copies in repeated.items():
            source_rows = identity_index.loc[
                identity_index["SourceIdentity"].eq(source_identity)
            ]
            source_files = sorted(source_rows["SourceFile"].astype(str).unique())
            group_rows.append({
                "SourceIdentity": source_identity,
                "source_copies": int(source_copies),
                "payload_variants": 0,
                "chosen_source_file": source_files[0],
                "equivalent_source_files": ";".join(source_files),
            })
        source_identity_groups_df = pd.DataFrame(group_rows, columns=[
            "SourceIdentity", "source_copies", "payload_variants",
            "chosen_source_file", "equivalent_source_files",
        ])
        if not source_identity_groups_df.empty:
            candidate_identities = set(
                source_identity_groups_df["SourceIdentity"].astype(str)
            )
            candidate_index_rows = identity_index.loc[
                identity_index["SourceIdentity"].isin(candidate_identities)
            ]
            candidate_sources = {
                (str(row.SourceIdentity).split("|", 1)[0], str(row.SourceFile))
                for row in candidate_index_rows.itertuples()
            }
            exact_parts = list(_read_exact_candidate_payloads(
                candidate_sources, candidate_identities
            ))
            if not exact_parts:
                raise RuntimeError(
                    "Repeated-source exact comparison found no candidate rows"
                )
            exact_payloads = pd.concat(exact_parts, ignore_index=True)
            expected_exact_rows = int(
                source_identity_groups_df["source_copies"].sum()
            )
            if len(exact_payloads) != expected_exact_rows:
                raise RuntimeError(
                    "Repeated-source exact comparison was incomplete: "
                    f"expected {expected_exact_rows:,} candidate rows, "
                    f"re-read {len(exact_payloads):,}"
                )
            exact_variants = (
                exact_payloads.drop_duplicates(
                    ["SourceIdentity", *IDENTITY_EQUIVALENCE_FIELDS]
                ).groupby("SourceIdentity", sort=False).size()
            )
            source_identity_groups_df["payload_variants"] = (
                source_identity_groups_df["SourceIdentity"]
                .map(exact_variants).astype("int64")
            )
        conflict_rows = source_identity_groups_df.loc[
            source_identity_groups_df["payload_variants"].gt(1)
        ].head(20)
        if not conflict_rows.empty:
            exact_conflict_payloads = exact_payloads.loc[
                exact_payloads["SourceIdentity"].isin(
                    set(conflict_rows["SourceIdentity"].astype(str))
                )
            ].copy()
        duplicate_identity_groups = int(len(source_identity_groups_df))
        equivalent_duplicate_rows = int(
            (source_identity_groups_df.loc[
                source_identity_groups_df["payload_variants"].eq(1),
                "source_copies",
            ] - 1).sum()
        ) if not source_identity_groups_df.empty else 0

    if not conflict_rows.empty:
        examples_list = []
        for row in conflict_rows.itertuples():
            rows = exact_conflict_payloads.loc[
                exact_conflict_payloads["SourceIdentity"].eq(row.SourceIdentity)
            ]
            differences = []
            for column in IDENTITY_EQUIVALENCE_FIELDS:
                variants = rows[column].drop_duplicates().tolist()
                if len(variants) > 1:
                    rendered = [repr(value)[:80] for value in variants[:3]]
                    differences.append(f"{column}={rendered}")
            examples_list.append(
                f"{row.SourceIdentity} copies={int(row.source_copies)} "
                f"payload_variants={int(row.payload_variants)} "
                f"files={row.equivalent_source_files} "
                f"differences={','.join(differences[:10])}"
            )
        examples = "; ".join(examples_list)
        raise SourceIdentityConflictError(
            "Conflicting canonical payloads reuse the same SourceIdentity; "
            "no copy was selected. " + examples
        )

    if mode == "contaminants" and unreviewed_category_counts:
        examples = "; ".join(
            f"DeterminandCode={code!r}, Test={test!r}, records={count:,}"
            for (code, test), count in sorted(
                unreviewed_category_counts.items(), key=lambda item: str(item[0])
            )[:100]
        )
        raise RuntimeError(
            "SCIENTIFIC CATEGORY REVIEW REQUIRED: retained analytes lack a "
            "reviewed category decision. " + examples
        )

    if not tmp_csv.exists() or total_clean_pre_dedup == 0:
        raise ValueError("No rows survived processing.")

    # Save crosswalk/source metadata now; these do not depend on finalisation.
    unit_crosswalk_df = pd.DataFrame([
        {
            "source_format": sf,
            "raw_unit": raw_u,
            "canonical_unit_before_global_conversion": can_u,
            "policy_status": policy_status,
            "records_seen": n,
        }
        for (sf, raw_u, can_u, policy_status), n in unit_crosswalk_counts.items()
    ])
    if not unit_crosswalk_df.empty:
        unit_crosswalk_df = unit_crosswalk_df.sort_values(
            ["source_format", "records_seen"], ascending=[True, False]
        )

    source_df = pd.DataFrame(source_summary)

    out_units = run_dir / "EA_unit_crosswalk_v2.csv"
    out_sources = run_dir / "EA_source_summary_v2.csv"
    out_schema = run_dir / "EA_schema_crosswalk_v2.csv"
    out_metadata = run_dir / "EA_metadata_v2.csv"
    out_stations = run_dir / "EA_stations_v2.csv"
    out_manifest = run_dir / "EA_source_manifest_v2.csv"
    out_dedup_audit = run_dir / "EA_source_identity_dedup_v2.csv"

    unit_crosswalk_df.to_csv(out_units, index=False)
    source_df.to_csv(out_sources, index=False)
    source_manifest.to_csv(out_manifest, index=False)
    if using_duckdb:
        dedup_sql = str(out_dedup_audit).replace("'", "''")
        con.execute(
            "COPY (SELECT *, CASE WHEN payload_variants=1 THEN "
            "'equivalent_same_source_record' ELSE 'CONFLICT' END AS reason "
            "FROM source_identity_groups ORDER BY SourceIdentity) "
            f"TO '{dedup_sql}' (FORMAT CSV, HEADER true)"
        )
    else:
        dedup_frame = source_identity_groups_df.copy()
        dedup_frame["reason"] = np.where(
            dedup_frame["payload_variants"].eq(1),
            "equivalent_same_source_record",
            "CONFLICT",
        )
        dedup_frame.sort_values("SourceIdentity").to_csv(
            out_dedup_audit, index=False
        )

    schema_crosswalk = pd.DataFrame([
        ("Sampling point code", "sample.samplingPoint.notation", "samplingPoint.notation", "SamplingPointCode"),
        ("Sampling point label", "sample.samplingPoint.label", "samplingPoint.prefLabel", "Sampling Point"),
        ("Date/time", "sample.sampleDateTime", "phenomenonTime", "Date"),
        ("Sample material", "sample.sampledMaterialType.label", "sampleMaterialType", "Type"),
        ("Determinand code", "determinand.notation", "determinand.notation", "DeterminandCode"),
        ("Determinand name", "determinand.definition", "determinand.prefLabel", "Test"),
        ("Result", "result", "result (may contain qualifier prefix)", "result"),
        ("Result qualifier", "resultQualifier.notation", "parsed from result prefix", "ResultQualifier"),
        ("Unit", "determinand.unit.label", "unit", "Unit"),
        ("Coordinates", "Easting/Northing", "longitude/latitude", "Longitude/Latitude"),
        ("Record ID", "@id", "id", "internal provenance"),
        ("Sample ID", "parsed from @id", "parsed from id", "internal de-duplication"),
        ("Region", "not supplied", "samplingPoint.region", "stations metadata"),
        ("Area", "not supplied", "samplingPoint.area", "stations metadata"),
        ("Sub-area", "not supplied", "samplingPoint.subArea", "stations metadata"),
    ], columns=["concept", "legacy_field", "api_field", "v2_field"])
    schema_crosswalk.to_csv(out_schema, index=False)

    # ==================================================================
    # FINALISATION — DUCKDB WHEN AVAILABLE, PANDAS FALLBACK OTHERWISE
    # ==================================================================

    # The primary scientific dataset is intentionally concise. Source and EA
    # identifiers remain available in the standalone provenance output and are
    # still used internally for validation, deterministic ordering and safe
    # source-identity deduplication.
    main_columns = [
        "Sampling Point", "Type", "Date", "Test", "result",
        "ResultQualifier", "Unit", "Season", "SourceYear",
        "Latitude", "Longitude",
    ]
    if category_map:
        main_columns.append("Category")

    provenance_columns = [
        "RecordID", "SamplingPointCode", "DeterminandCode",
        "Sampling Point", "Type", "Date", "Test", "result",
        "ResultQualifier", "Unit", "Season", "SourceYear",
        "Latitude", "Longitude", "SourceFormat", "SourceFile",
        "SourceRecordID", "SampleID", "ObservationKey", "RawUnit",
        "UnitPolicyStatus", "UnitConversionCode", "SourceIdentity",
        "SourceCopyCount", "EquivalentSourceFiles",
    ]
    if category_map:
        provenance_columns.insert(14, "Category")
    provenance_aliases = {
        "RecordID": "record_id",
        "SamplingPointCode": "sampling_point_code",
        "DeterminandCode": "determinand_code",
        "Sampling Point": "sampling_point",
        "Type": "type",
        "Date": "date",
        "Test": "test",
        "result": "result",
        "ResultQualifier": "result_qualifier",
        "Unit": "unit",
        "Season": "season",
        "SourceYear": "source_year",
        "Latitude": "latitude",
        "Longitude": "longitude",
        "Category": "category",
        "SourceFormat": "source_format",
        "SourceFile": "source_file",
        "SourceRecordID": "source_record_id",
        "SampleID": "sample_id",
        "ObservationKey": "observation_key",
        "RawUnit": "raw_unit",
        "UnitPolicyStatus": "unit_policy_status",
        "UnitConversionCode": "unit_conversion_code",
        "SourceIdentity": "source_identity",
        "SourceCopyCount": "source_copy_count",
        "EquivalentSourceFiles": "equivalent_source_files",
    }

    final_df: Any = None

    if using_duckdb:
        log("Finalising with DuckDB (SourceIdentity-safe, out-of-core) ...")
        tmp_csv_sql = str(tmp_csv).replace("'", "''")
        con.execute(
            f"CREATE VIEW staged_text AS SELECT * FROM read_csv_auto("
            f"'{tmp_csv_sql}', header=true, all_varchar=true, "
            "sample_size=200000, ignore_errors=false)"
        )

        staging_columns = [
            "SamplingPointCode", "Sampling Point", "Type", "Date",
            "DeterminandCode", "Test", "result", "ResultQualifier", "Unit",
            "Season", "SourceYear", "Latitude", "Longitude",
            "Region", "Area", "SubArea",
            "SamplingPointStatus", "SamplingPointType", "SamplingPurpose",
            "SourceFormat", "SourceFile", "SourceRecordID", "SampleID",
            "ObservationKey", "SourceIdentity", "RawUnit", "UnitPolicyStatus",
            "UnitConversionCode",
        ]
        if category_map:
            staging_columns.append("Category")
        staging_types = {
            "Date": "TIMESTAMP",
            "result": "DOUBLE",
            "SourceYear": "INTEGER",
            "Latitude": "DOUBLE",
            "Longitude": "DOUBLE",
        }
        typed_projection = ", ".join(
            f'CAST("{column}" AS {staging_types.get(column, "VARCHAR")}) '
            f'AS "{column}"'
            for column in staging_columns
        )
        con.execute(f"CREATE TABLE staged AS SELECT {typed_projection} FROM staged_text")
        con.execute("""
            CREATE TABLE final_internal AS
            SELECT * EXCLUDE (_rn)
            FROM (
                SELECT s.*,
                       'HS2-' || sha256(s.SourceIdentity) AS RecordID,
                       coalesce(g.source_copies, 1) AS SourceCopyCount,
                       coalesce(g.equivalent_source_files, s.SourceFile)
                           AS EquivalentSourceFiles,
                       row_number() OVER (
                           PARTITION BY s.SourceIdentity
                           ORDER BY s.SourceFile, s.SourceRecordID
                       ) AS _rn
                FROM staged s
                LEFT JOIN source_identity_groups g USING (SourceIdentity)
            )
            WHERE _rn = 1
        """)

        final_rows = int(con.execute("SELECT count(*) FROM final_internal").fetchone()[0])
        duplicates_removed = total_clean_pre_dedup - final_rows
        record_ids = int(con.execute(
            "SELECT count(DISTINCT RecordID) FROM final_internal"
        ).fetchone()[0])
        if record_ids != final_rows:
            raise RuntimeError("Stable RecordID collision detected; publication aborted")
        date_min, date_max = con.execute(
            'SELECT min("Date"), max("Date") FROM final_internal'
        ).fetchone()

        start_year = pd.Timestamp(date_min).year
        end_year = pd.Timestamp(date_max).year
        out_csv = run_dir / f"EA_clean_{start_year}_{end_year}_{mode}_v2.csv"
        out_pq = run_dir / f"EA_clean_{start_year}_{end_year}_{mode}_v2.parquet"
        out_provenance = run_dir / f"EA_provenance_{start_year}_{end_year}_{mode}_v2.csv"

        quoted = ", ".join(f'"{c}"' for c in main_columns)
        provenance_quoted = ", ".join(
            f'"{column}" AS "{provenance_aliases[column]}"'
            for column in provenance_columns
        )
        order_sql = ", ".join(
            f'"{column}" NULLS LAST' for column in FINAL_ORDER_COLUMNS + ["RecordID"]
        )
        out_csv_sql = str(out_csv).replace("'", "''")
        out_pq_sql = str(out_pq).replace("'", "''")
        out_provenance_sql = str(out_provenance).replace("'", "''")

        con.execute(
            f'COPY (SELECT {quoted} FROM final_internal ORDER BY {order_sql}) '
            f"TO '{out_csv_sql}' (FORMAT CSV, HEADER true)"
        )
        con.execute(
            f'COPY (SELECT {quoted} FROM final_internal ORDER BY {order_sql}) '
            f"TO '{out_pq_sql}' (FORMAT PARQUET, COMPRESSION ZSTD)"
        )
        con.execute(
            f'COPY (SELECT {provenance_quoted} FROM final_internal '
            f'ORDER BY {order_sql}) TO \'{out_provenance_sql}\' '
            "(FORMAT CSV, HEADER true)"
        )

        # Deterministic station ranking: valid coordinates, completeness, API,
        # recency, then stable provenance. Coordinate variants remain visible.
        out_stations_sql = str(out_stations).replace("'", "''")
        con.execute(f"""
            COPY (
                WITH scored AS (
                    SELECT *,
                           coalesce(nullif(trim(SamplingPointCode), ''),
                                    nullif(trim("Sampling Point"), ''), RecordID)
                               AS station_key,
                           CASE WHEN Latitude IS NOT NULL AND Longitude IS NOT NULL
                                THEN 1 ELSE 0 END AS valid_coordinates,
                           (CASE WHEN "Sampling Point" IS NOT NULL THEN 1 ELSE 0 END +
                            CASE WHEN Region IS NOT NULL THEN 1 ELSE 0 END +
                            CASE WHEN Area IS NOT NULL THEN 1 ELSE 0 END +
                            CASE WHEN SubArea IS NOT NULL THEN 1 ELSE 0 END +
                            CASE WHEN SamplingPointStatus IS NOT NULL THEN 1 ELSE 0 END +
                            CASE WHEN SamplingPointType IS NOT NULL THEN 1 ELSE 0 END)
                               AS metadata_completeness
                    FROM final_internal
                ), station_stats AS (
                    SELECT station_key, count(*) AS candidate_record_count,
                           count(DISTINCT CASE WHEN valid_coordinates=1 THEN
                               concat(CAST(Latitude AS VARCHAR), '|',
                                      CAST(Longitude AS VARCHAR)) END)
                               AS coordinate_variant_count
                    FROM scored GROUP BY station_key
                ), ranked AS (
                    SELECT *, row_number() OVER (
                        PARTITION BY station_key
                        ORDER BY valid_coordinates DESC,
                                 metadata_completeness DESC,
                                 CASE WHEN SourceFormat='api' THEN 0 ELSE 1 END,
                                 Date DESC NULLS LAST, SourceFormat,
                                 SourceRecordID, RecordID
                    ) AS station_rank
                    FROM scored
                )
                SELECT r.SamplingPointCode, r."Sampling Point", r.Latitude,
                       r.Longitude, r.Region, r.Area, r.SubArea,
                       r.SamplingPointStatus, r.SamplingPointType,
                       r.RecordID AS SelectedRecordID,
                       r.SourceFormat AS SelectedSourceFormat,
                       r.SourceFile AS SelectedSourceFile,
                       r.SourceRecordID AS SelectedSourceRecordID,
                       r.valid_coordinates, r.metadata_completeness,
                       s.candidate_record_count, s.coordinate_variant_count
                FROM ranked r JOIN station_stats s USING (station_key)
                WHERE r.station_rank=1
                ORDER BY r.SamplingPointCode NULLS LAST,
                         r."Sampling Point" NULLS LAST, r.RecordID
            ) TO '{out_stations_sql}' (FORMAT CSV, HEADER true)
        """)

        # Summary queries.
        unique_sites = int(con.execute(
            'SELECT count(DISTINCT "SamplingPointCode") FROM final_internal '
            'WHERE "SamplingPointCode" IS NOT NULL'
        ).fetchone()[0])
        unique_tests = int(con.execute(
            'SELECT count(DISTINCT "Test") FROM final_internal'
        ).fetchone()[0])
        unique_types = int(con.execute(
            'SELECT count(DISTINCT "Type") FROM final_internal'
        ).fetchone()[0])
        unique_units = int(con.execute(
            'SELECT count(DISTINCT "Unit") FROM final_internal'
        ).fetchone()[0])
        coordinate_count = int(con.execute(
            'SELECT count(*) FROM final_internal '
            'WHERE "Latitude" IS NOT NULL AND "Longitude" IS NOT NULL'
        ).fetchone()[0])
        qualifier_count = int(con.execute(
            "SELECT count(*) FROM final_internal WHERE "
            "\"ResultQualifier\" IS NOT NULL AND trim(\"ResultQualifier\") <> ''"
        ).fetchone()[0])
        rows_per_year = con.execute("""
            SELECT "SourceYear" AS SourceYear, count(*) AS rows
            FROM final_internal
            GROUP BY "SourceYear"
            ORDER BY "SourceYear"
        """).fetchdf()

    else:
        log("Finalising bounded small run with pandas (SourceIdentity-safe) ...")

        df_all = pd.read_csv(
            tmp_csv,
            low_memory=False,
            dtype=str,
            keep_default_na=False,
            na_values=[""],
        )
        string_columns = [
            column for column in df_all.columns
            if column not in {"Date", "result", "SourceYear", "Latitude",
                              "Longitude"}
        ]
        for column in string_columns:
            df_all[column] = df_all[column].astype("string")
        df_all["Date"] = pd.to_datetime(
            df_all["Date"], errors="raise", format="mixed"
        )
        df_all["result"] = pd.to_numeric(df_all["result"], errors="raise").astype("float64")
        df_all["SourceYear"] = pd.to_numeric(
            df_all["SourceYear"], errors="raise"
        ).astype("Int64")
        for column in ["Latitude", "Longitude"]:
            df_all[column] = pd.to_numeric(df_all[column], errors="coerce").astype("float64")
        if not source_identity_groups_df.empty:
            df_all = df_all.merge(
                source_identity_groups_df[[
                    "SourceIdentity", "source_copies", "equivalent_source_files"
                ]],
                on="SourceIdentity",
                how="left",
                validate="many_to_one",
            )
        else:
            df_all["source_copies"] = pd.NA
            df_all["equivalent_source_files"] = pd.NA
        df_all["SourceCopyCount"] = pd.to_numeric(
            df_all["source_copies"], errors="coerce"
        ).fillna(1).astype("Int64")
        df_all["EquivalentSourceFiles"] = (
            df_all["equivalent_source_files"].astype("string")
            .fillna(df_all["SourceFile"].astype("string"))
        )
        df_all = df_all.drop(
            columns=["source_copies", "equivalent_source_files"]
        )
        df_all = df_all.sort_values(
            ["SourceIdentity", "SourceFile", "SourceRecordID"],
            kind="mergesort",
            na_position="last",
        )
        d0 = len(df_all)
        df_all = df_all.drop_duplicates(subset=["SourceIdentity"], keep="first")
        duplicates_removed = d0 - len(df_all)
        df_all["RecordID"] = df_all["SourceIdentity"].map(
            lambda value: "HS2-" + hashlib.sha256(
                str(value).encode("utf-8")
            ).hexdigest()
        ).astype("string")
        if df_all["RecordID"].nunique(dropna=False) != len(df_all):
            raise RuntimeError("Stable RecordID collision detected; publication aborted")
        df_all = df_all.sort_values(
            FINAL_ORDER_COLUMNS + ["RecordID"],
            kind="mergesort",
            na_position="last",
        )
        final_rows = len(df_all)
        date_min = df_all["Date"].min()
        date_max = df_all["Date"].max()

        start_year = int(pd.Timestamp(date_min).year)
        end_year = int(pd.Timestamp(date_max).year)
        out_csv = run_dir / f"EA_clean_{start_year}_{end_year}_{mode}_v2.csv"
        out_pq = run_dir / f"EA_clean_{start_year}_{end_year}_{mode}_v2.parquet"
        out_provenance = run_dir / f"EA_provenance_{start_year}_{end_year}_{mode}_v2.csv"

        final_df = df_all.copy()
        final_df[main_columns].to_csv(out_csv, index=False)
        final_df[main_columns].to_parquet(
            out_pq, index=False, compression="zstd", engine="pyarrow"
        )
        final_df[provenance_columns].rename(
            columns=provenance_aliases
        ).to_csv(out_provenance, index=False)

        station_sort = final_df.copy()
        station_sort["_station_key"] = (
            station_sort["SamplingPointCode"].astype("string")
            .fillna(station_sort["Sampling Point"].astype("string"))
            .fillna(station_sort["RecordID"].astype("string"))
        )
        station_sort["valid_coordinates"] = (
            station_sort["Latitude"].notna() & station_sort["Longitude"].notna()
        )
        metadata_fields = [
            "Sampling Point", "Region", "Area", "SubArea",
            "SamplingPointStatus", "SamplingPointType",
        ]
        station_sort["metadata_completeness"] = station_sort[
            metadata_fields
        ].notna().sum(axis=1)
        station_sort["_api_priority"] = np.where(
            station_sort["SourceFormat"].eq("api"), 0, 1
        )
        station_sort["candidate_record_count"] = station_sort.groupby(
            "_station_key"
        )["RecordID"].transform("size")
        coordinate_key = (
            station_sort["Latitude"].astype("string") + "|" +
            station_sort["Longitude"].astype("string")
        ).where(station_sort["valid_coordinates"])
        station_sort["coordinate_variant_count"] = coordinate_key.groupby(
            station_sort["_station_key"]
        ).transform("nunique").fillna(0).astype("Int64")
        station_sort = station_sort.sort_values(
            ["_station_key", "valid_coordinates", "metadata_completeness",
             "_api_priority", "Date", "SourceFormat", "SourceRecordID", "RecordID"],
            ascending=[True, False, False, True, False, True, True, True],
            kind="mergesort",
            na_position="last",
        )
        stations_df = station_sort.drop_duplicates("_station_key", keep="first")[[
            "SamplingPointCode", "Sampling Point", "Latitude", "Longitude",
            "Region", "Area", "SubArea", "SamplingPointStatus",
            "SamplingPointType", "RecordID", "SourceFormat", "SourceFile",
            "SourceRecordID", "valid_coordinates", "metadata_completeness",
            "candidate_record_count", "coordinate_variant_count",
        ]]
        stations_df = stations_df.rename(columns={
            "RecordID": "SelectedRecordID",
            "SourceFormat": "SelectedSourceFormat",
            "SourceFile": "SelectedSourceFile",
            "SourceRecordID": "SelectedSourceRecordID",
        }).sort_values(
            ["SamplingPointCode", "Sampling Point", "SelectedRecordID"],
            kind="mergesort", na_position="last",
        )
        stations_df.to_csv(out_stations, index=False)

        unique_sites = int(final_df["SamplingPointCode"].nunique(dropna=True))
        unique_tests = int(final_df["Test"].nunique(dropna=True))
        unique_types = int(final_df["Type"].nunique(dropna=True))
        unique_units = int(final_df["Unit"].nunique(dropna=True))
        coordinate_count = int(
            (final_df["Latitude"].notna() & final_df["Longitude"].notna()).sum()
        )
        qualifier_count = int(
            final_df["ResultQualifier"].astype("string").fillna("").str.strip().ne("").sum()
        )
        rows_per_year = (
            final_df.groupby("SourceYear").size().reset_index(name="rows")
            .sort_values("SourceYear")
        )

    # ==================================================================
    # METADATA FILE
    # ==================================================================

    metadata_rows = [
        ("Sampling Point", "string", "EA sampling-point label supplied by the relevant source record."),
        ("Type", "string", "Sample material/water matrix retained by HydroStream."),
        ("Date", "datetime", "Observation/sample timestamp."),
        ("Test", "string", "Canonical analyte name: legacy determinand.definition or current API determinand.prefLabel."),
        ("result", "float", "Numeric reported value or numeric reporting limit; interpret with ResultQualifier."),
        ("ResultQualifier", "string", "Qualifier such as <, <=, > or >=. Blank for unqualified numeric results."),
        ("Unit", "string", "Canonical unit where a reviewed rule exists; otherwise the unreviewed raw label is retained and identified in provenance."),
        ("Season", "string", "Winter, Spring, Summer or Autumn derived from Date."),
        ("SourceYear", "integer", "Calendar year derived from Date."),
        ("Latitude", "float", "EPSG:4326 latitude. Legacy BNG coordinates converted; API latitude used directly."),
        ("Longitude", "float", "EPSG:4326 longitude. Legacy BNG coordinates converted; API longitude used directly."),
    ]
    if category_map:
        metadata_rows.append(
            ("Category", "string", "Category from the curated test-category workbook; otherwise uncategorized.")
        )
    pd.DataFrame(metadata_rows, columns=["variable", "type", "description"]).to_csv(
        out_metadata, index=False
    )

    # ==================================================================
    # STATISTICS
    # ==================================================================

    out_stats = run_dir / f"EA_statistics_{start_year}_{end_year}_{mode}_v2.xlsx"
    stats_output: Optional[Path] = None

    if generate_stats:
        log("Generating statistics workbook ...")
        try:
            if using_duckdb:
                test_stats = con.execute("""
                    SELECT "Test" AS Test, "Unit" AS Unit,
                           count(*) AS count,
                           min(result) AS min,
                           max(result) AS max,
                           avg(result) AS mean,
                           median(result) AS median,
                           stddev_samp(result) AS std,
                           quantile_cont(result, 0.10) AS p10,
                           quantile_cont(result, 0.25) AS p25,
                           quantile_cont(result, 0.75) AS p75,
                           quantile_cont(result, 0.90) AS p90
                    FROM final_internal
                    GROUP BY "Test", "Unit"
                    ORDER BY "Test", "Unit"
                """).fetchdf()
                type_stats = con.execute("""
                    SELECT "Type" AS Type, "Test" AS Test,
                           count(*) AS count,
                           avg(result) AS mean,
                           median(result) AS median,
                           stddev_samp(result) AS std
                    FROM final_internal
                    GROUP BY "Type", "Test"
                    ORDER BY "Type", "Test"
                """).fetchdf()
                season_stats = con.execute("""
                    SELECT "Season" AS Season, "Test" AS Test,
                           count(*) AS count,
                           avg(result) AS mean,
                           median(result) AS median
                    FROM final_internal
                    GROUP BY "Season", "Test"
                    ORDER BY "Season", "Test"
                """).fetchdf()
                category_stats = (
                    con.execute("""
                        SELECT "Category" AS Category, count(*) AS rows
                        FROM final_internal GROUP BY "Category" ORDER BY rows DESC
                    """).fetchdf()
                    if category_map else None
                )
            else:
                grouped = final_df.groupby(["Test", "Unit"])["result"]
                test_stats = grouped.agg([
                    "count", "min", "max", "mean", "median", "std",
                    ("p10", lambda x: x.quantile(.10)),
                    ("p25", lambda x: x.quantile(.25)),
                    ("p75", lambda x: x.quantile(.75)),
                    ("p90", lambda x: x.quantile(.90)),
                ]).reset_index()
                type_stats = final_df.groupby(["Type", "Test"])["result"].agg(
                    ["count", "mean", "median", "std"]
                ).reset_index()
                season_stats = final_df.groupby(["Season", "Test"])["result"].agg(
                    ["count", "mean", "median"]
                ).reset_index()
                category_stats = (
                    final_df.groupby("Category").size().reset_index(name="rows")
                    .sort_values("rows", ascending=False)
                    if category_map else None
                )

            coverage = pd.DataFrame([
                ("Final rows", final_rows),
                ("Unique sampling-point codes", unique_sites),
                ("Unique tests", unique_tests),
                ("Unique types", unique_types),
                ("Unique units", unique_units),
                ("Date range start", str(date_min)),
                ("Date range end", str(date_max)),
                ("Rows with coordinates", coordinate_count),
                ("Rows with result qualifier", qualifier_count),
                ("Duplicates removed", duplicates_removed),
                ("Mode", mode.upper()),
            ], columns=["Metric", "Value"])

            drop_audit = pd.DataFrame(
                list(drop_counts.items()) + [("duplicates_removed", duplicates_removed)],
                columns=["QA_step", "rows"],
            )

            with pd.ExcelWriter(out_stats, engine="openpyxl") as writer:
                coverage.to_excel(writer, sheet_name="Coverage", index=False)
                test_stats.to_excel(writer, sheet_name="Test_Statistics", index=False)
                type_stats.to_excel(writer, sheet_name="Type_Test_Stats", index=False)
                season_stats.to_excel(writer, sheet_name="Seasonal_Stats", index=False)
                rows_per_year.to_excel(writer, sheet_name="Rows_Per_Year", index=False)
                drop_audit.to_excel(writer, sheet_name="QA_Drop_Audit", index=False)
                source_df.to_excel(writer, sheet_name="Source_Files", index=False)
                unit_crosswalk_df.to_excel(writer, sheet_name="Unit_Crosswalk", index=False)
                schema_crosswalk.to_excel(writer, sheet_name="Schema_Crosswalk", index=False)
                if category_stats is not None:
                    category_stats.to_excel(writer, sheet_name="Rows_Per_Category", index=False)

            stats_output = out_stats
        except Exception as exc:
            raise RuntimeError(
                f"Requested statistics workbook failed; run not published: {exc}"
            ) from exc

    # ==================================================================
    # QA REPORT
    # ==================================================================

    out_qa = run_dir / f"EA_qa_report_{mode}_v2.html"
    qa_output: Optional[Path] = None

    if generate_qa_report:
        log("Generating QA report ...")
        try:
            if using_duckdb:
                top_types = con.execute("""
                    SELECT "Type" AS Type, count(*) AS rows
                    FROM final_internal GROUP BY "Type" ORDER BY rows DESC LIMIT 15
                """).fetchdf()
                top_tests = con.execute("""
                    SELECT "Test" AS Test, count(*) AS rows
                    FROM final_internal GROUP BY "Test" ORDER BY rows DESC LIMIT 20
                """).fetchdf()
                top_units = con.execute("""
                    SELECT "Unit" AS Unit, count(*) AS rows
                    FROM final_internal GROUP BY "Unit" ORDER BY rows DESC LIMIT 25
                """).fetchdf()
            else:
                top_types = final_df["Type"].value_counts().head(15).rename_axis("Type").reset_index(name="rows")
                top_tests = final_df["Test"].value_counts().head(20).rename_axis("Test").reset_index(name="rows")
                top_units = final_df["Unit"].value_counts().head(25).rename_axis("Unit").reset_index(name="rows")

            def _html_rows(frame: pd.DataFrame, label: str) -> str:
                lines = []
                for _, row in frame.iterrows():
                    n = int(row["rows"])
                    pct = n / final_rows * 100 if final_rows else 0
                    lines.append(
                        f"<tr><td>{html.escape(str(row[label]))}</td>"
                        f"<td>{n:,}</td><td>{pct:.2f}%</td></tr>"
                    )
                return "\n".join(lines)

            drop_html = "\n".join(
                f"<tr><td>{html.escape(k)}</td><td>{v:,}</td></tr>"
                for k, v in drop_counts.items()
            )
            source_html = "\n".join(
                f"<tr><td>{html.escape(str(r['source_format']))}</td>"
                f"<td>{html.escape(str(r['source_file']))}</td>"
                f"<td>{int(r['raw_rows_read']):,}</td>"
                f"<td>{int(r['clean_rows_pre_dedup']):,}</td></tr>"
                for _, r in source_df.iterrows()
            )

            report = f"""<!doctype html>
<html><head><meta charset="utf-8"><title>HydroStream V2 QA</title>
<style>
body{{font-family:Arial,sans-serif;margin:32px;color:#222}}
h1{{color:#255f49}}h2{{margin-top:28px;border-bottom:2px solid #8fbd45}}
table{{border-collapse:collapse;width:100%;margin:12px 0 24px}}
th,td{{border:1px solid #d4d8db;padding:7px 9px;text-align:left}}
th{{background:#f1f5f2}}.note{{background:#f6f8f7;padding:12px;border-left:4px solid #8fbd45}}
</style></head><body>
<h1>HydroStream V2 — QA Report</h1>
<p><b>Mode:</b> {html.escape(mode.upper())}<br>
<b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

<h2>Dataset overview</h2>
<table><tr><th>Metric</th><th>Value</th></tr>
<tr><td>Final rows</td><td>{final_rows:,}</td></tr>
<tr><td>Date range</td><td>{date_min} → {date_max}</td></tr>
<tr><td>Unique sampling-point codes</td><td>{unique_sites:,}</td></tr>
<tr><td>Unique tests</td><td>{unique_tests:,}</td></tr>
<tr><td>Unique water types</td><td>{unique_types:,}</td></tr>
<tr><td>Unique units</td><td>{unique_units:,}</td></tr>
<tr><td>Rows with coordinates</td><td>{coordinate_count:,} ({coordinate_count/final_rows*100:.2f}%)</td></tr>
<tr><td>Rows with qualifiers</td><td>{qualifier_count:,} ({qualifier_count/final_rows*100:.2f}%)</td></tr>
<tr><td>Duplicates removed</td><td>{duplicates_removed:,}</td></tr></table>

<h2>Legacy/API integration</h2>
<div class="note"><b>Transition recovery validation for these exact run
inputs: NOT TESTED.</b> The configured source boundary is still applied as
legacy strictly before {cutover.date()} and API from {cutover.date()} onward.
Use the provenance-linked focused transition audit before making a numerical
recovery claim for a source snapshot.</div>
<table><tr><th>Check</th><th>Count</th></tr>
<tr><td>Legacy rows excluded at/after cutover</td><td>{integration_counts['legacy_rows_excluded_at_cutover']:,}</td></tr>
<tr><td>API rows found before cutover</td><td>{integration_counts['api_rows_before_cutover']:,}</td></tr>
<tr><td>Qualified numeric API results parsed</td><td>{integration_counts['qualified_numeric_results']:,}</td></tr>
<tr><td>Less-than qualifiers</td><td>{integration_counts['less_than_results']:,}</td></tr>
<tr><td>Greater-than qualifiers</td><td>{integration_counts['greater_than_results']:,}</td></tr></table>

<h2>Rows removed during cleaning</h2>
<table><tr><th>Step</th><th>Rows</th></tr>{drop_html}</table>

<h2>Source files</h2>
<table><tr><th>Format</th><th>File</th><th>Raw read</th><th>Clean pre-dedup</th></tr>{source_html}</table>

<h2>Top water types</h2><table><tr><th>Type</th><th>Rows</th><th>%</th></tr>{_html_rows(top_types,'Type')}</table>
<h2>Top tests</h2><table><tr><th>Test</th><th>Rows</th><th>%</th></tr>{_html_rows(top_tests,'Test')}</table>
<h2>Top units</h2><table><tr><th>Unit</th><th>Rows</th><th>%</th></tr>{_html_rows(top_units,'Unit')}</table>
</body></html>"""
            out_qa.write_text(report, encoding="utf-8")
            qa_output = out_qa
        except Exception as exc:
            raise RuntimeError(
                f"Requested QA report failed; run not published: {exc}"
            ) from exc

    # ==================================================================
    # PROVENANCE, END-OF-RUN RAW INTEGRITY, AND TRANSACTIONAL PUBLICATION
    # ==================================================================

    out_raw_integrity = run_dir / "EA_raw_integrity_v2.csv"
    integrity_rows = []
    for row in source_manifest.loc[
        source_manifest["accepted"].fillna(False)
    ].itertuples():
        source_path = Path(row.source_path)
        after_size = int(source_path.stat().st_size)
        after_sha256 = _sha256_file(source_path)
        size_match = after_size == int(row.file_size_bytes)
        sha256_match = after_sha256 == str(row.sha256)
        integrity_rows.append({
            "source_format": row.source_format,
            "source_filename": row.source_filename,
            "before_size_bytes": int(row.file_size_bytes),
            "after_size_bytes": after_size,
            "before_sha256": row.sha256,
            "after_sha256": after_sha256,
            "size_match": size_match,
            "sha256_match": sha256_match,
            "status": "PASS" if size_match and sha256_match else "FAIL",
        })
    raw_integrity_df = pd.DataFrame(integrity_rows)
    raw_integrity_df.to_csv(out_raw_integrity, index=False)
    changed_sources = raw_integrity_df.loc[raw_integrity_df["status"].eq("FAIL")]
    if not changed_sources.empty:
        raise RuntimeError(
            "Raw source changed during the run; publication aborted: "
            + ", ".join(changed_sources["source_filename"].astype(str))
        )

    if cat_path is not None and category_provenance is not None:
        category_after_size = int(cat_path.stat().st_size)
        category_after_sha256 = _sha256_file(cat_path)
        category_provenance["after_file_size_bytes"] = category_after_size
        category_provenance["after_sha256"] = category_after_sha256
        category_provenance["unchanged_during_run"] = bool(
            category_after_size == category_provenance["file_size_bytes"]
            and category_after_sha256 == category_provenance["sha256"]
        )
        if not category_provenance["unchanged_during_run"]:
            raise RuntimeError(
                "Category workbook changed during the run; publication aborted: "
                f"{cat_path}"
            )

    finished_utc = datetime.now(timezone.utc)
    runtime_data = _runtime_provenance(
        selected_finalizer, duckdb,
        category_provenance, capacity,
    )
    runtime_data.update({
        "started_utc": started_utc.isoformat(),
        "finished_utc": finished_utc.isoformat(),
        "arguments": {
            "input_dir": str(root),
            "mode": mode,
            "years": sorted(year_set),
            "chunksize": int(chunksize),
            "min_test_count": int(min_test_count),
            "generate_stats": bool(generate_stats),
            "generate_qa_report": bool(generate_qa_report),
            "save_log": bool(save_log),
            "cutover_date": str(cutover),
            "duckdb_memory_limit": str(duckdb_memory_limit),
            "requested_finalizer": requested_finalizer,
            "selected_finalizer": selected_finalizer,
            "unitless_quantitative_allowlist": [
                [code, test] for code, test in sorted(unitless_allowlist_keys)
            ],
            "pandas_fallback_max_rows": int(pandas_fallback_max_rows),
        },
        "source_manifest_sha256": _sha256_file(out_manifest),
        "raw_integrity": {
            "sources_checked": int(len(raw_integrity_df)),
            "sources_unchanged": int(raw_integrity_df["status"].eq("PASS").sum()),
        },
        "deduplication": {
            "identity": "SourceFormat + SourceRecordID",
            "equivalence_check": (
                "candidate fingerprints followed by null-aware exact comparison "
                "of canonical scientific/provenance fields"
            ),
            "duplicate_identity_groups": duplicate_identity_groups,
            "equivalent_rows_removed": int(duplicates_removed),
            "raw_equivalent_duplicate_rows": equivalent_duplicate_rows,
            "conflicting_identity_groups": 0,
            "observation_key_deletion": False,
        },
    })
    out_runtime = run_dir / "EA_runtime_provenance_v2.json"
    out_runtime.write_text(
        json.dumps(runtime_data, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )

    completion_lines = [
        "",
        "=" * 78,
        "  HYDROSTREAM V2 COMPLETE",
        "=" * 78,
        f"  Final rows            : {final_rows:,}",
        f"  Date range            : {date_min} -> {date_max}",
        f"  Sampling-point codes  : {unique_sites:,}",
        f"  Tests                 : {unique_tests:,}",
        f"  Water types           : {unique_types:,}",
        f"  Units                 : {unique_units:,}",
        f"  Source copies removed : {duplicates_removed:,}",
        f"  Rows with qualifiers  : {qualifier_count:,}",
        f"  Output directory      : {out_dir}",
        f"  Finaliser             : {selected_finalizer}",
        f"  Finished UTC          : {finished_utc.isoformat()}",
        "=" * 78,
    ]
    for line in completion_lines:
        log_buffer.write(line + "\n")

    log_path: Optional[Path] = None
    if save_log:
        log_path = run_dir / f"EA_processing_log_{mode}_v2.txt"
        log_path.write_text(log_buffer.getvalue(), encoding="utf-8")

    if con is not None:
        con.close()
        con = None

    staged_outputs = [
        out_csv, out_pq, out_provenance, out_stations, out_metadata,
        out_units, out_schema, out_sources, out_manifest, out_dedup_audit,
        out_runtime, out_raw_integrity,
    ]
    if stats_output is not None:
        staged_outputs.append(stats_output)
    if qa_output is not None:
        staged_outputs.append(qa_output)
    if log_path is not None:
        staged_outputs.append(log_path)

    publication_pairs = [
        (stage_path, out_dir / stage_path.name) for stage_path in staged_outputs
    ]
    _publish_artifacts_transactionally(publication_pairs, run_dir)
    final_paths = {
        stage_path.name: final_path for stage_path, final_path in publication_pairs
    }

    # Only now is success announced. A failed mandatory writer/publication
    # raises above and the transactional publisher restores prior artifacts.
    for line in completion_lines:
        print(line)

    if duck_tmp_dir != run_dir / "duckdb_tmp" and duck_tmp_dir.exists():
        shutil.rmtree(duck_tmp_dir)
    if run_dir.exists():
        shutil.rmtree(run_dir)

    final_csv = final_paths[out_csv.name]
    final_parquet = final_paths[out_pq.name]
    final_provenance = final_paths[out_provenance.name]
    final_stations = final_paths[out_stations.name]
    final_metadata = final_paths[out_metadata.name]
    final_units = final_paths[out_units.name]
    final_schema = final_paths[out_schema.name]
    final_sources = final_paths[out_sources.name]
    final_manifest = final_paths[out_manifest.name]
    final_dedup_audit = final_paths[out_dedup_audit.name]
    final_runtime = final_paths[out_runtime.name]
    final_raw_integrity = final_paths[out_raw_integrity.name]
    final_stats = final_paths[stats_output.name] if stats_output is not None else None
    final_qa = final_paths[qa_output.name] if qa_output is not None else None
    final_log = final_paths[log_path.name] if log_path is not None else None

    return {
        "final_rows": int(final_rows),
        "output_dir": str(out_dir),
        "csv": str(final_csv),
        "parquet": str(final_parquet),
        "provenance": str(final_provenance),
        "statistics": str(final_stats) if final_stats else None,
        "qa_report": str(final_qa) if final_qa else None,
        "log": str(final_log) if final_log else None,
        "stations": str(final_stations),
        "metadata": str(final_metadata),
        "unit_crosswalk": str(final_units),
        "schema_crosswalk": str(final_schema),
        "source_summary": str(final_sources),
        "source_manifest": str(final_manifest),
        "deduplication_audit": str(final_dedup_audit),
        "runtime_provenance": str(final_runtime),
        "raw_integrity": str(final_raw_integrity),
        "finalizer": selected_finalizer,
        "capacity_preflight": dict(capacity),
        "drop_counts": dict(drop_counts),
        "integration_counts": dict(integration_counts),
        "data_quality": {
            "total_raw_rows_physically_read": int(total_raw),
            "clean_rows_pre_dedup": int(total_clean_pre_dedup),
            "duplicates_removed": int(duplicates_removed),
            "unique_sampling_point_codes": int(unique_sites),
            "unique_tests": int(unique_tests),
            "unique_types": int(unique_types),
            "unique_units": int(unique_units),
            "date_range": (str(date_min), str(date_max)),
            "records_with_coordinates": int(coordinate_count),
            "records_with_qualifier": int(qualifier_count),
            "duplicate_source_identity_groups": duplicate_identity_groups,
            "raw_sources_unchanged": int(len(raw_integrity_df)),
        },
    }


# ============================================================================
# SAFE COMMAND-LINE ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description=(
            "Build a validated HydroStream V2 dataset from an explicit project "
            "root containing legacy_raw/ and api_raw/."
        )
    )
    parser.add_argument(
        "--input-dir", required=True,
        help="Explicit HydroStream project root; current-directory guessing is disabled.",
    )
    parser.add_argument(
        "--mode", choices=["full", "electrochemistry", "contaminants"],
        default="full",
    )
    parser.add_argument("--start-year", type=int, default=2000)
    parser.add_argument("--end-year", type=int, default=2026)
    parser.add_argument("--chunksize", type=int, default=250_000)
    parser.add_argument("--min-test-count", type=int, default=50)
    parser.add_argument("--categories-file")
    parser.add_argument(
        "--finalizer", choices=["duckdb", "pandas"], default="duckdb"
    )
    parser.add_argument("--duckdb-memory-limit", default="6GB")
    parser.add_argument("--temp-dir")
    parser.add_argument("--no-stats", action="store_true")
    parser.add_argument("--no-qa", action="store_true")
    parser.add_argument("--no-log", action="store_true")
    args = parser.parse_args()
    if args.end_year < args.start_year:
        parser.error("--end-year must be greater than or equal to --start-year")

    result = hydrostream(
        input_dir=args.input_dir,
        mode=args.mode,
        categories_file=args.categories_file,
        years=range(args.start_year, args.end_year + 1),
        chunksize=args.chunksize,
        min_test_count=args.min_test_count,
        generate_stats=not args.no_stats,
        generate_qa_report=not args.no_qa,
        save_log=not args.no_log,
        cutover_date="2025-10-13",
        duckdb_memory_limit=args.duckdb_memory_limit,
        finalizer=args.finalizer,
        temp_dir=args.temp_dir,
    )

    print("\n" + "-" * 70)
    print("QUICK SUMMARY")
    print("-" * 70)
    print(f"Final rows : {result['final_rows']:,}")
    print(f"Output dir : {result['output_dir']}")
    print("Files:")
    for key in [
        "csv", "parquet", "provenance", "statistics", "qa_report", "log",
        "stations", "metadata", "unit_crosswalk", "schema_crosswalk",
        "source_summary", "source_manifest", "deduplication_audit",
        "runtime_provenance", "raw_integrity",
    ]:
        if result.get(key):
            print(f"  {key:<16}: {Path(result[key]).name}")
    print("-" * 70)
