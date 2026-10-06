from __future__ import annotations

import argparse
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

DEFAULT_INPUT = Path(__file__).resolve().with_name("EPI Database_CHDN.xlsx")
DEFAULT_OUTPUT = Path(__file__).resolve().with_name("CHDN_EPI_clean.xlsx")
DEFAULT_SOURCE_SHEET = "EPI-Child"
DEFAULT_TARGET_SHEET = "EPI-Child"
DEFAULT_PREGNANCY_SOURCE_SHEET = "EPI-Pregnancy"
DEFAULT_PREGNANCY_TARGET_SHEET = "EPI-Pregnancy"

DOSE_AUTOFILL_MAPPINGS = [
    ("BCG dose", "BCG other (Y/N)", "BCG Age ", "BCG Source"),
    ("OPV first time dose", "OPV first time dose other", "OPV1 Age ", "OPV1 Source"),
    ("OPV second time dose", "OPV second time dose other", "OPV2 Age", "OPV2 Source"),
    ("OPV third time dose", "OPV third time dose other", "OPV3 Age", "OPV3 Source"),
    ("Penta first time dose", "Penta first time dose other", "Penta1 Age", "Penta1 Source"),
    ("Penta second time dose", "Penta second time dose other", "Penta2 Age", "Penta2 Source"),
    ("Penta third time dose", "Penta third time dose other", "Penta3 Age", "Penta3 Source"),
    ("MMR first timeadose", "MMR first time dose other", "MMR1 Age", "MMR1 Source"),
    ("MMR second time dose", "MMR second time dose other", "MMR2 Age", "MMR2 source"),
    ("JE first time dose", "JE first time dose other", "JE1 Age", "JE1 Source"),
    ("JE second time dose", "JE second time dose other", "JE2 Age", "JE2 Source"),
    ("IPV dose", "IPV dose other", "IPV Age", "IPV Source"),
]

COMPLETE_COLUMNS = [
    "CompleteInQ4 2024",
    "CompleteInQ1 2025",
    "CompleteInQ2 2025",
    "CompleteInQ3 2025",
    "CompleteInQ4 2025",
    "CompleteInQ1 2026",
    "CompleteInQ2 2026",
    "CompleteInQ3 2026",
    "CompleteInQ4 2026",
]

QUARTER_WINDOWS = {
    "CompleteInQ4 2024": ("Q4", date(2024, 9, 21), date(2024, 12, 20)),
    "CompleteInQ1 2025": ("Q1", date(2024, 12, 21), date(2025, 3, 20)),
    "CompleteInQ2 2025": ("Q2", date(2025, 3, 21), date(2025, 6, 20)),
    "CompleteInQ3 2025": ("Q3", date(2025, 6, 21), date(2025, 9, 20)),
    "CompleteInQ4 2025": ("Q4", date(2025, 9, 21), date(2025, 12, 20)),
    "CompleteInQ1 2026": ("Q1", date(2025, 12, 21), date(2026, 3, 20)),
    "CompleteInQ2 2026": ("Q2", date(2026, 3, 21), date(2026, 6, 20)),
    "CompleteInQ3 2026": ("Q3", date(2026, 6, 21), date(2026, 9, 20)),
    "CompleteInQ4 2026": ("Q4", date(2026, 9, 21), date(2026, 12, 20)),
}

MIN_DOSE_INTERVAL_DAYS = 28

TARGET_FACTOR = 0.35

REFERENCE_QUARTER_TARGETS_BY_YEAR = {
    2025: {
        "Penta1 under 5-yr-old": [1500.0, 1500.0, 2750.0, 2750.0],
        "Penta3 under 5-yr-old": [1200.0, 1200.0, 2200.0, 2200.0],
        "MMR1 under 5-yr-old": [1350.0, 1350.0, 2475.0, 2475.0],
        "MMR2 under 5-yr-old": [1080.0, 1080.0, 1980.0, 1980.0],
        "Full dose under 5-yr-old": [1100.0, 1100.0, 1400.0, 1400.0],
        "At least one dose under 5-yr-old": [2750.0, 2750.0, 3500.0, 3500.0],
        "Td ALOD": [750.0, 750.0, 1000.0, 1000.0],
        "Td Two Doses": [487.5, 487.5, 650.0, 650.0],
        "Penta3 under 1-yr-old": [600.0, 600.0, 1000.0, 1000.0],
        "MMR1 under 1-yr-old": [525.0, 525.0, 875.0, 875.0],
    },
    2026: {
        "Penta1 under 5-yr-old": [1622.5, 1622.5, 544.5, 544.5],
        "Penta3 under 5-yr-old": [1460.5, 1460.5, 490.0, 490.0],
        "MMR1 under 5-yr-old": [1561.5, 1561.5, 393.0, 393.0],
        "MMR2 under 5-yr-old": [1405.5, 1405.5, 353.5, 353.5],
        "Full dose under 5-yr-old": [1819.0, 1819.0, 88.0, 88.0],
        "At least one dose under 5-yr-old": [3638.0, 3638.0, 960.5, 960.5],
        "Td ALOD": [1036.0, 1036.0, 964.0, 964.0],
        "Td Two Doses": [725.0, 725.0, 775.0, 775.0],
        "Penta3 under 1-yr-old": [679.5, 679.5, 850.0, 850.0],
        "MMR1 under 1-yr-old": [531.0, 531.0, 956.0, 956.0],
    },
}

# Completion rule alias lists as module-level constants so they are built once
# rather than being reconstructed inside every _completion_value() call.
_U1_PRESENCE_RULES: list[tuple[list[str], list[str]]] = [
    (["BCG dose", "BCG"], ["BCG other (Y/N)", "BCG other", "BCG other yn"]),
    (["OPV first time dose", "OPV first timeadose", "OPV1 dose", "OPV1"], ["OPV first time dose other", "OPV1 other", "OPV1 other yn"]),
    (["OPV second time dose", "OPV2 dose", "OPV2"], ["OPV second time dose other", "OPV2 other", "OPV2 other yn"]),
    (["OPV third time dose", "OPV3 dose", "OPV3"], ["OPV third time dose other", "OPV3 other", "OPV3 other yn"]),
    (["Penta first time dose", "Penta1 dose", "Penta1"], ["Penta first time dose other", "Penta1 other", "Penta1 other yn"]),
    (["Penta second time dose", "Penta2 dose", "Penta2"], ["Penta second time dose other", "Penta2 other", "Penta2 other yn"]),
    (["Penta third time dose", "Penta3 dose", "Penta3"], ["Penta third time dose other", "Penta3 other", "Penta3 other yn"]),
    (["MMR first time dose", "MMR first timeadose", "MMR1 dose", "MMR1"], ["MMR first time dose other", "MMR1 other", "MMR1 other yn"]),
]
_U5_PRESENCE_RULES = _U1_PRESENCE_RULES[1:]

_U1_DATE_SOURCE_RULES: list[tuple[list[str], list[str]]] = [
    (["BCG reporting month", "BCG report month"], ["BCG Source", "BCG source", "BCG src"]),
    (["OPV first time dose reporting month", "OPV1 reporting month"], ["OPV1 Source", "OPV1 source", "OPV1 src"]),
    (["OPV second time dose reporting month", "OPV2 reporting month"], ["OPV2 Source", "OPV2 source", "OPV2 src"]),
    (["OPV third time dose reporting month", "OPV3 reporting month"], ["OPV3 Source", "OPV3 source", "OPV3 src"]),
    (["Penta first time dose reporting month", "Penta1 reporting month"], ["Penta1 Source", "Penta1 source", "Penta1 src"]),
    (["Penta second time dose reporting month", "Penta2 reporting month"], ["Penta2 Source", "Penta2 source", "Penta2 src"]),
    (["Penta third time dose reporting month", "Penta3 reporting month"], ["Penta3 Source", "Penta3 source", "Penta3 src"]),
    (["MMR first time dose reporting month", "MMR1 reporting month"], ["MMR1 Source", "MMR1 source", "MMR1 src"]),
    (["MMR second time dose reporting month", "MMR2 reporting month"], ["MMR2 Source", "MMR2 source", "MMR2 src"]),
]
_U5_DATE_SOURCE_RULES = _U1_DATE_SOURCE_RULES[1:]
_COMPLETION_MAX_DATE_RULES = [
    ["OPV third time dose reporting month", "OPV3 reporting month"],
    ["Penta third time dose reporting month", "Penta3 reporting month"],
    ["MMR first time dose reporting month", "MMR1 reporting month"],
]

_FIRST_VISIT_DOSE_SOURCE_RULES: list[tuple[list[str], list[str]]] = [
    (["BCG dose", "BCG"], ["BCG Source", "BCG source", "BCG src"]),
    (["OPV first time dose", "OPV first timeadose", "OPV1 dose", "OPV1"], ["OPV1 Source", "OPV1 source", "OPV1 src"]),
    (["OPV second time dose", "OPV2 dose", "OPV2"], ["OPV2 Source", "OPV2 source", "OPV2 src"]),
    (["OPV third time dose", "OPV3 dose", "OPV3"], ["OPV3 Source", "OPV3 source", "OPV3 src"]),
    (["Penta first time dose", "Penta1 dose", "Penta1"], ["Penta1 Source", "Penta1 source", "Penta1 src"]),
    (["Penta second time dose", "Penta2 dose", "Penta2"], ["Penta2 Source", "Penta2 source", "Penta2 src"]),
    (["Penta third time dose", "Penta3 dose", "Penta3"], ["Penta3 Source", "Penta3 source", "Penta3 src"]),
    (["MMR first time dose", "MMR first timeadose", "MMR1 dose", "MMR1"], ["MMR1 Source", "MMR1 source", "MMR1 src"]),
    (["MMR second time dose", "MMR2 dose", "MMR2"], ["MMR2 source", "MMR2 Source", "MMR2 src"]),
    (["JE first time dose", "JE1 dose", "JE1"], ["JE1 Source", "JE1 source", "JE1 src"]),
    (["JE second time dose", "JE2 dose", "JE2"], ["JE2 Source", "JE2 source", "JE2 src"]),
    (["IPV dose", "IPV"], ["IPV Source", "IPV source", "IPV src"]),
]

_AGE_ALIASES = ["Age at first visit", "Age", "Age in months", "Age at first visit "]
FIRST_VISIT_DATE_HEADER = "first_visit_date"
AGE_AT_FIRST_VISIT_HEADER = "Age at first visit"


def _normalize_col_name(name: str) -> str:
    return "".join(ch for ch in str(name).strip().lower() if ch.isalnum())


def _find_dob_col_idx(ws) -> int | None:
    aliases = {"dob", "dateofbirth", "birthdate", "datebirth"}
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value
        if header is None:
            continue
        norm = _normalize_col_name(header)
        if ("dob" in norm) or (norm in aliases) or ("dateofbirth" in norm):
            return col_idx
    if ws.max_column >= 11:
        return 11
    return None


def _header_col_map(ws) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value
        if header is None:
            continue
        mapping[_normalize_col_name(header)] = col_idx
    return mapping


def _get_or_add_column(ws, header_name: str, header_map: dict[str, int]) -> int:
    key = _normalize_col_name(header_name)
    existing = header_map.get(key)
    if existing is not None:
        return existing
    new_idx = ws.max_column + 1
    ws.cell(row=1, column=new_idx, value=header_name)
    header_map[key] = new_idx
    return new_idx


def _is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def _to_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.date()
    if isinstance(value, (int, float)):
        try:
            parsed = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
            if pd.isna(parsed):
                return None
            return parsed.date()
        except Exception:
            return None
    return None


def _is_yes(value) -> bool:
    return str(value).strip().upper() == "YES"


def _is_no(value) -> bool:
    return str(value).strip().upper() == "NO"


def _derive_source_value(dose_value, other_value) -> str:
    if _is_yes(other_value):
        return "Other"
    if not _is_blank(dose_value):
        return "CHDN"
    if _is_blank(dose_value):
        return "Not received yet"
    return ""


def _calculate_dose_age_value(dob_date, dose_value, other_value):
    try:
        if _is_yes(other_value):
            return 999
        dose_date = _to_date(dose_value)
        if dob_date is not None and dose_date is not None:
            return _datedif_months(dob_date, dose_date)
        return 1111
    except Exception:
        return ""


def _datedif_months(start_date: date, end_date: date) -> int:
    months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)
    if end_date.day < start_date.day:
        months -= 1
    return months


def _precompute_dose_source_indices(header_index_map: dict[str, int]) -> list[tuple]:
    """Pre-compute (dose_idx, source_idx) pairs for first-visit calculations.

    Called once per file so that first-visit calculations no longer call _find_header_index
    inside the per-row loop.
    """
    return [
        (
            _find_header_index(header_index_map, dose_aliases),
            _find_header_index(header_index_map, source_aliases),
        )
        for dose_aliases, source_aliases in _FIRST_VISIT_DOSE_SOURCE_RULES
    ]


def _precompute_completion_indices(header_index_map: dict[str, int]) -> dict:
    """Pre-compute all column indices needed for completion-value calculations.

    Called once per file before the per-row loop to avoid O(rows × quarters × rules)
    string-normalization work.
    """
    u1_presence = [
        (
            _find_header_index(header_index_map, dose_aliases),
            _find_header_index(header_index_map, other_aliases),
        )
        for dose_aliases, other_aliases in _U1_PRESENCE_RULES
    ]
    u1_date_source = [
        (
            _find_header_index(header_index_map, date_aliases),
            _find_header_index(header_index_map, source_aliases),
        )
        for date_aliases, source_aliases in _U1_DATE_SOURCE_RULES
    ]
    return {
        "age": _find_header_index(header_index_map, _AGE_ALIASES),
        "u1_presence": u1_presence,
        "u5_presence": u1_presence[1:],
        "u1_date_source": u1_date_source,
        "u5_date_source": u1_date_source[1:],
        "max_date_cols": [
            _find_header_index(header_index_map, aliases)
            for aliases in _COMPLETION_MAX_DATE_RULES
        ],
    }


def _presence_ok_fast(row_values, dose_idx, other_idx) -> bool:
    """Return True when the dose exists or the paired other cell is YES."""
    n = len(row_values)
    dose_value = row_values[dose_idx] if dose_idx is not None and dose_idx < n else None
    other_value = row_values[other_idx] if other_idx is not None and other_idx < n else None
    return (not _is_blank(dose_value)) or _is_yes(other_value)


def _date_source_ok_fast(row_values, date_idx, source_idx, start: date, end: date) -> bool:
    """Return True when the reporting-month date falls inside the window and Source is CHDN."""
    n = len(row_values)
    date_value = row_values[date_idx] if date_idx is not None and date_idx < n else None
    row_date = _to_date(date_value)
    if row_date is None or not (start <= row_date <= end):
        return False
    source_value = row_values[source_idx] if source_idx is not None and source_idx < n else None
    return str(source_value).strip().upper() == "CHDN"


def _max_date_ok_fast(row_values, date_indices: list[int | None], end: date) -> bool:
    n = len(row_values)
    dates = [
        _to_date(row_values[idx])
        for idx in date_indices
        if idx is not None and idx < n
    ]
    dates = [item for item in dates if item is not None]
    if not dates:
        return True
    return max(dates) <= end


def _compute_first_visit_date(row_values, header_index_map: dict[str, int]) -> "date | None":
    """Return the earliest dose date whose paired Source cell is CHDN."""

    earliest: "date | None" = None
    for dose_aliases, source_aliases in _FIRST_VISIT_DOSE_SOURCE_RULES:
        dose_idx = _find_header_index(header_index_map, dose_aliases)
        source_idx = _find_header_index(header_index_map, source_aliases)
        if dose_idx is None or source_idx is None:
            continue
        source_val = row_values[source_idx] if source_idx < len(row_values) else None
        if str(source_val).strip().upper() != "CHDN":
            continue
        dose_date = _to_date(row_values[dose_idx] if dose_idx < len(row_values) else None)
        if dose_date is None:
            continue
        if earliest is None or dose_date < earliest:
            earliest = dose_date
    return earliest


def _compute_age_at_first_visit(row_values, header_index_map: dict[str, int], dob: "date | None", error_value: int | None = None) -> "int | None":
    """Return age in completed months from DOB to the CHDN first_visit_date.

    Mirrors an Excel-style MINIFS/DATEDIF pattern:
      =DATEDIF(DOB, MINIFS(reporting_month_cols, source_cols, "CHDN"), "M")
    """
    if dob is None:
        return error_value
    first_visit_date = _compute_first_visit_date(row_values, header_index_map)
    if first_visit_date is None:
        return error_value
    return _datedif_months(dob, first_visit_date)


def _compute_first_visit_date_fast(row_values, dose_source_indices: list[tuple]) -> "date | None":
    """Fast variant of first_visit_date calculation using pre-computed pairs."""

    earliest: "date | None" = None
    n = len(row_values)
    for dose_idx, source_idx in dose_source_indices:
        if dose_idx is None or source_idx is None:
            continue
        source_val = row_values[source_idx] if source_idx < n else None
        if str(source_val).strip().upper() != "CHDN":
            continue
        dose_date = _to_date(row_values[dose_idx] if dose_idx < n else None)
        if dose_date is None:
            continue
        if earliest is None or dose_date < earliest:
            earliest = dose_date
    return earliest


def _compute_age_at_first_visit_fast(row_values, dose_source_indices: list[tuple], dob: "date | None", error_value: int | None = None) -> "int | None":
    """Fast variant of _compute_age_at_first_visit using pre-computed pairs.

    Avoids repeated _find_header_index calls inside the per-row hot loop.
    """
    if dob is None:
        return error_value
    first_visit_date = _compute_first_visit_date_fast(row_values, dose_source_indices)
    if first_visit_date is None:
        return error_value
    return _datedif_months(dob, first_visit_date)


def _apply_age_source_autofill(input_file: Path, source_sheet_name: str) -> tuple[Path, dict]:
    wb = load_workbook(input_file)
    if source_sheet_name not in wb.sheetnames:
        raise ValueError(f"Source sheet not found for autofill: {source_sheet_name}")

    ws = wb[source_sheet_name]
    dob_col_idx = _find_dob_col_idx(ws)
    if dob_col_idx is None:
        raise ValueError("DOB column not found in EPI-Child sheet.")

    header_map = _header_col_map(ws)
    stats = {
        "age_filled": 0,
        "first_visit_date_filled": 0,
        "age_at_first_visit_filled": 0,
        "source_filled": 0,
        "age_columns_added": 0,
        "source_columns_added": 0,
        "first_visit_date_columns_added": 0,
    }

    first_visit_date_key = _normalize_col_name(FIRST_VISIT_DATE_HEADER)
    first_visit_date_was_missing = first_visit_date_key not in header_map
    first_visit_date_idx = _get_or_add_column(ws, FIRST_VISIT_DATE_HEADER, header_map)
    if first_visit_date_was_missing:
        stats["first_visit_date_columns_added"] += 1

    age_at_first_visit_key = _normalize_col_name(AGE_AT_FIRST_VISIT_HEADER)
    age_at_first_visit_was_missing = age_at_first_visit_key not in header_map
    age_at_first_visit_idx = _get_or_add_column(ws, AGE_AT_FIRST_VISIT_HEADER, header_map)
    if age_at_first_visit_was_missing:
        stats["age_columns_added"] += 1

    prepared_mappings: list[tuple[int, int, int, int]] = []
    for dose_name, other_name, age_name, source_name in DOSE_AUTOFILL_MAPPINGS:
        dose_idx = header_map.get(_normalize_col_name(dose_name))
        other_idx = header_map.get(_normalize_col_name(other_name))
        if dose_idx is None or other_idx is None:
            continue
        age_key = _normalize_col_name(age_name)
        source_key = _normalize_col_name(source_name)
        age_was_missing = age_key not in header_map
        source_was_missing = source_key not in header_map
        age_idx = _get_or_add_column(ws, age_name, header_map)
        source_idx = _get_or_add_column(ws, source_name, header_map)
        if age_was_missing:
            stats["age_columns_added"] += 1
        if source_was_missing:
            stats["source_columns_added"] += 1
        prepared_mappings.append((dose_idx, other_idx, age_idx, source_idx))

    if not prepared_mappings:
        raise ValueError("No dose/other column pairs found for autofill mapping.")

    dose_source_indices = [
        (dose_idx - 1 if dose_idx is not None else None, source_idx - 1 if source_idx is not None else None)
        for dose_idx, source_idx in _precompute_dose_source_indices(header_map)
    ]

    max_col = ws.max_column
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=max_col, values_only=False):
        row_values = [cell.value for cell in row]
        if not any(not _is_blank(value) for value in row_values):
            continue
        dob_cell = row[dob_col_idx - 1]
        dob_date = _to_date(dob_cell.value)
        for dose_idx, other_idx, _age_idx, source_idx in prepared_mappings:
            dose_cell = row[dose_idx - 1]
            other_cell = row[other_idx - 1]
            source_cell = row[source_idx - 1] if source_idx <= len(row) else None
            dose_raw = dose_cell.value
            other_raw = other_cell.value
            if source_cell is not None:
                try:
                    derived_source = _derive_source_value(dose_raw, other_raw)
                    if derived_source and source_cell.value != derived_source:
                        source_cell.value = derived_source
                        stats["source_filled"] += 1
                except Exception:
                    source_cell.value = ""

        current_row_values = [cell.value for cell in row]
        first_visit_date_cell = row[first_visit_date_idx - 1] if first_visit_date_idx <= len(row) else None
        if first_visit_date_cell is not None and _is_blank(first_visit_date_cell.value):
            first_visit_date_cell.value = _compute_first_visit_date_fast(current_row_values, dose_source_indices)
            stats["first_visit_date_filled"] += 1

        age_at_first_visit_cell = row[age_at_first_visit_idx - 1] if age_at_first_visit_idx <= len(row) else None
        if age_at_first_visit_cell is not None and _is_blank(age_at_first_visit_cell.value):
            first_visit_date = _to_date(first_visit_date_cell.value) if first_visit_date_cell is not None else None
            age_at_first_visit_cell.value = _datedif_months(dob_date, first_visit_date) if dob_date is not None and first_visit_date is not None else 9999
            stats["age_at_first_visit_filled"] += 1
        for dose_idx, other_idx, age_idx, source_idx in prepared_mappings:
            dose_cell = row[dose_idx - 1]
            other_cell = row[other_idx - 1]
            age_cell = row[age_idx - 1] if age_idx <= len(row) else None
            dose_raw = dose_cell.value
            other_raw = other_cell.value
            if age_cell is not None:
                age_cell.value = _calculate_dose_age_value(dob_date, dose_raw, other_raw)
                stats["age_filled"] += 1
    temp_dir = Path(tempfile.mkdtemp(prefix="chdn_clean_copy_"))
    temp_input = temp_dir / input_file.name
    wb.save(temp_input)
    return temp_input, stats


def _find_children_code_column(columns) -> int | None:
    aliases = {"childrencodetccode", "childrencode", "tccode", "tcode"}
    for idx, col in enumerate(columns, start=1):
        if _normalize_col_name(col) in aliases:
            return idx
    return None


def _find_pregnance_code_column(columns) -> int | None:
    aliases = {"pregnancecode", "pregnancycode", "pregnance_code", "pregnancy_code", "pwcode", "pw_code"}
    norm_aliases = {a.replace("_", "") for a in aliases}
    for idx, col in enumerate(columns, start=1):
        if _normalize_col_name(col) in norm_aliases:
            return idx
    return None


def _find_header_index(header_index_map: dict[str, int], aliases: list[str]) -> int | None:
    for alias in aliases:
        if alias in header_index_map:
            return header_index_map[alias]
    normalized_aliases = {_normalize_col_name(alias) for alias in aliases}
    for header_name, index in header_index_map.items():
        if _normalize_col_name(header_name) in normalized_aliases:
            return index
    return None


def _row_value(row_values, header_index_map: dict[str, int], aliases: list[str]):
    index = _find_header_index(header_index_map, aliases)
    if index is None:
        return ""
    return row_values[index]


def _presence_ok(row_values, header_index_map: dict[str, int], dose_aliases: list[str], other_aliases: list[str]) -> bool:
    dose_value = _row_value(row_values, header_index_map, dose_aliases)
    other_value = _row_value(row_values, header_index_map, other_aliases)
    if not _is_blank(dose_value):
        return True
    if not _is_blank(other_value):
        return True
    return False


def _date_source_ok(row_values, header_index_map: dict[str, int], date_aliases: list[str], source_aliases: list[str], start: date, end: date) -> bool:
    row_date = _to_date(_row_value(row_values, header_index_map, date_aliases))
    if row_date is None or not (start <= row_date <= end):
        return False

    # Count the row as completed when the reporting month falls inside the quarter.
    # Source values vary across forms and are not reliable for blocking completion.
    return True


def _completion_value(row_values, column_name: str, precomp: dict, ce_age: "int | None" = None) -> str:
    """Compute the completion label for *column_name* using pre-computed column indices.

    *precomp* must be the dict returned by _precompute_completion_indices().
    Using pre-computed indices avoids calling _find_header_index (with its
    string-normalization scan) for every row × quarter combination.
    """
    quarter_label, start_date, end_date = QUARTER_WINDOWS[column_name]
    year_label = column_name.rsplit(" ", 1)[-1]
    age_num: "float | None" = None
    if ce_age is not None:
        age_num = float(ce_age)
    else:
        age_idx = precomp.get("age")
        if age_idx is not None and age_idx < len(row_values):
            try:
                age_num = float(row_values[age_idx])
            except (TypeError, ValueError):
                age_num = None
    if age_num is not None and age_num <= 11:
        has_all_presence = all(
            _presence_ok_fast(row_values, d_idx, o_idx)
            for d_idx, o_idx in precomp["u1_presence"]
        )
        has_any_date_source = any(
            _date_source_ok_fast(row_values, dt_idx, s_idx, start_date, end_date)
            for dt_idx, s_idx in precomp["u1_date_source"]
        )
        if has_all_presence and has_any_date_source and _max_date_ok_fast(row_values, precomp["max_date_cols"], end_date):
            return f"U1 complete in {quarter_label}_{year_label}"
    if age_num is not None and 11 < age_num <= 59:
        has_all_presence = all(
            _presence_ok_fast(row_values, d_idx, o_idx)
            for d_idx, o_idx in precomp["u5_presence"]
        )
        has_any_date_source = any(
            _date_source_ok_fast(row_values, dt_idx, s_idx, start_date, end_date)
            for dt_idx, s_idx in precomp["u5_date_source"]
        )
        if has_all_presence and has_any_date_source and _max_date_ok_fast(row_values, precomp["max_date_cols"], end_date):
            return f"1-5 complete in {quarter_label}_{year_label}"
    return ""


def _find_source_columns(headers: list) -> list[int]:
    """Return 1-based indices of columns whose headers contain 'source'."""
    source_cols = []
    for idx, header in enumerate(headers, start=1):
        if header is not None and "source" in str(header).strip().lower():
            source_cols.append(idx)
    return source_cols


def _warn_missing_source_columns(headers: list) -> None:
    col_count = len(_find_source_columns(headers))
    if col_count == 0:
        print("WARNING: No source columns found in the output sheet.")
    else:
        print(f"INFO: {col_count} source column(s) verified in the output sheet.")


def _warn_code_quality(sheet_or_values, code_col_idx_or_label, code_label: str | None = None, row_numbers: list[int] | None = None) -> None:
    if code_label is None and row_numbers is None:
        code_values = sheet_or_values
        label = code_col_idx_or_label
    else:
        sheet = sheet_or_values
        code_col_idx = code_col_idx_or_label
        code_values = [sheet.cell(row=row_num, column=code_col_idx).value for row_num in row_numbers or []]
        label = code_label
    normalized_codes = [str(value).strip() for value in code_values if not _is_blank(value)]
    missing_count = sum(1 for value in code_values if _is_blank(value))
    duplicate_codes = sorted({value for value in normalized_codes if normalized_codes.count(value) > 1})
    if missing_count > 0:
        print(f"WARNING: {label} missing. Missing count = {missing_count}")
    else:
        print(f"No missing {label}")
    if duplicate_codes:
        print(f"WARNING: duplication of {label} ({','.join(duplicate_codes)})")


def _verify_cleaned_child_rows(row_values_list: list[tuple], header_index_map: dict[str, int], code_col_idx: int) -> dict[str, int]:
    first_visit_idx = _find_header_index(header_index_map, [FIRST_VISIT_DATE_HEADER])
    dob_idx = _find_header_index(header_index_map, ["DOB", "Date of Birth", "DateOfBirth", "Birth Date", "Birthdate"])

    missing_code_count = 0
    duplicate_code_count = 0
    dob_after_first_visit_count = 0
    later_before_prior_count = 0
    later_while_prior_not_received_count = 0
    interval_under_min_count = 0

    code_values = []
    for row_values in row_values_list:
        code_value = row_values[code_col_idx - 1] if code_col_idx is not None and code_col_idx - 1 < len(row_values) else None
        if _is_blank(code_value):
            missing_code_count += 1
        else:
            code_values.append(str(code_value).strip())

        dob_date = _to_date(row_values[dob_idx]) if dob_idx is not None and dob_idx < len(row_values) else None
        first_visit_date = _to_date(row_values[first_visit_idx]) if first_visit_idx is not None and first_visit_idx < len(row_values) else None
        if dob_date is not None and first_visit_date is not None and dob_date > first_visit_date:
            dob_after_first_visit_count += 1

        def _dose_date(header_aliases: list[str]) -> date | None:
            idx = _find_header_index(header_index_map, header_aliases)
            return _to_date(row_values[idx]) if idx is not None and idx < len(row_values) else None

        def _source_text(header_aliases: list[str]) -> str:
            idx = _find_header_index(header_index_map, header_aliases)
            value = row_values[idx] if idx is not None and idx < len(row_values) else None
            return str(value).strip().upper()

        dose_pairs = [
            ("OPV", ["OPV first time dose"], ["OPV1 Source", "OPV1 source", "OPV1 src"]),
            ("OPV", ["OPV second time dose"], ["OPV2 Source", "OPV2 source", "OPV2 src"]),
            ("OPV", ["OPV third time dose"], ["OPV3 Source", "OPV3 source", "OPV3 src"]),
            ("Penta", ["Penta first time dose"], ["Penta1 Source", "Penta1 source", "Penta1 src"]),
            ("Penta", ["Penta second time dose"], ["Penta2 Source", "Penta2 source", "Penta2 src"]),
            ("Penta", ["Penta third time dose"], ["Penta3 Source", "Penta3 source", "Penta3 src"]),
            ("MMR", ["MMR first timeadose", "MMR first time dose"], ["MMR1 Source", "MMR1 source", "MMR1 src"]),
            ("MMR", ["MMR second time dose"], ["MMR2 source", "MMR2 Source", "MMR2 src"]),
        ]

        ordered_sequences = [
            (["OPV first time dose"], ["OPV second time dose"], ["OPV1 Source", "OPV1 source", "OPV1 src"]),
            (["OPV second time dose"], ["OPV third time dose"], ["OPV2 Source", "OPV2 source", "OPV2 src"]),
            (["Penta first time dose"], ["Penta second time dose"], ["Penta1 Source", "Penta1 source", "Penta1 src"]),
            (["Penta second time dose"], ["Penta third time dose"], ["Penta2 Source", "Penta2 source", "Penta2 src"]),
            (["MMR first timeadose", "MMR first time dose"], ["MMR second time dose"], ["MMR1 Source", "MMR1 source", "MMR1 src"]),
        ]

        for prior_aliases, later_aliases, prior_source_aliases in ordered_sequences:
            prior_date = _dose_date(prior_aliases)
            later_date = _dose_date(later_aliases)
            prior_source = _source_text(prior_source_aliases)
            if prior_date is not None and later_date is not None and later_date < prior_date:
                later_before_prior_count += 1
            if prior_date is not None and later_date is not None and 0 <= (later_date - prior_date).days < MIN_DOSE_INTERVAL_DAYS:
                interval_under_min_count += 1
                print(f"WARNING: interval < {MIN_DOSE_INTERVAL_DAYS} days between {prior_aliases[0]} and {later_aliases[0]} for children_code {code_value}: {(later_date - prior_date).days} days")
            if later_date is not None and prior_source == "NOT RECEIVED YET":
                later_while_prior_not_received_count += 1

    duplicate_code_count = sum(1 for value in set(code_values) if code_values.count(value) > 1)

    print("Child verification summary:")
    print(f"  children_code missing: {missing_code_count}")
    print(f"  duplicate children_code values: {duplicate_code_count}")
    print(f"  DOB later than first_visit_date: {dob_after_first_visit_count}")
    print(f"  later dose earlier than prior dose: {later_before_prior_count}")
    print(f"  later dose with prior dose not received yet: {later_while_prior_not_received_count}")
    print(f"  interval between doses < {MIN_DOSE_INTERVAL_DAYS} days (OPV/Penta/MMR): {interval_under_min_count}")

    return {
        "children_code_missing": missing_code_count,
        "children_code_duplicate_values": duplicate_code_count,
        "dob_after_first_visit": dob_after_first_visit_count,
        "later_before_prior": later_before_prior_count,
        "later_while_prior_not_received": later_while_prior_not_received_count,
        "interval_under_min_days": interval_under_min_count,
    }


def _verify_cleaned_pregnancy_rows(row_values_list: list[tuple], header_index_map: dict[str, int], code_col_idx: int | None) -> dict[str, int]:
    missing_code_count = 0
    duplicate_code_count = 0
    td_later_before_prior_count = 0
    td2_while_td1_not_received_count = 0

    code_values = []
    for row_values in row_values_list:
        code_value = row_values[code_col_idx - 1] if code_col_idx is not None and code_col_idx - 1 < len(row_values) else None
        if _is_blank(code_value):
            missing_code_count += 1
        else:
            code_values.append(str(code_value).strip())

        def _date_value(aliases: list[str]) -> date | None:
            idx = _find_header_index(header_index_map, aliases)
            return _to_date(row_values[idx]) if idx is not None and idx < len(row_values) else None

        def _text_value(aliases: list[str]) -> str:
            idx = _find_header_index(header_index_map, aliases)
            value = row_values[idx] if idx is not None and idx < len(row_values) else None
            return str(value).strip().upper()

        td1_date = _date_value(["Td first time dose", "Td1 dose", "Td1", "Td ALOD", "TD1"])
        td2_date = _date_value(["Td second time dose", "Td2 dose", "Td2", "Td Two Doses", "TD2"])
        td1_source = _text_value(["Td1 Source", "TD1 Source", "Td first time dose source", "Td ALOD Source"])

        if td1_date is not None and td2_date is not None and td2_date < td1_date:
            td_later_before_prior_count += 1
        if td2_date is not None and td1_source == "NOT RECEIVED YET":
            td2_while_td1_not_received_count += 1

    duplicate_code_count = sum(1 for value in set(code_values) if code_values.count(value) > 1)

    print("PW verification summary:")
    print(f"  pw_code missing: {missing_code_count}")
    print(f"  duplicate pw_code values: {duplicate_code_count}")
    print(f"  Td2 earlier than Td1: {td_later_before_prior_count}")
    print(f"  Td2 received while Td1 not received yet: {td2_while_td1_not_received_count}")

    return {
        "pw_code_missing": missing_code_count,
        "pw_code_duplicate_values": duplicate_code_count,
        "td2_earlier_than_td1": td_later_before_prior_count,
        "td2_while_td1_not_received": td2_while_td1_not_received_count,
    }


def _format_issue_status(issues: list[str]) -> str:
    unique_issues: list[str] = []
    for issue in issues:
        if issue not in unique_issues:
            unique_issues.append(issue)
    return "OK" if not unique_issues else "; ".join(unique_issues)


def _build_child_row_status(row_values: tuple, header_index_map: dict[str, int], code_col_idx: int, code_counts: dict[str, int]) -> str:
    issues = []
    code_value = row_values[code_col_idx - 1] if code_col_idx - 1 < len(row_values) else None
    code_text = str(code_value).strip() if not _is_blank(code_value) else ""
    if _is_blank(code_value):
        issues.append("children_code missing")
    elif code_counts.get(code_text, 0) > 1:
        issues.append("children_code duplicate")

    dob_idx = _find_header_index(header_index_map, ["DOB", "Date of Birth", "DateOfBirth", "Birth Date", "Birthdate"])
    first_visit_idx = _find_header_index(header_index_map, [FIRST_VISIT_DATE_HEADER])
    dob_date = _to_date(row_values[dob_idx]) if dob_idx is not None and dob_idx < len(row_values) else None
    first_visit_date = _to_date(row_values[first_visit_idx]) if first_visit_idx is not None and first_visit_idx < len(row_values) else None
    if dob_date is not None and first_visit_date is not None and dob_date > first_visit_date:
        issues.append("DOB later than first_visit_date")

    ordered_checks = [
        (["OPV first time dose"], ["OPV second time dose"], ["OPV1 Source", "OPV1 source", "OPV1 src"]),
        (["OPV second time dose"], ["OPV third time dose"], ["OPV2 Source", "OPV2 source", "OPV2 src"]),
        (["Penta first time dose"], ["Penta second time dose"], ["Penta1 Source", "Penta1 source", "Penta1 src"]),
        (["Penta second time dose"], ["Penta third time dose"], ["Penta2 Source", "Penta2 source", "Penta2 src"]),
        (["MMR first timeadose", "MMR first time dose"], ["MMR second time dose"], ["MMR1 Source", "MMR1 source", "MMR1 src"]),
    ]
    for prior_aliases, later_aliases, prior_source_aliases in ordered_checks:
        prior_idx = _find_header_index(header_index_map, prior_aliases)
        later_idx = _find_header_index(header_index_map, later_aliases)
        prior_source_idx = _find_header_index(header_index_map, prior_source_aliases)
        prior_date = _to_date(row_values[prior_idx]) if prior_idx is not None and prior_idx < len(row_values) else None
        later_date = _to_date(row_values[later_idx]) if later_idx is not None and later_idx < len(row_values) else None
        prior_source = str(row_values[prior_source_idx]).strip().upper() if prior_source_idx is not None and prior_source_idx < len(row_values) and not _is_blank(row_values[prior_source_idx]) else ""
        if prior_date is not None and later_date is not None and later_date < prior_date:
            issues.append("later dose earlier than prior dose")
        if prior_date is not None and later_date is not None and 0 <= (later_date - prior_date).days < MIN_DOSE_INTERVAL_DAYS:
            prior_label = prior_source_aliases[0].split()[0]
            later_label = prior_label[:-1] + str(int(prior_label[-1]) + 1)
            issues.append(f"dose interval < {MIN_DOSE_INTERVAL_DAYS} days ({prior_label} to {later_label})")
        if later_date is not None and prior_source == "NOT RECEIVED YET":
            issues.append("later dose with prior not received")

    return _format_issue_status(issues)


def _build_pw_row_status(row_values: tuple, header_index_map: dict[str, int], code_col_idx: int | None) -> str:
    issues = []
    code_value = row_values[code_col_idx - 1] if code_col_idx is not None and code_col_idx - 1 < len(row_values) else None
    if _is_blank(code_value):
        issues.append("pw_code missing")

    td1_idx = _find_header_index(header_index_map, ["Td1", "Td ALOD", "Td first time dose", "Td1 dose"])
    td2_idx = _find_header_index(header_index_map, ["Td2", "Td Two Doses", "Td second time dose", "Td2 dose"])
    td1_source_idx = _find_header_index(header_index_map, ["Td1 Source", "TD1 Source", "Td first time dose source", "Td ALOD Source"])
    td1_date = _to_date(row_values[td1_idx]) if td1_idx is not None and td1_idx < len(row_values) else None
    td2_date = _to_date(row_values[td2_idx]) if td2_idx is not None and td2_idx < len(row_values) else None
    td1_source = str(row_values[td1_source_idx]).strip().upper() if td1_source_idx is not None and td1_source_idx < len(row_values) and not _is_blank(row_values[td1_source_idx]) else ""

    if td1_date is not None and td2_date is not None and td2_date < td1_date:
        issues.append("Td2 earlier than Td1")
    if td2_date is not None and td1_source == "NOT RECEIVED YET":
        issues.append("Td2 received while Td1 not received")

    return _format_issue_status(issues)


def _write_rows_sheet(out_wb, sheet_name: str, columns: list[str], rows: list[dict]) -> None:
    if sheet_name in out_wb.sheetnames:
        del out_wb[sheet_name]
    ws = out_wb.create_sheet(title=sheet_name)

    ws.append(columns)
    for row in rows:
        ws.append([row.get(col_name) for col_name in columns])


def _build_indicator_target_columns(year: int, indicator_name: str) -> dict[str, float | None]:
    year_targets = REFERENCE_QUARTER_TARGETS_BY_YEAR.get(year, {})
    if indicator_name not in year_targets:
        return {
            "Q1 Target": None,
            "Q2 Target": None,
            "Q3 Target": None,
            "Q4 Target": None,
        }

    q1_ref, q2_ref, q3_ref, q4_ref = year_targets[indicator_name]
    return {
        "Q1 Target": q1_ref * TARGET_FACTOR,
        "Q2 Target": q2_ref * TARGET_FACTOR,
        "Q3 Target": q3_ref * TARGET_FACTOR,
        "Q4 Target": q4_ref * TARGET_FACTOR,
    }


def _build_annual_target_column(year: int, indicator_name: str) -> float | None:
    year_targets = REFERENCE_QUARTER_TARGETS_BY_YEAR.get(year, {})
    if indicator_name not in year_targets:
        return None
    return sum(year_targets[indicator_name]) * TARGET_FACTOR


def _add_default_template_sheets(out_wb) -> list[str]:
    years = [2024, 2025, 2026]
    created: list[str] = []

    indicator_names = [
        "Penta3 under 1-yr-old",
        "Penta3 under 5-yr-old",
        "MMR1 under 1-yr-old",
        "MMR1 under 5-yr-old",
        "MMR2 under 5-yr-old",
        "Penta1 under 5-yr-old",
        "At least one dose under 5-yr-old",
        "Full dose under 5-yr-old",
        "Td ALOD",
        "Td Two Doses",
    ]
    quarter_cols = [
        "Q1 Target",
        "Q1 U1 Male",
        "Q1 U1 Female",
        "Q1 1-5 Male ",
        "Q1 1-5 Female",
        "Q2 Target",
        "Q2 U1 Male",
        "Q2 U1 Female",
        "Q2 1-5 Male ",
        "Q2 1-5 Female",
        "Q3 Target",
        "Q3 U1 Male",
        "Q3 U1 Female",
        "Q3 1-5 Male ",
        "Q3 1-5 Female",
        "Q4 Target",
        "Q4 U1 Male",
        "Q4 U1 Female",
        "Q4 1-5 Male ",
        "Q4 1-5 Female",
    ]
    indicator_columns = ["Period", "Organization", "Project Name", "indicator", *quarter_cols]
    indicator_rows = []
    for year in years:
        for name in indicator_names:
            row = {
                "Period": year,
                "Organization": "CHDN",
                "Project Name": "REACH-KK",
                "indicator": name,
            }
            row.update(_build_indicator_target_columns(year, name))
            for col in quarter_cols:
                row.setdefault(col, None)
            indicator_rows.append(row)
    _write_rows_sheet(out_wb, "indicators", indicator_columns, indicator_rows)
    created.append("indicators")

    td_alod_columns = ["Period", "Organization ", "Project Name", "Indicators ", "Annual Target", "Annual Achievement"]
    td_alod_rows = [
        {
            "Period": year,
            "Organization ": "CHDN",
            "Project Name": "REACH-KK",
            "Indicators ": "Td ALOD",
            "Annual Target": _build_annual_target_column(year, "Td ALOD"),
            "Annual Achievement": None,
        }
        for year in years
    ]
    _write_rows_sheet(out_wb, "Td_ALOD", td_alod_columns, td_alod_rows)
    created.append("Td_ALOD")

    alod_cummu_columns = [
        "Period",
        "Organization ",
        "Project Name",
        "Indicator",
        "Annual Target",
        "Annual U1 Male",
        "Annaul U1 Female",
        "Annual 1-5 Male ",
        "Annual 1-5 Female",
    ]
    alod_cummu_rows = [
        {
            "Period": year,
            "Organization ": "CHDN",
            "Project Name": "REACH-KK",
            "Indicator": "ALOD cumulative",
            "Annual Target": _build_annual_target_column(year, "At least one dose under 5-yr-old"),
            "Annual U1 Male": None,
            "Annaul U1 Female": None,
            "Annual 1-5 Male ": None,
            "Annual 1-5 Female": None,
        }
        for year in years
    ]
    _write_rows_sheet(out_wb, "ALOD_cummu", alod_cummu_columns, alod_cummu_rows)
    created.append("ALOD_cummu")

    idp_columns = [
        "Period ",
        "Organization ",
        "Project Name",
        "indicator",
        "Q1 IDP Male ",
        "Q1 IDP Female",
        "Q1 non-IDP Male ",
        "Q1 non-IDP Female",
        "Q2 IDP Male ",
        "Q2 IDP Female",
        "Q2 non-IDP Male ",
        "Q2 non-IDP Female",
        "Q3 IDP Male ",
        "Q3 IDP Female",
        "Q3 non-IDP Male ",
        "Q3 non-IDP Female",
        "Q4 IDP Male ",
        "Q4 IDP Female",
        "Q4 non-IDP Male ",
        "Q4 non-IDP Female",
    ]
    idp_rows = [
        {
            "Period ": year,
            "Organization ": "CHDN",
            "Project Name": "REACH-KK",
            "indicator": "Penta1 under 5-yr-old",
        }
        for year in years
    ]
    _write_rows_sheet(out_wb, "IDP", idp_columns, idp_rows)
    created.append("IDP")

    td2_columns = [
        "Period",
        "Organization ",
        "Project Name",
        "Indicators ",
        "Q1 Target",
        "Q1 Achievement",
        "Q2 Target",
        "Q2 Achievement",
        "Q3 Target",
        "Q3 Achievement",
        "Q4 Target",
        "Q4 Achievement",
    ]
    td2_rows = [
        {
            "Period": year,
            "Organization ": "CHDN",
            "Project Name": "REACH-KK",
            "Indicators ": "Td Two Doses",
            "Q1 Target": _build_indicator_target_columns(year, "Td Two Doses")["Q1 Target"],
            "Q1 Achievement": None,
            "Q2 Target": _build_indicator_target_columns(year, "Td Two Doses")["Q2 Target"],
            "Q2 Achievement": None,
            "Q3 Target": _build_indicator_target_columns(year, "Td Two Doses")["Q3 Target"],
            "Q3 Achievement": None,
            "Q4 Target": _build_indicator_target_columns(year, "Td Two Doses")["Q4 Target"],
            "Q4 Achievement": None,
        }
        for year in years
    ]
    _write_rows_sheet(out_wb, "Td2_indicator", td2_columns, td2_rows)
    created.append("Td2_indicator")

    return created


def _get_kept_rows(values_sheet, data_start_row: int, max_col: int) -> tuple[list[int], list[tuple], int]:
    """Return (row_numbers, row_data, removed_count) for non-blank rows.

    Extracting both row numbers and data in one pass avoids a second cell-by-cell
    scan when writing the pregnancy output sheet.
    """
    kept_rows: list[int] = []
    kept_data: list[tuple] = []
    removed_count = 0
    for row_num, row in enumerate(
        values_sheet.iter_rows(min_row=data_start_row, max_row=values_sheet.max_row, min_col=1, max_col=max_col, values_only=True),
        start=data_start_row,
    ):
        if any(not _is_blank(v) for v in row):
            kept_rows.append(row_num)
            kept_data.append(row)
        else:
            removed_count += 1
    return kept_rows, kept_data, removed_count


def build_clean_sheet(
    input_file: Path | str,
    output_file: Path | str,
    source_sheet_name: str = "EPI-Child",
    target_sheet_name: str = "EPI-Child",
    pregnancy_source_sheet_name: str = "EPI-Pregnancy",
    pregnancy_target_sheet_name: str = "EPI-Pregnancy",
) -> int:
    input_path = Path(input_file).resolve()
    output_path = Path(output_file).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    try:
        prefilled_input, stats = _apply_age_source_autofill(input_path, source_sheet_name)
    except Exception as exc:
        print(f"ERROR: Failed to prefill age/source values: {exc}")
        return 2

    print("Autofill summary (EPI-Child):")
    print(f"  Age values filled: {stats['age_filled']}")
    print(f"  First visit date filled: {stats['first_visit_date_filled']}")
    print(f"  Age at first visit filled: {stats['age_at_first_visit_filled']}")
    print(f"  Source values filled: {stats['source_filled']}")
    print(f"  Age columns added: {stats['age_columns_added']}")
    print(f"  First visit date columns added: {stats['first_visit_date_columns_added']}")
    print(f"  Source columns added: {stats['source_columns_added']}")

    wb = load_workbook(prefilled_input, data_only=True, read_only=False)
    if source_sheet_name not in wb.sheetnames:
        raise ValueError(f"Source sheet not found: {source_sheet_name}")

    child_ws = wb[source_sheet_name]
    max_col = child_ws.max_column
    # Read the header row once using iter_rows instead of individual cell accesses.
    source_headers = list(next(child_ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=max_col, values_only=True)))
    code_col_idx = _find_children_code_column(source_headers)
    if code_col_idx is None:
        raise ValueError("Could not find children_code(T/C Code) column in the source sheet")

    child_rows_to_keep: list[tuple] = []
    child_code_values: list[object] = []
    child_removed_count = 0

    # Use iter_rows(values_only=True) for bulk sequential reads — much faster than
    # individual child_ws.cell(row, col) accesses in a nested loop.
    for row_values in child_ws.iter_rows(min_row=2, max_row=child_ws.max_row, min_col=1, max_col=max_col, values_only=True):
        if any(not _is_blank(value) for value in row_values):
            child_rows_to_keep.append(row_values)
            child_code_values.append(row_values[code_col_idx - 1])
        else:
            child_removed_count += 1

    if child_removed_count > 0:
        print(f"INFO: Removed {child_removed_count} formula-only rows from {source_sheet_name}")

    _warn_code_quality(child_code_values, "children_code")

    out_wb = Workbook()
    child_out_ws = out_wb.active
    child_out_ws.title = target_sheet_name
    child_export_columns = list(source_headers) + list(COMPLETE_COLUMNS) + ["verification status"]
    child_out_ws.append(child_export_columns)
    _warn_missing_source_columns(child_export_columns)

    header_index_map = {header_name: idx for idx, header_name in enumerate(source_headers)}
    dob_idx = _find_header_index(header_index_map, ["DOB", "Date of Birth", "DateOfBirth", "Birth Date", "Birthdate"])
    red_fill = PatternFill(fill_type="solid", fgColor="FFC7CE")

    child_code_counts: dict[str, int] = {}
    for row_values in child_rows_to_keep:
        code_value = row_values[code_col_idx - 1] if code_col_idx - 1 < len(row_values) else None
        if not _is_blank(code_value):
            code_text = str(code_value).strip()
            child_code_counts[code_text] = child_code_counts.get(code_text, 0) + 1

    # Pre-compute all column indices once so the per-row loop does O(1) lookups
    # instead of O(headers) string-normalization scans for every row × quarter.
    completion_precomp = _precompute_completion_indices(header_index_map)

    for row_values in child_rows_to_keep:
        completed_values = [_completion_value(row_values, col_name, completion_precomp) for col_name in COMPLETE_COLUMNS]
        verification_status = _build_child_row_status(row_values, header_index_map, code_col_idx, child_code_counts)
        # Default absent source values to blank string so the row length stays consistent
        safe_row = [v if v is not None else "" for v in row_values]
        child_out_ws.append(safe_row + completed_values + [verification_status])
        if verification_status == "Issue":
            code_cell = child_out_ws.cell(row=child_out_ws.max_row, column=code_col_idx)
            code_cell.fill = red_fill

    _verify_cleaned_child_rows(child_rows_to_keep, header_index_map, code_col_idx)

    # Load pregnancy sheet from the original input file to preserve Excel-cached formula values
    # that would be lost after re-saving through _apply_age_source_autofill.
    orig_wb = load_workbook(input_path, data_only=True, read_only=True)
    if pregnancy_source_sheet_name in orig_wb.sheetnames:
        preg_ws = orig_wb[pregnancy_source_sheet_name]
        preg_max_col = preg_ws.max_column
        preg_header_row = 2
        preg_headers = list(next(
            preg_ws.iter_rows(min_row=preg_header_row, max_row=preg_header_row, min_col=1, max_col=preg_max_col, values_only=True)
        ))

        # _get_kept_rows now returns both row numbers (for _warn_code_quality) and
        # row data (for writing output) in a single pass over the pregnancy sheet.
        preg_rows_to_keep, preg_data_to_keep, preg_removed_count = _get_kept_rows(preg_ws, data_start_row=3, max_col=preg_max_col)
        if preg_removed_count > 0:
            print(f"INFO: Removed {preg_removed_count} formula-only rows from {pregnancy_source_sheet_name}")

        preg_code_col_idx = _find_pregnance_code_column(preg_headers)
        if preg_code_col_idx is None:
            print("WARNING: Could not find Pregnance_code column in EPI-Pregnancy sheet.")
        else:
            _warn_code_quality(preg_ws, preg_code_col_idx, "pw_code", preg_rows_to_keep)

        preg_header_index_map = {header_name: idx for idx, header_name in enumerate(preg_headers)}
        _verify_cleaned_pregnancy_rows(preg_data_to_keep, preg_header_index_map, preg_code_col_idx)

        preg_out_ws = out_wb.create_sheet(title=pregnancy_target_sheet_name)
        preg_out_ws.append(preg_headers + ["verification status"])
        _warn_missing_source_columns(preg_headers)
        for row_values in preg_data_to_keep:
            # Default absent source values to blank string to ensure all cells are written
            safe_row = [v if v is not None else "" for v in row_values]
            status_value = _build_pw_row_status(row_values, preg_header_index_map, preg_code_col_idx)
            preg_out_ws.append(safe_row + [status_value])
    else:
        print(f"WARNING: Pregnancy source sheet not found: {pregnancy_source_sheet_name}")

    created_templates = _add_default_template_sheets(out_wb)

    try:
        out_wb.save(output_path)
    except Exception as exc:
        raise RuntimeError(f"Failed to save workbook: {exc}") from exc

    print(f"Done: created {output_path.name}")
    if created_templates:
        print("Created template sheets: " + ", ".join(created_templates))
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Create CHDN_EPI_clean.xlsx with completion columns")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    sys.exit(build_clean_sheet(args.input, args.output))
