"""
Mining Assay Data Processing Pipeline
======================================
Loads messy lab Excel assay files, cleans and standardises data,
computes AuEq grades, finds best mineralised intervals, and writes
a professional multi-sheet Excel workbook.

Entry point
-----------
    run_pipeline(filepath, prices, cutoff, output_path)

prices : dict  e.g. {"Au": 1950, "Ag": 24, "Pb": 0.95, "Zn": 1.20}
            Precious metals (Au, Ag) : USD / troy oz
            Base metals (Pb, Zn, Cu): USD / lb
cutoff : float  minimum AuEq g/t for a qualifying interval
"""

from __future__ import annotations

import re
import warnings
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side,
)
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# ---------------------------------------------------------------------------
# 1. Element / column-name catalogue
# ---------------------------------------------------------------------------

_ELEMENT_ALIASES: dict[str, str] = {
    "gold":      "Au", "au":   "Au",
    "silver":    "Ag", "ag":   "Ag",
    "lead":      "Pb", "pb":   "Pb",
    "zinc":      "Zn", "zn":   "Zn",
    "copper":    "Cu", "cu":   "Cu",
    "iron":      "Fe", "fe":   "Fe",
    "arsenic":   "As", "as":   "As",
    "antimony":  "Sb", "sb":   "Sb",
    "bismuth":   "Bi", "bi":   "Bi",
    "cadmium":   "Cd", "cd":   "Cd",
    "cobalt":    "Co", "co":   "Co",
    "chromium":  "Cr", "cr":   "Cr",
    "manganese": "Mn", "mn":   "Mn",
    "molybdenum":"Mo", "mo":   "Mo",
    "nickel":    "Ni", "ni":   "Ni",
    "phosphorus":"P",  "p":    "P",
    "sulfur":    "S",  "s":    "S",
    "tin":       "Sn", "sn":   "Sn",
    "tungsten":  "W",  "w":    "W",
    "vanadium":  "V",  "v":    "V",
    "barium":    "Ba", "ba":   "Ba",
    "calcium":   "Ca", "ca":   "Ca",
    "potassium": "K",  "k":    "K",
    "magnesium": "Mg", "mg":   "Mg",
    "sodium":    "Na", "na":   "Na",
    "silicon":   "Si", "si":   "Si",
    "titanium":  "Ti", "ti":   "Ti",
}

# Priced in USD/troy oz; all others assumed USD/lb (base metals in %)
_GT_ELEMENTS = {"Au", "Ag", "Pt", "Pd"}

_UNIT_CONVERSIONS: dict[str, float] = {
    "oz/t":  31.1035,
    "opt":   31.1035,
    "ppm":   0.0001,    # ppm -> % for base metals
    "ppb":   0.0000001,
    "g/t":   1.0,
    "%":     1.0,
}


# ---------------------------------------------------------------------------
# 2. Header / data-row auto-detection
# ---------------------------------------------------------------------------

_HEADER_RE = re.compile(
    r"\b(au|ag|pb|zn|cu|fe|sample|hole|from|to|depth|interval|assay)\b",
    re.IGNORECASE,
)


def _score_row_as_header(row: pd.Series) -> int:
    score = 0
    for val in row.dropna():
        txt = str(val).strip()
        if _HEADER_RE.search(txt):
            score += 2
        if len(txt) <= 6 and txt.isalpha():
            score += 1
    return score


def detect_header_row(raw: pd.DataFrame) -> tuple[int, int]:
    """
    Scan the first 40 rows to find the header row and first data row.
    Returns (header_row_index, data_start_row_index) -- zero-based.
    """
    best_score, best_idx = -1, 0
    for i in range(min(40, len(raw))):
        score = _score_row_as_header(raw.iloc[i])
        if score > best_score:
            best_score, best_idx = score, i
    data_start = best_idx + 1
    while data_start < len(raw) and raw.iloc[data_start].isna().all():
        data_start += 1
    return best_idx, data_start


# ---------------------------------------------------------------------------
# 3. Load Excel file
# ---------------------------------------------------------------------------

def load_raw(
    filepath: str | Path,
    header_row: Optional[int] = None,
    data_start_row: Optional[int] = None,
    sheet: int | str = 0,
) -> pd.DataFrame:
    """
    Load an assay Excel file into a clean DataFrame.

    Parameters
    ----------
    filepath       : path to .xlsx / .xls file
    header_row     : zero-based row index of the header; None => auto-detect
    data_start_row : zero-based row index of first data row; None => auto-detect
    sheet          : sheet index or name (default 0)
    """
    filepath = Path(filepath)
    raw = pd.read_excel(
        filepath, sheet_name=sheet,
        header=None, dtype=str, keep_default_na=False,
    )
    raw.replace("", np.nan, inplace=True)

    if header_row is None:
        header_row, data_start_auto = detect_header_row(raw)
        if data_start_row is None:
            data_start_row = data_start_auto
    elif data_start_row is None:
        data_start_row = header_row + 1

    col_names = [
        str(v).strip() if pd.notna(v) else f"_col{i}"
        for i, v in enumerate(raw.iloc[header_row])
    ]
    df = raw.iloc[data_start_row:].copy()
    df.columns = col_names
    df = df.loc[:, ~df.columns.str.startswith("_col")]
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ---------------------------------------------------------------------------
# 4. Detection-limit cleaning
# ---------------------------------------------------------------------------

_DL_LESS_THAN    = re.compile(r"^[<<]\s*([\d.]+)\s*$")
_DL_GREATER_THAN = re.compile(r"^[>>]\s*([\d.]+)\s*$")


def clean_detection_limits(value):
    """
    '<0.005'  -> 0.0
    '>10000'  -> 10000.0
    Anything else is coerced to float or returned as-is.
    """
    if not isinstance(value, str):
        return value
    s = value.strip()
    m = _DL_LESS_THAN.match(s)
    if m:
        return 0.0
    m = _DL_GREATER_THAN.match(s)
    if m:
        return float(m.group(1))
    try:
        return float(s)
    except ValueError:
        return s


# ---------------------------------------------------------------------------
# 5. Unit metadata extraction
# ---------------------------------------------------------------------------

_UNIT_RE = re.compile(
    r"\((" + "|".join(re.escape(u) for u in _UNIT_CONVERSIONS) + r")\)",
    re.IGNORECASE,
)


def extract_units_from_columns(columns: list[str]) -> dict[str, str]:
    """Parse units embedded in column headers like 'Au (g/t)' -> {'Au (g/t)': 'g/t'}"""
    return {
        col: m.group(1).lower()
        for col in columns
        if (m := _UNIT_RE.search(col))
    }


# ---------------------------------------------------------------------------
# 6. Column-name standardisation
# ---------------------------------------------------------------------------

def _parse_col_to_element(col: str) -> Optional[str]:
    base = _UNIT_RE.sub("", col).strip().lower()
    base = re.sub(r"[^a-z]", "", base)
    return _ELEMENT_ALIASES.get(base)


def standardise_columns(df: pd.DataFrame):
    """
    Rename messy assay columns to canonical element symbols.

    Returns (df, col_map, raw_units)
    """
    raw_units = extract_units_from_columns(df.columns.tolist())
    col_map: dict[str, str] = {}
    seen: dict[str, int] = {}

    for col in df.columns:
        elem = _parse_col_to_element(col)
        if elem:
            seen[elem] = seen.get(elem, 0) + 1
            col_map[col] = elem if seen[elem] == 1 else f"{elem}_{seen[elem]}"
        else:
            col_map[col] = col

    df = df.rename(columns=col_map)
    return df, col_map, raw_units


# ---------------------------------------------------------------------------
# 7. Unit conversion
# ---------------------------------------------------------------------------

def apply_unit_conversions(
    df: pd.DataFrame,
    raw_units: dict[str, str],
    col_map: dict[str, str],
) -> pd.DataFrame:
    """
    Convert element columns to canonical units:
      - Au, Ag  -> g/t
      - others  -> %
    """
    df = df.copy()
    for old_col, unit in raw_units.items():
        new_col = col_map.get(old_col, old_col)
        if new_col not in df.columns:
            continue
        elem = _parse_col_to_element(old_col)
        if not elem:
            continue

        u = unit.lower()
        if elem in _GT_ELEMENTS:
            if u in ("oz/t", "opt"):
                factor = 31.1035
            elif u == "ppm":
                factor = 0.001  # ppm -> g/t (1 ppm = 0.001 g/t)
            else:
                factor = 1.0   # already g/t
        else:
            if u == "ppm":
                factor = 0.0001  # ppm -> %
            elif u == "ppb":
                factor = 0.0000001
            else:
                factor = 1.0   # already %

        if factor != 1.0:
            df[new_col] = pd.to_numeric(df[new_col], errors="coerce") * factor

    return df


# ---------------------------------------------------------------------------
# 8. Parse hole ID + depth strings
# ---------------------------------------------------------------------------

_HOLEID_DEPTH_RE  = re.compile(
    r"^([A-Za-z]{1,6}[-_]?\d+)\s+(\d+\.?\d*)\s*[-–_]\s*(\d+\.?\d*)$"
)
_HOLEID_DEPTH_RE2 = re.compile(
    r"^([A-Za-z]{1,6}[-_]?\d+)[-_](\d+\.?\d*)[-_](\d+\.?\d*)$"
)


def _try_parse_combined(val: str):
    for pat in (_HOLEID_DEPTH_RE, _HOLEID_DEPTH_RE2):
        m = pat.match(str(val).strip())
        if m:
            return m.group(1), float(m.group(2)), float(m.group(3))
    return None


def _find_col(df: pd.DataFrame, patterns: list[str]) -> Optional[str]:
    for pat in patterns:
        for col in df.columns:
            if re.search(pat, col, re.IGNORECASE):
                return col
    return None


def parse_hole_depth_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure df has Hole_ID, From, To columns."""
    df = df.copy()

    hole_col = _find_col(df, [r"hole", r"bhid", r"dh.?id", r"drill", r"holeid"])
    from_col = _find_col(df, [r"^from$", r"from_", r"_from", r"depth.?from", r"start"])
    to_col   = _find_col(df, [r"^to$",   r"to_",   r"_to",   r"depth.?to",   r"end"])

    if hole_col and from_col and to_col:
        rename = {}
        if hole_col != "Hole_ID": rename[hole_col] = "Hole_ID"
        if from_col != "From":    rename[from_col] = "From"
        if to_col   != "To":      rename[to_col]   = "To"
        df.rename(columns=rename, inplace=True)
        df["From"] = pd.to_numeric(df["From"], errors="coerce")
        df["To"]   = pd.to_numeric(df["To"],   errors="coerce")
        return df

    # Try combined column
    for col in df.columns:
        sample = df[col].dropna().head(10)
        parsed = [_try_parse_combined(v) for v in sample]
        parsed = [p for p in parsed if p]
        if len(parsed) >= max(3, len(sample) // 2):
            extracted = df[col].apply(
                lambda v: _try_parse_combined(str(v)) if pd.notna(v) else None
            )
            df["Hole_ID"] = extracted.apply(lambda x: x[0] if x else np.nan)
            df["From"]    = extracted.apply(lambda x: float(x[1]) if x else np.nan)
            df["To"]      = extracted.apply(lambda x: float(x[2]) if x else np.nan)
            df.drop(columns=[col], inplace=True)
            return df

    if hole_col:
        df.rename(columns={hole_col: "Hole_ID"}, inplace=True)
    if "Hole_ID" not in df.columns:
        df.insert(0, "Hole_ID", "UNKNOWN")

    num_cols = [
        c for c in df.columns
        if c != "Hole_ID"
        and pd.to_numeric(df[c], errors="coerce").notna().sum() > len(df) * 0.5
    ]
    if len(num_cols) >= 2 and "From" not in df.columns:
        df.rename(columns={num_cols[0]: "From", num_cols[1]: "To"}, inplace=True)

    df["From"] = pd.to_numeric(df.get("From", pd.Series(dtype=float)), errors="coerce")
    df["To"]   = pd.to_numeric(df.get("To",   pd.Series(dtype=float)), errors="coerce")
    return df


# ---------------------------------------------------------------------------
# 9. Full data cleaning pass
# ---------------------------------------------------------------------------

def clean_dataframe(df: pd.DataFrame):
    """
    Run the full cleaning pipeline on a raw DataFrame.
    Returns (cleaned_df, col_map).
    """
    df = df.map(clean_detection_limits)
    df, col_map, raw_units = standardise_columns(df)
    df = apply_unit_conversions(df, raw_units, col_map)
    df = parse_hole_depth_columns(df)

    for col in [c for c in df.columns if c in _ELEMENT_ALIASES.values()]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df.dropna(subset=["From", "To"], how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df, col_map


# ---------------------------------------------------------------------------
# 10. AuEq calculation (Python-side, for interval finding)
# ---------------------------------------------------------------------------
#
# Pricing conventions
# -------------------
# Precious metals (Au, Ag) : USD / troy oz
#   AuEq g/t contribution  = grade_g/t * price_$/oz / au_price_$/oz
#
# Base metals (Pb, Zn, Cu) : USD / lb
#   1% grade = 10,000 g/tonne = 10,000 / 453.592 lb/tonne = 22.046 lb/t
#   $/tonne  = grade_% * 22.046 * price_$/lb * recovery
#   AuEq g/t = $/tonne / (au_price_$/oz / 31.1035 g/oz)
#            = grade_% * 22.046 * 31.1035 * price_$/lb * recovery / au_price
#            = grade_% * 685.94 * price_$/lb * recovery / au_price
#
_LB_SCALE = 10_000 * 31.1035 / 453.592  # ~685.94 -- lb/t scaling constant


def compute_aueq_series(
    df: pd.DataFrame,
    prices: dict[str, float],
    recovery: float = 0.75,
) -> pd.Series:
    """Compute AuEq (g/t) for every sample row using prices dict."""
    au_price = prices.get("Au", 1) or 1
    aueq = pd.Series(0.0, index=df.index)

    for elem, price in prices.items():
        if elem not in df.columns:
            continue
        grade = pd.to_numeric(df[elem], errors="coerce").fillna(0)
        if elem in _GT_ELEMENTS:
            aueq = aueq + grade * price / au_price
        else:
            aueq = aueq + grade * _LB_SCALE * price * recovery / au_price

    return aueq


# ---------------------------------------------------------------------------
# 11. Sliding-window best-interval finder
# ---------------------------------------------------------------------------

def find_best_intervals(
    df: pd.DataFrame,
    prices: dict[str, float],
    cutoff: float,
    drop_tolerance: float = 0.5,
    min_length_ft: float = 10.0,
    recovery: float = 0.75,
) -> pd.DataFrame:
    """
    Find the best mineralised intervals per hole using a
    peak-comparison weighted-average sliding window.

    Parameters
    ----------
    df            : cleaned DataFrame with Hole_ID, From, To, element columns
    prices        : metal price dict used for AuEq
    cutoff        : minimum AuEq g/t to qualify
    drop_tolerance: fraction of cutoff that can fall below before breaking
    min_length_ft : minimum interval length in feet
    recovery      : metallurgical recovery for base metals

    Returns
    -------
    DataFrame with qualifying intervals + "No significant intersections" rows.
    """
    df = df.copy()
    df["_AuEq"] = compute_aueq_series(df, prices, recovery=recovery)

    results: list[dict] = []
    elem_cols = [e for e in prices if e in df.columns]
    lower_cut = cutoff * (1.0 - drop_tolerance)

    # $/tonne value scaling
    lb_val = 10_000 / 453.592  # ~22.046 lb per tonne per 1%

    for hole_id, hole_df in df.groupby("Hole_ID", sort=False):
        hole_df = hole_df.sort_values("From").reset_index(drop=True)
        n = len(hole_df)
        best: Optional[dict] = None
        best_aueq = -np.inf

        i = 0
        while i < n:
            if hole_df.at[i, "_AuEq"] < lower_cut:
                i += 1
                continue

            j = i
            while j < n and hole_df.at[j, "_AuEq"] >= lower_cut:
                j += 1

            window = hole_df.iloc[i:j]
            length_ft = float(window["To"].max() - window["From"].min())
            if length_ft < min_length_ft:
                i = j
                continue

            weights = (window["To"] - window["From"]).clip(lower=0)
            if weights.sum() == 0:
                i = j
                continue

            wav_aueq = float(np.average(window["_AuEq"], weights=weights))
            if wav_aueq >= cutoff and wav_aueq > best_aueq:
                best_aueq = wav_aueq
                best = {"window": window, "AuEq_avg": wav_aueq}
            i = j

        if best is None:
            results.append({
                "Hole_ID": hole_id, "From": None, "To": None,
                "Length_ft": None, "Length_m": None,
                "AuEq_avg": None, "Value_per_tonne": None,
                "no_intersection": True,
                **{e: None for e in elem_cols},
            })
        else:
            w = best["window"]
            length_ft = float(w["To"].max() - w["From"].min())
            weights   = (w["To"] - w["From"]).clip(lower=0)

            row: dict = {
                "Hole_ID":        hole_id,
                "From":           float(w["From"].min()),
                "To":             float(w["To"].max()),
                "Length_ft":      round(length_ft, 2),
                "Length_m":       round(length_ft * 0.3048, 2),
                "AuEq_avg":       round(best["AuEq_avg"], 4),
                "no_intersection": False,
            }

            value_pt = 0.0
            for elem in elem_cols:
                g = float(np.average(
                    pd.to_numeric(w[elem], errors="coerce").fillna(0),
                    weights=weights
                )) if elem in w.columns else 0.0
                row[elem] = round(g, 4)
                price = prices.get(elem, 0)
                if elem in _GT_ELEMENTS:
                    value_pt += g * price / 31.1035
                else:
                    value_pt += g * lb_val * price * recovery

            row["Value_per_tonne"] = round(value_pt, 2)
            results.append(row)

    return pd.DataFrame(results) if results else pd.DataFrame()


# ---------------------------------------------------------------------------
# 12. Excel formatting helpers
# ---------------------------------------------------------------------------

_BLUE_FONT   = Font(color="0070C0", bold=True)
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_GREY_FONT   = Font(color="808080", italic=True)
_DARK_FILL   = PatternFill("solid", fgColor="2F4F4F")
_LIGHT_FILL  = PatternFill("solid", fgColor="D9E1F2")
_PINK_FILL   = PatternFill("solid", fgColor="FFD7D7")
_PARAM_FILL  = PatternFill("solid", fgColor="EBF3FB")
_THIN        = Side(style="thin", color="CCCCCC")
_THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_FMT_GRADE = "0.0000"
_FMT_MONEY = '#,##0.00'


def _sheet_ref(sheet_name: str, cell_addr: str) -> str:
    """Build a cross-sheet cell reference safe for names with spaces."""
    return f"'{sheet_name}'!{cell_addr}"


def _build_aueq_formula(
    param_cells: dict[str, str],
    elem_col_map: dict[str, str],
    row: int,
    results_sheet: str = "Results",
) -> str:
    """
    Build a live Excel AuEq formula for a data row.

    param_cells  : {"Au": "$B$2", "Ag": "$B$3", ..., "recovery": "$B$8"}
    elem_col_map : {"Au": "D", "Ag": "E", "Pb": "F", "Zn": "G"}
    row          : 1-based Excel row number
    """
    au_ref  = _sheet_ref(results_sheet, param_cells["Au"]) if "Au" in param_cells else "1950"
    rec_ref = _sheet_ref(results_sheet, param_cells.get("recovery", "$B$99"))
    LB = "685.94"

    parts = []
    for elem, col_letter in elem_col_map.items():
        if elem not in param_cells:
            continue
        price_ref = _sheet_ref(results_sheet, param_cells[elem])
        g = f"IF(ISNUMBER({col_letter}{row}),{col_letter}{row},0)"
        if elem in _GT_ELEMENTS:
            parts.append(f"({g}*{price_ref}/{au_ref})")
        else:
            parts.append(f"({g}*{LB}*{price_ref}*{rec_ref}/{au_ref})")

    return ("=" + "+".join(parts)) if parts else "=0"


def _build_value_formula(
    param_cells: dict[str, str],
    elem_col_map: dict[str, str],
    row: int,
    results_sheet: str = "Results",
) -> str:
    """Build a live $/tonne Excel formula for a data row."""
    rec_ref = _sheet_ref(results_sheet, param_cells.get("recovery", "$B$99"))
    LB_VAL  = "22.046"  # 10000 / 453.592

    parts = []
    for elem, col_letter in elem_col_map.items():
        if elem not in param_cells:
            continue
        price_ref = _sheet_ref(results_sheet, param_cells[elem])
        g = f"IF(ISNUMBER({col_letter}{row}),{col_letter}{row},0)"
        if elem in _GT_ELEMENTS:
            parts.append(f"({g}*{price_ref}/31.1035)")
        else:
            parts.append(f"({g}*{LB_VAL}*{price_ref}*{rec_ref})")

    return ("=" + "+".join(parts)) if parts else "=0"


# ---------------------------------------------------------------------------
# 13. Parameters block writer
# ---------------------------------------------------------------------------

def _write_parameters_block(
    ws,
    prices: dict[str, float],
    cutoff: float,
    start_row: int = 1,
) -> dict[str, str]:
    """
    Write the editable Parameters block.
    Returns {param_name: "$B$n"} addresses for formula use.
    """
    # Title banner
    ws.merge_cells(f"A{start_row}:D{start_row}")
    c = ws[f"A{start_row}"]
    c.value = "PARAMETERS  -  edit the blue cells to recalculate the workbook"
    c.font  = Font(bold=True, size=12, color="FFFFFF")
    c.fill  = _DARK_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[start_row].height = 22

    # Sub-header
    r = start_row + 1
    for ci, h in enumerate(["Parameter", "Value", "Unit", "Notes"], 1):
        cell = ws.cell(row=r, column=ci, value=h)
        cell.font   = Font(bold=True)
        cell.fill   = _LIGHT_FILL
        cell.border = _THIN_BORDER

    r += 1
    param_cells: dict[str, str] = {}

    _PRICE_META = {
        "Au": ("Au Price",  "USD/troy oz", "Gold spot price"),
        "Ag": ("Ag Price",  "USD/troy oz", "Silver spot price"),
        "Pb": ("Pb Price",  "USD/lb",      "Lead spot price"),
        "Zn": ("Zn Price",  "USD/lb",      "Zinc spot price"),
        "Cu": ("Cu Price",  "USD/lb",      "Copper spot price"),
        "Mo": ("Mo Price",  "USD/lb",      "Molybdenum spot price"),
        "Ni": ("Ni Price",  "USD/lb",      "Nickel spot price"),
        "Co": ("Co Price",  "USD/lb",      "Cobalt spot price"),
    }

    for elem, price in prices.items():
        label, unit, note = _PRICE_META.get(
            elem, (f"{elem} Price", "USD/unit", f"{elem} spot price")
        )
        ws.cell(row=r, column=1).value = label
        c = ws.cell(row=r, column=2)
        c.value = price
        c.font  = _BLUE_FONT
        c.fill  = _PARAM_FILL
        ws.cell(row=r, column=3).value = unit
        ws.cell(row=r, column=4).value = note
        for ci in range(1, 5):
            ws.cell(row=r, column=ci).border = _THIN_BORDER
        param_cells[elem] = f"$B${r}"
        r += 1

    # Cutoff row
    ws.cell(row=r, column=1).value = "AuEq Cutoff"
    c = ws.cell(row=r, column=2)
    c.value = cutoff
    c.font  = _BLUE_FONT
    c.fill  = _PARAM_FILL
    ws.cell(row=r, column=3).value = "g/t AuEq"
    ws.cell(row=r, column=4).value = "Minimum grade for a qualifying interval"
    for ci in range(1, 5):
        ws.cell(row=r, column=ci).border = _THIN_BORDER
    param_cells["cutoff"] = f"$B${r}"
    r += 1

    # Recovery row
    ws.cell(row=r, column=1).value = "Base Metal Recovery"
    c = ws.cell(row=r, column=2)
    c.value = 0.75
    c.font  = _BLUE_FONT
    c.fill  = _PARAM_FILL
    c.number_format = "0%"
    ws.cell(row=r, column=3).value = "fraction"
    ws.cell(row=r, column=4).value = "Assumed metallurgical recovery for Pb, Zn, Cu etc."
    for ci in range(1, 5):
        ws.cell(row=r, column=ci).border = _THIN_BORDER
    param_cells["recovery"] = f"$B${r}"

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 44

    return param_cells


# ---------------------------------------------------------------------------
# 14. Results sheet
# ---------------------------------------------------------------------------

def _write_results_sheet(
    wb: openpyxl.Workbook,
    intervals: pd.DataFrame,
    prices: dict[str, float],
    cutoff: float,
    param_cells: dict,
    sheet_name: str = "Results",
) -> None:
    ws = wb.active
    ws.title = sheet_name

    # Parameters block (fills param_cells in-place)
    param_cells.update(
        _write_parameters_block(ws, prices, cutoff, start_row=1)
    )
    param_end = 1 + 1 + len(prices) + 2  # banner + sub-header + prices + cutoff + recovery
    table_start = param_end + 2

    elem_cols = [e for e in prices if e in intervals.columns]

    # Table header
    headers = (
        ["Hole ID", "From (ft)", "To (ft)", "Length (ft)", "Length (m)", "AuEq (g/t)"]
        + [f"{e} (g/t)" if e in _GT_ELEMENTS else f"{e} (%)" for e in elem_cols]
        + ["Value ($/t)"]
    )

    r = table_start
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font   = _HEADER_FONT
        c.fill   = _DARK_FILL
        c.border = _THIN_BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[r].height = 30

    aueq_col    = 6
    elem_start  = 7
    value_col   = elem_start + len(elem_cols)

    # elem -> column letter map for formula building
    elem_col_map = {
        elem: get_column_letter(elem_start + i)
        for i, elem in enumerate(elem_cols)
    }

    r += 1
    for _, row_data in intervals.iterrows():
        is_ni = bool(row_data.get("no_intersection", False))
        ws.cell(row=r, column=1, value=str(row_data["Hole_ID"]))

        if is_ni:
            last = len(headers)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last)
            mc = ws.cell(row=r, column=2, value="No significant intersections")
            mc.font = _GREY_FONT
            mc.alignment = Alignment(horizontal="center")
            for ci in range(1, last + 1):
                ws.cell(row=r, column=ci).border = _THIN_BORDER
        else:
            ws.cell(row=r, column=2, value=row_data.get("From"))
            ws.cell(row=r, column=3, value=row_data.get("To"))
            ws.cell(row=r, column=4, value=row_data.get("Length_ft"))
            ws.cell(row=r, column=5, value=row_data.get("Length_m"))

            # Live AuEq formula
            ws.cell(row=r, column=aueq_col,
                    value=_build_aueq_formula(param_cells, elem_col_map, r, sheet_name)
                    ).number_format = _FMT_GRADE

            # Element grades (static values)
            for i, elem in enumerate(elem_cols):
                c = ws.cell(row=r, column=elem_start + i, value=row_data.get(elem))
                c.number_format = _FMT_GRADE

            # Live value formula
            ws.cell(row=r, column=value_col,
                    value=_build_value_formula(param_cells, elem_col_map, r, sheet_name)
                    ).number_format = _FMT_MONEY

            for ci in range(1, len(headers) + 1):
                ws.cell(row=r, column=ci).border = _THIN_BORDER

        r += 1

    # Column widths
    widths = [14, 10, 10, 12, 12, 12] + [12] * len(elem_cols) + [14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = ws[f"A{table_start + 1}"]


# ---------------------------------------------------------------------------
# 15. Raw Data sheet
# ---------------------------------------------------------------------------

def _write_raw_data_sheet(
    wb: openpyxl.Workbook,
    df: pd.DataFrame,
    prices: dict[str, float],
    param_cells: dict[str, str],
    results_sheet: str = "Results",
    sheet_name: str = "Raw Data",
) -> None:
    ws = wb.create_sheet(title=sheet_name)

    elem_cols = [e for e in prices if e in df.columns]
    header_labels = (
        ["Hole ID", "From (ft)", "To (ft)"]
        + [f"{e} (g/t)" if e in _GT_ELEMENTS else f"{e} (%)" for e in elem_cols]
        + ["AuEq (g/t)"]
    )

    for ci, h in enumerate(header_labels, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font   = _HEADER_FONT
        c.fill   = _DARK_FILL
        c.border = _THIN_BORDER
        c.alignment = Alignment(horizontal="center")

    elem_start = 4
    elem_col_map = {
        elem: get_column_letter(elem_start + i)
        for i, elem in enumerate(elem_cols)
    }
    aueq_col_letter = get_column_letter(elem_start + len(elem_cols))

    for ri, (_, row_data) in enumerate(df.iterrows(), start=2):
        ws.cell(row=ri, column=1, value=row_data.get("Hole_ID"))
        ws.cell(row=ri, column=2, value=row_data.get("From"))
        ws.cell(row=ri, column=3, value=row_data.get("To"))
        for ci in range(1, 4):
            ws.cell(row=ri, column=ci).border = _THIN_BORDER

        for i, elem in enumerate(elem_cols):
            c = ws.cell(row=ri, column=elem_start + i,
                        value=row_data.get(elem))
            c.number_format = _FMT_GRADE
            c.border = _THIN_BORDER

        # Live AuEq formula
        aueq_c = ws.cell(
            row=ri,
            column=elem_start + len(elem_cols),
            value=_build_aueq_formula(param_cells, elem_col_map, ri, results_sheet),
        )
        aueq_c.number_format = _FMT_GRADE
        aueq_c.border = _THIN_BORDER

    # Conditional formatting: highlight rows where AuEq >= cutoff
    last_row   = 1 + len(df)
    first_col  = "A"
    cutoff_ref = _sheet_ref(results_sheet, param_cells["cutoff"])
    cf_range   = f"{first_col}2:{aueq_col_letter}{last_row}"
    ws.conditional_formatting.add(
        cf_range,
        FormulaRule(
            formula=[f"${aueq_col_letter}2>={cutoff_ref}"],
            fill=_PINK_FILL,
        ),
    )

    # Column widths
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 10
    for letter in elem_col_map.values():
        ws.column_dimensions[letter].width = 12
    ws.column_dimensions[aueq_col_letter].width = 12

    ws.freeze_panes = ws["A2"]


# ---------------------------------------------------------------------------
# 16. Main Excel writer
# ---------------------------------------------------------------------------

def write_excel(
    cleaned_df: pd.DataFrame,
    intervals: pd.DataFrame,
    prices: dict[str, float],
    cutoff: float,
    output_path: str | Path,
) -> Path:
    """
    Write the full Excel workbook with Results and Raw Data sheets.
    Returns the output path.
    """
    output_path = Path(output_path)
    wb = openpyxl.Workbook()
    param_cells: dict[str, str] = {}

    _write_results_sheet(wb, intervals, prices, cutoff, param_cells)
    _write_raw_data_sheet(wb, cleaned_df, prices, param_cells)

    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# 17. Main entry point
# ---------------------------------------------------------------------------

def run_pipeline(
    filepath: str | Path,
    prices: dict[str, float],
    cutoff: float,
    output_path: str | Path,
    *,
    header_row: Optional[int] = None,
    data_start_row: Optional[int] = None,
    sheet: int | str = 0,
    drop_tolerance: float = 0.5,
    min_length_ft: float = 10.0,
    recovery: float = 0.75,
) -> Path:
    """
    Full pipeline: raw Excel -> cleaned data -> intervals -> output Excel.

    Parameters
    ----------
    filepath       : path to input assay Excel file
    prices         : {"Au": 1950, "Ag": 24, "Pb": 0.95, "Zn": 1.20}
                     Precious metals in USD/troy oz, base metals in USD/lb
    cutoff         : AuEq g/t cut-off for qualifying intervals
    output_path    : where to save the output .xlsx
    header_row     : override auto-detection (0-based)
    data_start_row : override auto-detection (0-based)
    sheet          : sheet index or name in the input file
    drop_tolerance : fraction below cutoff allowed inside an interval
    min_length_ft  : minimum qualifying interval length in feet
    recovery       : metallurgical recovery for base metals (default 0.75)

    Returns
    -------
    Path to the written output file.
    """
    print(f"[1/4] Loading {filepath} ...")
    raw_df = load_raw(filepath, header_row=header_row,
                      data_start_row=data_start_row, sheet=sheet)

    print(f"[2/4] Cleaning data  ({len(raw_df)} rows) ...")
    cleaned_df, _ = clean_dataframe(raw_df)
    elem_found = [e for e in prices if e in cleaned_df.columns]
    print(f"      Elements found: {elem_found}")

    print(f"[3/4] Finding intervals  (cutoff={cutoff} g/t AuEq) ...")
    intervals = find_best_intervals(
        cleaned_df, prices, cutoff,
        drop_tolerance=drop_tolerance,
        min_length_ft=min_length_ft,
        recovery=recovery,
    )
    if len(intervals):
        n_qual = int((~intervals["no_intersection"]).sum())
        print(f"      {n_qual}/{len(intervals)} holes with qualifying intervals")
    else:
        print("      No holes processed.")

    print(f"[4/4] Writing Excel -> {output_path} ...")
    out = write_excel(cleaned_df, intervals, prices, cutoff, output_path)
    print("      Done.")
    return out


# ---------------------------------------------------------------------------
# 18. CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser(
        description="Mining Assay Pipeline -- process lab Excel files"
    )
    parser.add_argument("input",  help="Input assay Excel file (.xlsx)")
    parser.add_argument("output", help="Output Excel file (.xlsx)")
    parser.add_argument(
        "--prices",
        default='{"Au": 1950, "Ag": 24, "Pb": 0.95, "Zn": 1.20}',
        help="JSON dict of metal prices e.g. '{\"Au\":1950,\"Ag\":24}'",
    )
    parser.add_argument("--cutoff",     type=float, default=0.5,
                        help="AuEq g/t cutoff (default 0.5)")
    parser.add_argument("--header-row", type=int,   default=None,
                        help="Override header row (0-based)")
    parser.add_argument("--data-start", type=int,   default=None,
                        help="Override data start row (0-based)")
    parser.add_argument("--sheet",      default="0",
                        help="Sheet index or name (default 0)")
    parser.add_argument("--drop-tol",   type=float, default=0.5,
                        help="Drop tolerance fraction (default 0.5)")
    parser.add_argument("--min-length", type=float, default=10.0,
                        help="Minimum interval length in feet (default 10)")
    parser.add_argument("--recovery",   type=float, default=0.75,
                        help="Base metal metallurgical recovery (default 0.75)")

    args = parser.parse_args()
    prices_in = json.loads(args.prices)
    sheet_in  = int(args.sheet) if args.sheet.isdigit() else args.sheet

    run_pipeline(
        filepath       = args.input,
        prices         = prices_in,
        cutoff         = args.cutoff,
        output_path    = args.output,
        header_row     = args.header_row,
        data_start_row = args.data_start,
        sheet          = sheet_in,
        drop_tolerance = args.drop_tol,
        min_length_ft  = args.min_length,
        recovery       = args.recovery,
    )
