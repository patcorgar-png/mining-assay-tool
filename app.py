"""
Mining Assay Tool — Streamlit UI
Run with:  streamlit run app.py
"""

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

# ── page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Mining Assay Tool",
    page_icon="⛏️",
    layout="wide",
)

# ── import pipeline (must be in the same folder) ───────────────────────────
try:
    from pipeline import (
        load_raw, clean_dataframe, find_best_intervals,
        write_excel, compute_aueq_series,
    )
except ImportError:
    st.error("❌ `pipeline.py` not found. Make sure it's in the same folder as `app.py`.")
    st.stop()


# ── helpers ────────────────────────────────────────────────────────────────

def _df_to_excel_bytes(cleaned_df, intervals, prices, cutoff) -> bytes:
    buf = io.BytesIO()
    write_excel(cleaned_df, intervals, prices, cutoff, buf)
    return buf.getvalue()


def _colour_no_intersection(val):
    return "color: #aaaaaa; font-style: italic;" if val == "No significant intersections" else ""


def _highlight_aueq(row, cutoff):
    col = "AuEq (g/t)"
    if col not in row.index:
        return [""] * len(row)
    try:
        v = float(row[col])
        if v >= cutoff:
            return ["background-color: #ffe0e0"] * len(row)
    except (ValueError, TypeError):
        pass
    return [""] * len(row)


# ── sidebar ────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("⛏️ Mining Assay Tool")
    st.caption("Upload a lab Excel file, configure prices, and get instant results.")

    st.divider()
    st.subheader("Metal Prices")
    st.caption("Precious metals in USD/troy oz · Base metals in USD/lb")

    col1, col2 = st.columns(2)
    with col1:
        au_price = st.number_input("Au ($/oz)", value=1950.0,  min_value=0.0, step=10.0)
        pb_price = st.number_input("Pb ($/lb)", value=0.95,    min_value=0.0, step=0.01, format="%.3f")
        cu_price = st.number_input("Cu ($/lb)", value=3.80,    min_value=0.0, step=0.01, format="%.3f")
    with col2:
        ag_price = st.number_input("Ag ($/oz)", value=24.0,    min_value=0.0, step=0.5)
        zn_price = st.number_input("Zn ($/lb)", value=1.20,    min_value=0.0, step=0.01, format="%.3f")
        ni_price = st.number_input("Ni ($/lb)", value=7.00,    min_value=0.0, step=0.05, format="%.2f")

    st.divider()
    st.subheader("Interval Parameters")
    cutoff      = st.number_input("AuEq Cutoff (g/t)",       value=0.5,  min_value=0.0, step=0.1, format="%.2f")
    drop_tol    = st.slider(      "Drop Tolerance",           min_value=0.0, max_value=1.0, value=0.5, step=0.05,
                                  help="Fraction below cutoff allowed inside an interval before it's split")
    min_length  = st.number_input("Min Interval Length (ft)", value=10.0, min_value=0.0, step=5.0)
    recovery    = st.slider(      "Base Metal Recovery",      min_value=0.0, max_value=1.0, value=0.75, step=0.05,
                                  help="Assumed metallurgical recovery for Pb, Zn, Cu etc.")

    st.divider()
    st.subheader("Advanced (optional)")
    header_row_override   = st.number_input("Header row override (0-based, -1 = auto)",
                                            value=-1, min_value=-1, step=1)
    data_start_override   = st.number_input("Data start row override (0-based, -1 = auto)",
                                            value=-1, min_value=-1, step=1)
    sheet_input = st.text_input("Sheet name or index", value="0",
                                help="Leave as 0 for the first sheet, or type the sheet name")

prices = {
    "Au": au_price, "Ag": ag_price,
    "Pb": pb_price, "Zn": zn_price,
    "Cu": cu_price, "Ni": ni_price,
}
# Drop elements with zero price (user doesn't want them)
prices = {k: v for k, v in prices.items() if v > 0}


# ── main area ──────────────────────────────────────────────────────────────

st.header("Mining Assay Data Processor")

uploaded = st.file_uploader(
    "Upload your lab assay Excel file (.xlsx or .xls)",
    type=["xlsx", "xls"],
    help="Handles messy files with letterhead rows, mixed units, detection-limit strings, etc.",
)

if uploaded is None:
    st.info("👆 Upload a file to get started.")
    st.stop()


# ── write upload to temp file (needed for openpyxl) ───────────────────────

suffix = Path(uploaded.name).suffix
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
    tmp.write(uploaded.read())
    tmp_path = tmp.name

# Show available sheets so user can pick the right one
try:
    import openpyxl as _opx
    _wb = _opx.load_workbook(tmp_path, read_only=True, data_only=True)
    sheet_names = _wb.sheetnames
    _wb.close()
    if len(sheet_names) > 1:
        st.info(f"📋 **Sheets detected:** {', '.join(sheet_names)}  — use the Sheet field in Advanced to pick one.")
except Exception:
    sheet_names = []

# ── run pipeline on upload ─────────────────────────────────────────────────

hr   = None if header_row_override < 0 else int(header_row_override)
dsr  = None if data_start_override < 0 else int(data_start_override)
sht  = int(sheet_input) if sheet_input.strip().isdigit() else sheet_input.strip()

with st.spinner("Reading and cleaning data…"):
    try:
        raw_df     = load_raw(tmp_path, header_row=hr, data_start_row=dsr, sheet=sht)
        cleaned_df, _ = clean_dataframe(raw_df)
    except ValueError as e:
        st.error(f"**Data format error:** {e}")
        st.caption("Try adjusting the Header row or Data start row in Advanced settings, or check the Sheet name.")
        st.stop()
    except Exception as e:
        st.error(f"**Failed to load file:** {e}")
        with st.expander("Show full error"):
            import traceback
            st.code(traceback.format_exc())
        st.stop()

elem_found = [e for e in prices if e in cleaned_df.columns]
if not elem_found:
    st.warning(
        "No priced elements found in the file. "
        "Check that column headers contain element symbols (Au, Ag, Pb, Zn…) "
        "and that the relevant prices above are non-zero."
    )
    st.stop()

st.success(
    f"Loaded **{len(cleaned_df):,} samples** from **{cleaned_df['Hole_ID'].nunique()} holes**. "
    f"Elements found: {', '.join(elem_found)}"
)


# ── run interval finder whenever prices/cutoff change ─────────────────────

with st.spinner("Finding best intervals…"):
    try:
        intervals = find_best_intervals(
            cleaned_df, prices, cutoff,
            drop_tolerance=drop_tol,
            min_length_ft=min_length,
            recovery=recovery,
        )
    except Exception as e:
        st.error(f"Interval finder error: {e}")
        st.stop()


# ── results summary ────────────────────────────────────────────────────────

n_holes = len(intervals)
n_qual  = int((~intervals["no_intersection"]).sum()) if n_holes else 0

colA, colB, colC = st.columns(3)
colA.metric("Total Holes",           n_holes)
colB.metric("Qualifying Intervals",  n_qual)
colC.metric("Below Cutoff",          n_holes - n_qual)

st.divider()


# ── results table ──────────────────────────────────────────────────────────

st.subheader("Best Intervals")

display_cols = ["Hole_ID", "From", "To", "Length_ft", "Length_m", "AuEq_avg"]
display_cols += [e for e in elem_found]
display_cols += ["Value_per_tonne"]

rename_map = {
    "Hole_ID": "Hole ID", "From": "From (ft)", "To": "To (ft)",
    "Length_ft": "Length (ft)", "Length_m": "Length (m)",
    "AuEq_avg": "AuEq (g/t)", "Value_per_tonne": "Value ($/t)",
    **{e: (f"{e} (g/t)" if e in {"Au","Ag","Pt","Pd"} else f"{e} (%)") for e in elem_found},
}

disp = intervals[[c for c in display_cols if c in intervals.columns]].copy()
disp = disp.rename(columns=rename_map)

# Fill "No significant intersections" text into From column for display
ni_mask = intervals["no_intersection"].fillna(False).astype(bool)
for col in ["From (ft)", "To (ft)", "Length (ft)", "Length (m)", "AuEq (g/t)",
            "Value ($/t)"] + list(rename_map.values()):
    if col in disp.columns:
        disp.loc[ni_mask, col] = ""

disp.loc[ni_mask, "From (ft)"] = "No significant intersections"

# Format numbers
num_cols = [c for c in disp.columns if disp[c].dtype in (float,) or
            (disp[c].apply(lambda x: isinstance(x, float)).any())]
for c in disp.columns:
    if c not in ("Hole ID", "From (ft)"):
        disp[c] = disp[c].apply(lambda x: f"{x:.3f}" if isinstance(x, float) else x)

styled = (
    disp.style
    .map(lambda v: "color:#aaaaaa;font-style:italic;" if v == "No significant intersections" else "", subset=["From (ft)"])
    .apply(_highlight_aueq, cutoff=cutoff, axis=1)
    .set_properties(**{"text-align": "center"})
    .set_table_styles([{"selector": "th", "props": [("text-align", "center"), ("background-color", "#2F4F4F"), ("color", "white")]}])
)
st.dataframe(styled, use_container_width=True, hide_index=True)


# ── raw data preview ───────────────────────────────────────────────────────

with st.expander("Raw Data Preview (cleaned)", expanded=False):
    preview_cols = ["Hole_ID", "From", "To"] + elem_found
    preview = cleaned_df[[c for c in preview_cols if c in cleaned_df.columns]].copy()

    # Add computed AuEq column for preview
    preview["AuEq (g/t)"] = compute_aueq_series(cleaned_df, prices, recovery=recovery).round(4)

    # Highlight rows above cutoff
    def _highlight_raw(row):
        try:
            if float(row.get("AuEq (g/t)", 0)) >= cutoff:
                return ["background-color: #ffe0e0"] * len(row)
        except (ValueError, TypeError):
            pass
        return [""] * len(row)

    st.caption(f"{len(preview):,} rows · pink rows ≥ {cutoff} g/t AuEq")
    st.dataframe(
        preview.style.apply(_highlight_raw, axis=1),
        use_container_width=True,
        hide_index=True,
        height=300,
    )


# ── download ───────────────────────────────────────────────────────────────

st.divider()
st.subheader("Download Results")

with st.spinner("Building Excel workbook…"):
    try:
        xlsx_bytes = _df_to_excel_bytes(cleaned_df, intervals, prices, cutoff)
    except Exception as e:
        st.error(f"Excel generation error: {e}")
        st.stop()

out_name = Path(uploaded.name).stem + "_processed.xlsx"
st.download_button(
    label="⬇️  Download Excel workbook",
    data=xlsx_bytes,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)
st.caption(
    "The workbook includes a **Results** sheet (intervals + parameters block) "
    "and a **Raw Data** sheet with live AuEq formulas. "
    "Changing prices in the Parameters section recalculates every formula instantly."
)
