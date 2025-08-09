# file: app/mileiq_dashboard.py
from __future__ import annotations

import re
from io import BytesIO
from typing import Iterable, List, Tuple

import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter


# --------------------------- Streamlit Page Config ---------------------------
st.set_page_config(page_title="MileIQ Summary & PDF Merger", layout="wide")


# --------------------------------- Helpers ---------------------------------
POSTCODE_FULL_RE = re.compile(
    r"\b([A-Z]{1,2}[0-9R][0-9A-Z]?) ?[0-9][ABD-HJLNP-UW-Z]{2}\b",
    re.IGNORECASE,
)
POSTCODE_PARTIAL_RE = re.compile(r"\b([A-Z]{1,2}[0-9R][0-9A-Z]?)\b", re.IGNORECASE)


def extract_postcode(location: str) -> str:
    """Return outward UK postcode (e.g., UB3) from a freeform location.

    Why: downstream grouping needs stable, short codes; we normalize common aliases.
    """
    if not isinstance(location, str):
        return ""
    loc = location.strip().lower()

    # Normalization of frequent places
    if loc == "home":
        return "UB3"
    if "rico pudo" in loc:  # business alias
        return "UB7"

    # Prefer full postcode when available
    full = POSTCODE_FULL_RE.search(location)
    if full:
        return full.group(1).upper()

    # Fallback to outward code
    partial = POSTCODE_PARTIAL_RE.search(location)
    return partial.group(1).upper() if partial else ""


def read_excel(file) -> pd.DataFrame:
    """Read MileIQ export, handling both .xls and .xlsx.

    Why: xlrd>=2 drops .xlsx; we choose engine explicitly to avoid surprises.
    """
    name = getattr(file, "name", "").lower()
    usecols = [1, 2, 4, 7]  # Start Time, Start Loc, End Loc, Miles (per provided layout)

    if name.endswith(".xls"):
        return pd.read_excel(file, skiprows=39, header=None, usecols=usecols, engine="xlrd")
    if name.endswith(".xlsx"):
        return pd.read_excel(file, skiprows=39, header=None, usecols=usecols, engine="openpyxl")
    raise ValueError("Unsupported file type. Please upload .xls or .xlsx.")


def remove_consecutive_duplicates(postcodes_str: str) -> str:
    """Remove *consecutive* duplicates after splitting by comma.

    Why: same outward code via round-trips shouldn't repeat in the sequence.
    """
    if not postcodes_str:
        return ""
    parts = [p.strip() for p in postcodes_str.split(",") if p.strip()]
    if not parts:
        return ""
    filtered = [parts[0]]
    for i in range(1, len(parts)):
        if parts[i] != parts[i - 1]:
            filtered.append(parts[i])
    return ",".join(filtered)


@st.cache_data(show_spinner=False)
def process_file(file) -> Tuple[pd.DataFrame, float]:
    """Transform raw rows into a per-day summary with postcodes chain and miles sum."""
    df = read_excel(file)
    df.columns = ["Start Time", "Start Location", "End Location", "Miles"]

    # Parse/clean
    df["Miles"] = pd.to_numeric(df["Miles"], errors="coerce").fillna(0.0)
    df["Date"] = pd.to_datetime(df["Start Time"], errors="coerce", dayfirst=True).dt.date

    # Extract outward codes
    df["Start Postcode"] = df["Start Location"].apply(extract_postcode)
    df["End Postcode"] = df["End Location"].apply(extract_postcode)

    # Build visit chain per row (omit blanks)
    df["Postcodes"] = df[["Start Postcode", "End Postcode"]].apply(
        lambda x: ",".join([p for p in x if p]), axis=1
    )

    # Aggregate per day
    grouped = (
        df.groupby("Date")
        .agg(
            Miles=("Miles", "sum"),
            Postcodes=("Postcodes", lambda s: remove_consecutive_duplicates(",".join(s))),
        )
        .reset_index()
    )

    # Format
    grouped["Date"] = pd.to_datetime(grouped["Date"]).dt.strftime("%d-%b-%Y")
    grouped["Miles"] = grouped["Miles"].round(1)
    total_miles = float(grouped["Miles"].sum())
    return grouped, total_miles


@st.cache_data(show_spinner=False)
def convert_df_to_excel(df: pd.DataFrame) -> BytesIO:
    """Return Excel bytes for download, cached by content."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Summary")
    output.seek(0)
    return output


def _sort_key_numeric_first(name: str) -> Tuple[int, str]:
    m = re.search(r"(\d+)", name)
    return (int(m.group(1)) if m else 10**12, name.lower())


def _validate_pdf_readable(file) -> None:
    reader = PdfReader(file)
    if reader.is_encrypted:
        try:
            reader.decrypt("")  # try empty password
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"Encrypted PDF not supported: {getattr(file, 'name', '')}") from exc


def merge_pdfs_by_filename(files: Iterable) -> Tuple[BytesIO, List[str]]:
    """Merge PDFs using numeric-aware filename ordering.

    Why: predictable order by run number; gracefully skip unreadable pages.
    """
    files = list(files)
    if not files:
        raise ValueError("No PDF files provided.")

    # Ensure all are readable first (fail-fast)
    for f in files:
        _validate_pdf_readable(f)

    writer = PdfWriter()

    # Sort filenames by first integer, then by name
    sorted_files = sorted(files, key=lambda f: _sort_key_numeric_first(getattr(f, "name", "")))

    for file in sorted_files:
        file.seek(0)
        reader = PdfReader(file)
        for page in reader.pages:
            writer.add_page(page)

    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output, [getattr(f, "name", "") for f in sorted_files]


# ----------------------------------- UI ------------------------------------
summary_tab, merge_tab = st.tabs(["📊 MileIQ Summary", "📎 Merge PDFs"])\

with summary_tab:
    st.title("📊 MileIQ Mileage Summary")
    uploaded_file = st.file_uploader("Upload your MileIQ file (.xlsx or .xls)", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            summary_df, total_miles = process_file(uploaded_file)
            st.metric(label="🚗 Total Miles", value=f"{total_miles:.1f} mi")
            st.dataframe(summary_df, use_container_width=True)

            excel_data = convert_df_to_excel(summary_df)
            st.download_button(
                label="💾 Download Summary as Excel",
                data=excel_data,
                file_name="mileiq_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:  # noqa: BLE001
            st.error(f"❌ Error processing file: {e}")

with merge_tab:
    st.title("📎 Merge PDFs")
    pdf_files = st.file_uploader(
        "Upload PDFs to merge (numeric filenames recommended)",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if pdf_files:
        try:
            st.write(f"📄 {len(pdf_files)} file(s) uploaded:")

            # Show sorted names that will be used
            sorted_names = sorted((f.name for f in pdf_files), key=_sort_key_numeric_first)
            for name in sorted_names:
                st.markdown(f"- {name}")

            merged_pdf, _ = merge_pdfs_by_filename(pdf_files)
            st.download_button(
                label="📥 Download Merged PDF",
                data=merged_pdf,
                file_name="merged_output.pdf",
                mime="application/pdf",
            )
        except Exception as e:  # noqa: BLE001
            st.error(f"❌ Error merging PDFs: {e}")


# --------------------------------- Footer ----------------------------------
st.caption(
    "Made with ❤️  | Tip: Use numeric prefixes in PDF filenames (e.g., 001_.pdf, 010_.pdf) for reliable order."
)
