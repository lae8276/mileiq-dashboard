# file: app/mileiq_dashboard.py
from __future__ import annotations

import re
from io import BytesIO
from typing import Iterable, List, Tuple
import math

import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="MileIQ Summary, Overtime & PDF Merger", layout="wide")

POSTCODE_FULL_RE = re.compile(r"\b([A-Z]{1,2}[0-9R][0-9A-Z]?) ?[0-9][ABD-HJLNP-UW-Z]{2}\b", re.IGNORECASE)
POSTCODE_PARTIAL_RE = re.compile(r"\b([A-Z]{1,2}[0-9R][0-9A-Z]?)\b", re.IGNORECASE)

def extract_postcode(location: str) -> str:
    if not isinstance(location, str):
        return ""
    loc = location.strip().lower()
    if loc == "home":
        return "UB3"
    if "rico pudo" in loc:
        return "UB7"
    full = POSTCODE_FULL_RE.search(location)
    if full:
        return full.group(1).upper()
    partial = POSTCODE_PARTIAL_RE.search(location)
    return partial.group(1).upper() if partial else ""

def read_excel(file) -> pd.DataFrame:
    name = getattr(file, "name", "").lower()
    usecols = [1, 2, 4, 7]
    if name.endswith(".xls"):
        return pd.read_excel(file, skiprows=39, header=None, usecols=usecols, engine="xlrd")
    if name.endswith(".xlsx"):
        return pd.read_excel(file, skiprows=39, header=None, usecols=usecols, engine="openpyxl")
    raise ValueError("Unsupported file type. Please upload .xls or .xlsx.")

def remove_consecutive_duplicates(postcodes_str: str) -> str:
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
def process_file(file) -> Tuple[pd.DataFrame, float, pd.DataFrame]:
    df = read_excel(file)
    df.columns = ["Start Time", "Start Location", "End Location", "Miles"]
    df["Miles"] = pd.to_numeric(df["Miles"], errors="coerce").fillna(0.0)
    df["Start Time Parsed"] = pd.to_datetime(df["Start Time"], errors="coerce", dayfirst=True)
    df["Date"] = df["Start Time Parsed"].dt.date
    df["Start Postcode"] = df["Start Location"].apply(extract_postcode)
    df["End Postcode"] = df["End Location"].apply(extract_postcode)
    df["Postcodes"] = df[["Start Postcode", "End Postcode"]].apply(lambda x: ",".join([p for p in x if p]), axis=1)
    grouped = (
        df.groupby("Date")
        .agg(
            Miles=("Miles", "sum"),
            Postcodes=("Postcodes", lambda s: remove_consecutive_duplicates(",".join(s))),
        )
        .reset_index()
    )
    grouped["Date"] = pd.to_datetime(grouped["Date"]).dt.strftime("%d-%b-%Y")
    grouped["Miles"] = grouped["Miles"].round(1)
    total_miles = float(grouped["Miles"].sum())
    return grouped, total_miles, df

@st.cache_data(show_spinner=False)
def convert_df_to_excel(df: pd.DataFrame) -> BytesIO:
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
            reader.decrypt("")
        except Exception as exc:
            raise ValueError(f"Encrypted PDF not supported: {getattr(file, 'name', '')}") from exc

def merge_pdfs_by_filename(files: Iterable) -> Tuple[BytesIO, List[str]]:
    files = list(files)
    if not files:
        raise ValueError("No PDF files provided.")
    for f in files:
        _validate_pdf_readable(f)
    writer = PdfWriter()
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

# UI Tabs
summary_tab, overtime_tab, merge_tab = st.tabs(["📊 MileIQ Summary", "⏱️ Overtime Calculator", "📎 Merge PDFs"])

with summary_tab:
    st.title("📊 MileIQ Mileage Summary")
    uploaded_file = st.file_uploader("Upload your MileIQ file (.xlsx or .xls)", type=["xlsx", "xls"], key="summary_upload")
    if uploaded_file is not None:
        try:
            summary_df, total_miles, _ = process_file(uploaded_file)
            st.metric(label="🚗 Total Miles", value=f"{total_miles:.1f} mi")
            st.dataframe(summary_df, use_container_width=True)
            excel_data = convert_df_to_excel(summary_df)
            st.download_button("💾 Download Summary as Excel", data=excel_data, file_name="mileiq_summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")

with overtime_tab:
    st.title("⏱️ Overtime Calculator")
    overtime_file = st.file_uploader("Upload your MileIQ file (.xlsx or .xls)", type=["xlsx", "xls"], key="overtime_upload")
    if overtime_file is not None:
        try:
            _, _, raw_df = process_file(overtime_file)
            raw_df["End Time Parsed"] = pd.to_datetime(raw_df["Start Time"], errors="coerce", dayfirst=True)
            raw_df["DayOfWeek"] = raw_df["End Time Parsed"].dt.dayofweek

            def get_cutoff_time(day_of_week: int):
                return pd.to_datetime("16:30" if day_of_week >= 5 else "17:30", format="%H:%M").time()

            overtime_rows = []
            for date, group in raw_df.groupby(raw_df["End Time Parsed"].dt.date):
                latest_home = group[group["End Postcode"] == "UB3"].sort_values("End Time Parsed").tail(1)
                if latest_home.empty:
                    continue
                arrival_time = latest_home["End Time Parsed"].iloc[0]
                cutoff_time = get_cutoff_time(arrival_time.weekday())
                if arrival_time.time() > cutoff_time:
                    hours = (pd.Timestamp.combine(pd.Timestamp.today(), arrival_time.time()) - pd.Timestamp.combine(pd.Timestamp.today(), cutoff_time)).seconds / 3600
                    # Round up to nearest 0.5 hour
                    hours = math.ceil(hours * 2) / 2
                    overtime_rows.append({
                        "Date": date.strftime("%d-%b-%Y"),
                        "Day": arrival_time.strftime("%A"),
                        "Home Arrival": arrival_time.strftime("%H:%M"),
                        "Overtime Hours": hours
                    })

            overtime_df = pd.DataFrame(overtime_rows)
            total_overtime = overtime_df["Overtime Hours"].sum() if not overtime_df.empty else 0.0

            st.metric(label="⏱️ Total Overtime", value=f"{total_overtime:.2f} hrs")
            st.dataframe(overtime_df, use_container_width=True)
        except Exception as e:
            st.error(f"❌ Error calculating overtime: {e}")

with merge_tab:
    st.title("📎 Merge PDFs")
    pdf_files = st.file_uploader("Upload PDFs to merge (numeric filenames recommended)", type=["pdf"], accept_multiple_files=True)
    if pdf_files:
        try:
            st.write(f"📄 {len(pdf_files)} file(s) uploaded:")
            sorted_names = sorted((f.name for f in pdf_files), key=_sort_key_numeric_first)
            for name in sorted_names:
                st.markdown(f"- {name}")
            merged_pdf, _ = merge_pdfs_by_filename(pdf_files)
            st.download_button("📥 Download Merged PDF", data=merged_pdf, file_name="merged_output.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"❌ Error merging PDFs: {e}")
