# file: app/mileiq_dashboard.py
from __future__ import annotations

import math
import re
from datetime import date, timedelta
from io import BytesIO
from typing import Iterable, List, Tuple

import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="MileIQ Summary, Overtime & PDF Merger", layout="wide")

# -------------------------- Regex & Utilities --------------------------
POSTCODE_FULL_RE = re.compile(r"\b([A-Z]{1,2}[0-9R][0-9A-Z]?) ?[0-9][ABD-HJLNP-UW-Z]{2}\b", re.IGNORECASE)
POSTCODE_PARTIAL_RE = re.compile(r"\b([A-Z]{1,2}[0-9R][0-9A-Z]?)\b", re.IGNORECASE)

def extract_postcode(location: str) -> str:
    if not isinstance(location, str):
        return ""
    loc = (location or "").strip().lower()
    if loc == "home":
        return "UB3"
    if "rico pudo" in loc:
        return "UB7"
    full = POSTCODE_FULL_RE.search(location or "")
    if full:
        return full.group(1).upper()
    partial = POSTCODE_PARTIAL_RE.search(location or "")
    return partial.group(1).upper() if partial else ""

def read_excel(file) -> pd.DataFrame:
    name = getattr(file, "name", "").lower()
    usecols = [1, 2, 4, 7]  # Start Time, Start Location, End Location, Miles
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
    df["Event Time"] = pd.to_datetime(df["Start Time"], errors="coerce", dayfirst=True)
    df["Date"] = df["Event Time"].dt.date
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
    m = re.search(r"(\d+)", name or "")
    return (int(m.group(1)) if m else 10**12, (name or "").lower())

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

# ------------------------------- UI Tabs -------------------------------
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

    # mandatory pattern: you provide S (worked Sunday) and T (worked Saturday)
    col_s, col_t = st.columns(2)
    with col_s:
        s_date = st.date_input("S: Sunday you worked (S)", value=None, format="DD/MM/YYYY")
    with col_t:
        t_date = st.date_input("T: Saturday you worked (T)", value=None, format="DD/MM/YYYY")

    def weekend_cutoff(d: date):
        # Sat/Sun -> 16:30, else 17:30
        return pd.to_datetime("16:30" if pd.Timestamp(d).weekday() in (5, 6) else "17:30").time()

    if overtime_file is not None and s_date and t_date:
        try:
            _, _, raw_df = process_file(overtime_file)  # has Event Time, Date, End Postcode
            # compute OffDays based on your core rule
            s = pd.to_datetime(s_date).date()
            t = pd.to_datetime(t_date).date()
            off_days = {
                s - timedelta(days=2),  # S-2 (Fri)
                s - timedelta(days=1),  # S-1 (Sat)
                t + timedelta(days=1),  # T+1 (Sun)
                t + timedelta(days=2),  # T+2 (Mon)
            }
            work_days_extra = {s, t}

            # pre-index trips by date
            by_date = dict(tuple(raw_df.groupby("Date")))

            overtime_rows = []
            for d in sorted(by_date.keys()):
                day_df = by_date[d]

                # any activity on an OffDay => full 7.5h
                if d in off_days and not day_df.empty:
                    overtime_rows.append({
                        "SortDate": pd.to_datetime(d),
                        "Date": pd.to_datetime(d).strftime("%d-%b-%Y"),
                        "Day": pd.to_datetime(d).strftime("%A"),
                        "Home Arrival": "",  # not needed for full-day OT
                        "Overtime Hours": 7.5,
                        "Flag": "🔴 o"
                    })
                    continue

                # otherwise: calculate based on arriving home after cutoff
                home_rows = day_df[day_df["End Postcode"].str.upper() == "UB3"]
                if home_rows.empty:
                    continue  # no return home recorded; no overtime

                latest_arrival = home_rows.sort_values("Event Time").iloc[-1]["Event Time"]
                cut = pd.Timestamp.combine(pd.Timestamp(d).date(), weekend_cutoff(d))
                diff_hours = (latest_arrival - cut).total_seconds() / 3600.0
                if diff_hours <= 0:
                    continue

                hours = math.ceil(diff_hours * 2) / 2.0
                hours = min(hours, 7.5)
                overtime_rows.append({
                    "SortDate": pd.to_datetime(d),
                    "Date": pd.to_datetime(d).strftime("%d-%b-%Y"),
                    "Day": pd.to_datetime(d).strftime("%A"),
                    "Home Arrival": pd.to_datetime(latest_arrival).strftime("%H:%M"),
                    "Overtime Hours": hours,
                    "Flag": "🔴 o" if hours == 7.5 and (d in work_days_extra or d not in off_days) else ""
                })

            # build table safely even if empty
            if overtime_rows:
                overtime_df = pd.DataFrame(overtime_rows).sort_values("SortDate").drop(columns=["SortDate"]).reset_index(drop=True)
            else:
                overtime_df = pd.DataFrame(columns=["Date", "Day", "Home Arrival", "Overtime Hours", "Flag"])

            total_overtime = float(overtime_df["Overtime Hours"].sum()) if not overtime_df.empty else 0.0

            st.metric(label="⏱️ Total Overtime", value=f"{total_overtime:.2f} hrs")
            st.dataframe(overtime_df, use_container_width=True)

        except Exception as e:
            st.error(f"❌ Error calculating overtime: {e}")
    else:
        st.info("Select both S (worked Sunday) and T (worked Saturday) to compute overtime.")

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
