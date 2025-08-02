# file: mileiq_dashboard.py

import pandas as pd
import re
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="MileIQ Summary", layout="wide")

def extract_postcode(location: str) -> str:
    if not isinstance(location, str):
        return ''
    loc = location.lower().strip()
    if loc == 'home':
        return 'UB3'
    if 'rico pudo' in loc.replace(' ', ''):
        return 'UB6'

    # Try full postcode first
    full = re.search(r'\b([A-Z]{1,2}[0-9R][0-9A-Z]?) ?[0-9][ABD-HJLNP-UW-Z]{2}\b', location, re.IGNORECASE)
    if full:
        return full.group(1).upper()

    # Fallback: just outward part
    partial = re.search(r'\b([A-Z]{1,2}[0-9R][0-9A-Z]?)\b', location, re.IGNORECASE)
    return partial.group(1).upper() if partial else ''

def read_excel(file) -> pd.DataFrame:
    ext = file.name.lower()
    if ext.endswith('.xls'):
        return pd.read_excel(file, skiprows=39, header=None, usecols=[1, 2, 4, 7], engine='xlrd')
    else:
        return pd.read_excel(file, skiprows=39, header=None, usecols=[1, 2, 4, 7], engine='openpyxl')

def process_file(file) -> tuple[pd.DataFrame, float]:
    df = read_excel(file)
    df.columns = ['Start Time', 'Start Location', 'End Location', 'Miles']
    df['Date'] = pd.to_datetime(df['Start Time'], errors='coerce', dayfirst=True).dt.date
    df['Start Postcode'] = df['Start Location'].apply(extract_postcode)
    df['End Postcode'] = df['End Location'].apply(extract_postcode)
    df['Postcodes'] = df[['Start Postcode', 'End Postcode']].apply(lambda x: ','.join([p for p in x if p]), axis=1)
    grouped = df.groupby('Date').agg({
        'Miles': 'sum',
        'Postcodes': lambda x: ','.join(x)
    }).reset_index()
    total_miles = grouped['Miles'].sum()
    return grouped, total_miles

def convert_df_to_excel(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Summary')
    output.seek(0)
    return output

# Streamlit UI
st.title("\U0001F4CA MileIQ Mileage Summary")
uploaded_file = st.file_uploader("Upload your MileIQ file (.xlsx or .xls)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        summary_df, total_miles = process_file(uploaded_file)

        st.subheader("\U0001F4DD Daily Summary")
        for _, row in summary_df.iterrows():
            st.markdown(f"""
            <div ondblclick="navigator.clipboard.writeText(this.innerText)" style="padding:8px;border:1px solid #ddd;border-radius:5px;margin-bottom:4px;font-family:monospace">
            <b>Date:</b> {row['Date']}<br>
            <b>Miles:</b> {row['Miles']}<br>
            <b>Postcodes:</b> {row['Postcodes']}
            </div>
            """, unsafe_allow_html=True)

        excel_data = convert_df_to_excel(summary_df)
        st.download_button(
            label="\U0001F4BE Download Summary as Excel",
            data=excel_data,
            file_name="mileiq_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"\u274C Error processing file: {e}")
