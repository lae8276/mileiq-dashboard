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
    if 'rico pudo' in loc:
        return 'UB6'
    match = re.search(r'\b([A-Z]{1,2}[0-9R][0-9A-Z]?) ?([0-9][ABD-HJLNP-UW-Z]{2})\b', location, re.IGNORECASE)
    return (match.group(1) + match.group(2)).upper() if match else ''

def process_file(file) -> tuple[pd.DataFrame, float]:
    df = pd.read_excel(file, skiprows=39, header=None, usecols=[1, 2, 4, 7])
    df.columns = ['Start Time', 'Start Location', 'End Location', 'Miles']
    df['Date'] = pd.to_datetime(df['Start Time'], errors='coerce').dt.date
    df['Start Postcode'] = df['Start Location'].apply(extract_postcode)
    df['End Postcode'] = df['End Location'].apply(extract_postcode)
    df['Postcodes'] = df['Start Postcode'] + ',' + df['End Postcode']

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
st.title("📊 MileIQ Mileage Summary")
uploaded_file = st.file_uploader("Upload your MileIQ .xlsx file", type=['xlsx'])

if uploaded_file:
    summary_df, total_miles = process_file(uploaded_file)

    st.metric(label="🚗 Total Miles", value=f"{total_miles:.1f} mi")
    st.dataframe(summary_df, use_container_width=True)

    excel_data = convert_df_to_excel(summary_df)
    st.download_button(
        label="💾 Download Summary as Excel",
        data=excel_data,
        file_name="mileiq_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
