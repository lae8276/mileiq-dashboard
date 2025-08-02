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
        return 'UB7'
    full = re.search(r'\b([A-Z]{1,2}[0-9R][0-9A-Z]?) ?[0-9][ABD-HJLNP-UW-Z]{2}\b', location, re.IGNORECASE)
    if full:
        return full.group(1).upper()
    partial = re.search(r'\b([A-Z]{1,2}[0-9R][0-9A-Z]?)\b', location, re.IGNORECASE)
    return partial.group(1).upper() if partial else ''

def read_excel(file) -> pd.DataFrame:
    ext = file.name.lower()
    if ext.endswith('.xls'):
        return pd.read_excel(file, skiprows=39, header=None, usecols=[1, 2, 4, 7], engine='xlrd')
    else:
        return pd.read_excel(file, skiprows=39, header=None, usecols=[1, 2, 4, 7], engine='openpyxl')

def remove_consecutive_duplicates(postcodes_str: str) -> str:
    parts = postcodes_str.split(',')
    filtered = [parts[0]] if parts else []
    for i in range(1, len(parts)):
        if parts[i] != parts[i-1]:
            filtered.append(parts[i])
    return ','.join(filtered)

def process_file(file) -> tuple[pd.DataFrame, float]:
    df = read_excel(file)
    df.columns = ['Start Time', 'Start Location', 'End Location', 'Miles']
    
    # Convert Miles to float safely
    df['Miles'] = pd.to_numeric(df['Miles'], errors='coerce').fillna(0.0)

    df['Date'] = pd.to_datetime(df['Start Time'], errors='coerce', dayfirst=True).dt.date
    df['Start Postcode'] = df['Start Location'].apply(extract_postcode)
    df['End Postcode'] = df['End Location'].apply(extract_postcode)
    df['Postcodes'] = df[['Start Postcode', 'End Postcode']].apply(
        lambda x: ','.join([p for p in x if p]), axis=1
    )

    grouped = df.groupby('Date').agg({
        'Miles': 'sum',
        'Postcodes': lambda x: remove_consecutive_duplicates(','.join(x))
    }).reset_index()

    # Format date + round miles
    grouped['Date'] = pd.to_datetime(grouped['Date']).dt.strftime('%d-%b-%Y')
    grouped['Miles'] = grouped['Miles'].round(1)

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
uploaded_file = st.file_uploader("Upload your MileIQ file (.xlsx or .xls)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
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
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
