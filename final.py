import streamlit as st
import pandas as pd
from io import BytesIO

# --- Mappings ---
territory_region_map = {
    'T1': 'QC', 'T2': 'QC', 'T3': 'QC',
    'T4': 'ON', 'T5': 'ON', 'T6': 'ON',
    'T7': 'BC',
    'T8': 'PR',
    'T9': 'AR',
    'Territory T1': 'QC', 'Territory T2': 'QC', 'Territory T3': 'QC',
    'Territory T4': 'ON', 'Territory T5': 'ON', 'Territory T6': 'ON',
    'Territory T7': 'BC',
    'Territory T8': 'PR',
    'Territory T9': 'AR'
}

model_map = {
    'CO45': 'Outlander',
    'COEV': 'Outlander PHEV',
    'CE45': 'Eclipse Cross',
    'CS45': 'RVR',
    'CG44': 'Mirage'
}

origin_map = {
    'ClickShop': 'ClickShop',
    'MMSCAN Brand Website': 'CWS',
    'Autotrader Storefront': 'AutoTrader',
    'Meta': 'Meta'
}

# --- Page Setup ---
st.set_page_config(page_title="Excel Report Generator", layout="wide")
st.title("📊 Weekly Report Generator")
st.markdown("Upload your **Corporate Leads**, **Footfall**, and **Sales** Excel files to generate clean summary reports by region.")

# --- File Uploads ---
# --- File Uploads ---
st.sidebar.header("Upload Excel Files")

corporate_file = st.sidebar.file_uploader("📂 Corporate Leads File", type=["xlsx"])
footfall_file = st.sidebar.file_uploader("📂 Footfall File", type=["xlsx"])
sales_file = st.sidebar.file_uploader("📂 Sales File", type=["xlsx"])

# --- Sheet Selection ---
corporate_sheet = footfall_sheet = sales_sheet = None

if corporate_file:
    xls = pd.ExcelFile(corporate_file)
    corporate_sheet = st.sidebar.selectbox("📝 Select sheet for Corporate Leads", xls.sheet_names, key="corp")

if footfall_file:
    xls = pd.ExcelFile(footfall_file)
    footfall_sheet = st.sidebar.selectbox("📝 Select sheet for Footfall", xls.sheet_names, key="foot")

if sales_file:
    xls = pd.ExcelFile(sales_file)
    sales_sheet = st.sidebar.selectbox("📝 Select sheet for Sales", xls.sheet_names, key="sales")


# --- Date Inputs ---
start_date = st.sidebar.date_input("📅 Start Date")
end_date = st.sidebar.date_input("📅 End Date")

# --- Processing Functions ---
def process_corporate(upload, sheet_name):
    df = pd.read_excel(upload, sheet_name=sheet_name)
    df = df[df['Origin'].isin(origin_map)]
    df['Origin'] = df['Origin'].map(origin_map)
    df['Region'] = df['Territory'].map(territory_region_map)
    df.dropna(subset=['Origin', 'Region'], inplace=True)
    pivot = pd.pivot_table(df, index='Origin', columns='Region', values='Province', aggfunc='count', fill_value=0)
    pivot['Total'] = pivot.sum(axis=1)
    desired_order = ['Total', 'BC', 'PR', 'ON', 'QC', 'AR']
    return pivot[[col for col in desired_order if col in pivot.columns]]

def process_sales(upload, start, end, sheet_name):
    df = pd.read_excel(upload, sheet_name=sheet_name)
    df['Calendar Date'] = pd.to_datetime(df['Calendar Date'], errors='coerce')
    df = df[(df['Calendar Date'] >= pd.to_datetime(start)) & (df['Calendar Date'] <= pd.to_datetime(end))]
    df = df[df['Model Code'].isin(model_map)]
    df['Model'] = df['Model Code'].map(model_map)
    df['Region'] = df['Terr.'].map(territory_region_map)
    df.dropna(subset=['Model', 'Region'], inplace=True)
    pivot = pd.pivot_table(df, index='Model', columns='Region', values='Retail Count', aggfunc='sum', fill_value=0)
    pivot['Total'] = pivot.sum(axis=1)
    desired_order = ['Total', 'BC', 'PR', 'ON', 'QC', 'AR']
    return pivot[[col for col in desired_order if col in pivot.columns]]

def process_footfall(upload, sheet_name):
    df = pd.read_excel(upload, sheet_name=sheet_name)
    df['Region'] = df['Region'].map(territory_region_map)
    df.dropna(subset=['Region', 'Model', 'Traffic'], inplace=True)
    pivot = pd.pivot_table(df, index='Model', columns='Region', values='Traffic', aggfunc='sum', fill_value=0)
    pivot['Total'] = pivot.sum(axis=1)
    ordered_cols = ['Total', 'BC', 'PR', 'ON', 'QC', 'AR']
    for col in ordered_cols:
        if col not in pivot.columns:
            pivot[col] = 0
    return pivot[ordered_cols]


def download_excel(df, name):
    buffer = BytesIO()
    df.to_excel(buffer, index=True)
    buffer.seek(0)
    return buffer

# --- Processing & Output ---
if corporate_file and footfall_file and sales_file:
    st.success("✅ All files uploaded successfully!")
    
    generate = st.button("🚀 Generate Reports")

    if generate:
        st.subheader("📈 Corporate Leads Report")
        corporate_df = process_corporate(corporate_file, corporate_sheet)
        st.dataframe(corporate_df.style.format(na_rep="0"), use_container_width=True)
        st.download_button("⬇️ Download Corporate Report", download_excel(corporate_df, "Corporate_Report.xlsx"), file_name="Corporate_Report.xlsx")

        st.subheader("🚶 Footfall Report")
        footfall_df = process_footfall(footfall_file, footfall_sheet)
        st.dataframe(footfall_df.style.format(na_rep="0"), use_container_width=True)
        st.download_button("⬇️ Download Footfall Report", download_excel(footfall_df, "Footfall_Report.xlsx"), file_name="Footfall_Report.xlsx")

        st.subheader("🛒 Sales Report")
        sales_df = process_sales(sales_file, start_date, end_date, sales_sheet)
        st.dataframe(sales_df.style.format(na_rep="0"), use_container_width=True)
        st.download_button("⬇️ Download Sales Report", download_excel(sales_df, "Sales_Report.xlsx"), file_name="Sales_Report.xlsx")
else:
    st.info("📄 Please upload all three files to enable the 'Generate Reports' button.")
