import subprocess

subprocess.check_call(["pip", "install", "openpyxl"])

import pandas as pd
import openpyxl
import streamlit as st


# Load the Excel file
@st.cache_data
def load_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        sheets_data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in sheet_names}
        return sheets_data
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return {}

# Main app
def main():
    st.markdown("<h1 style='text-align: center; color: #4CAF50;'>Retrofit Strategies Database</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Explore energy retrofit strategies tailored to different climate classifications. Search and filter below!</p>", unsafe_allow_html=True)

    # File uploader for Excel file
    uploaded_file = st.file_uploader("Upload an Excel file:", type=["xlsx"])
    if uploaded_file is not None:
        sheets = load_data(uploaded_file)

        if sheets:
            # Choose sheet
            sheet_choice = st.selectbox("Select a category:", list(sheets.keys()))

            df = sheets[sheet_choice]

            # Search box
            search = st.text_input("Search for keywords (e.g., insulation, HVAC, lighting):", "")
            if search:
                # Optimized search using pandas' string filtering
                df = df[df.apply(lambda row: row.astype(str).str.contains(search, case=False, na=False).any(), axis=1)]

            st.dataframe(df, use_container_width=True)
        else:
            st.warning("No data found in the uploaded file.")
    else:
        st.info("Please upload an Excel file to get started.")

# Run the app
if __name__ == "__main__":
    main()
