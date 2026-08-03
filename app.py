import streamlit as st
import pandas as pd

from support import load_excel_preview
from verification import filter_report, run_full_verification

st.set_page_config(page_title="New EPI Verification App", layout="wide")
st.title("New EPI Verification App")
st.write("Upload an Excel workbook to run verification and generate a report.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        data = load_excel_preview(uploaded_file)
        st.success("File loaded successfully")
        st.subheader("Sheets")
        st.write(data["sheet_names"])

        st.subheader("Verification")
        raw_df, report_df = run_full_verification(uploaded_file)

        st.subheader("Report Filters")
        col1, col2, col3 = st.columns(3)
        district_filter = col1.text_input("District")
        township_filter = col2.text_input("Township")
        status_filter = col3.text_input("Status")

        filtered_report = filter_report(report_df, district=district_filter, township=township_filter, status=status_filter)

        st.subheader("Verification Summary")
        st.dataframe(filtered_report.head(50))

        st.subheader("Raw Data Preview")
        st.dataframe(raw_df.head(10))
    except RuntimeError as error:
        st.error(str(error))
