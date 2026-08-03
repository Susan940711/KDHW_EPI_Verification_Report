import streamlit as st
from support import load_excel_preview

st.set_page_config(page_title="New EPI Verification App", layout="wide")
st.title("New EPI Verification App")
st.write("Upload an Excel workbook to preview its contents.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    data = load_excel_preview(uploaded_file)
    st.success("File loaded successfully")
    st.subheader("Sheets")
    st.write(data["sheet_names"])

    st.subheader("Preview")
    st.dataframe(data["preview"].head(10))
