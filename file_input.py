import streamlit as st
import pandas as pd

# Title of the app
st.title("Excel File Upload")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read the Excel file
    try:
        df = pd.read_excel(uploaded_file)  # Corrected indentation
        st.write("Data from the uploaded file:")  # Corrected indentation
        st.dataframe(df)  # Display the dataframe
    except Exception as e:
        st.error(f"Error reading the file: {e}")  # Corrected indentation