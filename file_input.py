import streamlit as st
import pandas as pd

# Title of the app
st.title("RPA Assignment - Sort Requisition by Date")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read the Excel file
    try:
        df = pd.read_excel(uploaded_file)
        st.write("Data from the uploaded requisition form:")
        st.dataframe(df)

        # Check if 'Required Date' column exists
        if 'Required Date' in df.columns:
            # Convert 'Required Date' to datetime
            df['Required Date'] = pd.to_datetime(df['Required Date'])

            # Sort the DataFrame by 'Required Date'
            sorted_requisition = df.sort_values(by='Required Date')
            st.write("Sorted Requisition Data by Date:")
            st.dataframe(sorted_requisition)
        else:
            st.error("The 'Required Date' column is not found in the uploaded file.")

    except Exception as e:
        st.error(f"Error reading the file: {e}")