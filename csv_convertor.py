import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from io import BytesIO

def convert_csv_to_excel(df):
    # Create a new Excel workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active

    # Write the DataFrame to the Excel worksheet
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    # Adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column_letter].width = adjusted_width

    # Save to BytesIO object
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    return excel_buffer

def main():
    st.set_page_config(
        page_title="CSV to Excel Converter",
        page_icon="📊",
        layout="centered"
    )

    st.title("CSV to Excel Converter")
    st.write("Upload your CSV file and convert it to Excel format")

    # File uploader
    uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

    if uploaded_file is not None:
        try:
            # Read CSV file
            df = pd.read_csv(uploaded_file)
            
            # Show preview of the data
            st.subheader("Preview of your CSV data")
            st.dataframe(df.head())
            
            # Show basic information about the data
            st.subheader("Data Information")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Rows", df.shape[0])
            with col2:
                st.metric("Columns", df.shape[1])
            with col3:
                st.metric("Size (KB)", round(uploaded_file.size/1024, 2))

            # Convert button
            if st.button("Convert to Excel", type="primary"):
                with st.spinner("Converting..."):
                    # Convert to Excel
                    excel_buffer = convert_csv_to_excel(df)
                    
                    # Generate download button
                    st.download_button(
                        label="📥 Download Excel file",
                        data=excel_buffer.getvalue(),
                        file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("✅ Conversion completed! Click the download button above to get your Excel file.")

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
