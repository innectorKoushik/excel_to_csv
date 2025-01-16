import streamlit as st
import pandas as pd
from io import BytesIO

def convert_to_csv(df):
    """Convert a DataFrame to CSV bytes."""
    buffer = BytesIO()
    df.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer.getvalue()

def main():
    st.title("Excel to CSV Converter")
    
    # File uploader
    uploaded_file = st.file_uploader("Upload an Excel Workbook", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Load Excel file
        try:
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names
            st.success(f"Successfully loaded workbook with {len(sheet_names)} sheets.")
            
            # Display sheet options
            selected_sheets = st.multiselect(
                "Select sheets to convert to CSV",
                options=sheet_names,
                default=sheet_names
            )
            
            # If sheets are selected
            if selected_sheets:
                st.write("Selected Sheets:", selected_sheets)
                
                csv_buffers = {}
                for sheet in selected_sheets:
                    df = excel_data.parse(sheet_name=sheet)
                    csv_data = convert_to_csv(df)
                    csv_buffers[sheet] = csv_data
                
                # Provide download links
                st.subheader("Download CSV Files")
                for sheet, csv_data in csv_buffers.items():
                    st.download_button(
                        label=f"Download {sheet}.csv",
                        data=csv_data,
                        file_name=f"{sheet}.csv",
                        mime="text/csv"
                    )
            else:
                st.warning("Please select at least one sheet to convert.")
        
        except Exception as e:
            st.error(f"Error processing the file: {e}")

if __name__ == "__main__":
    main()
