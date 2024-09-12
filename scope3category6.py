import streamlit as st
import pandas as pd
import io

# Function to process the Excel file
def process_excel(file):
    # Load the Excel file
    excel_data = pd.ExcelFile(file)
    
    # Define the specified sheets
    specified_sheets = ['SSLL', 'OEL Aviation', 'OEL', 'TWSM', 'TILPL', 'DWC', 'TLPL', 'TSLPL', 'FZE', 'TWH', 'TALPL']
    
    # Initialize an empty DataFrame for storing the merged data
    merged_data = pd.DataFrame()
    
    # Loop through the specified sheet names and merge them
    for sheet_name in specified_sheets:
        if sheet_name in excel_data.sheet_names:
            df = pd.read_excel(file, sheet_name=sheet_name)
            merged_data = pd.concat([merged_data, df], ignore_index=True)
    
    # Define the path for the template workbook
    template_workbook_path = r'Air-Sample.xlsx'
    
    # Define column mapping
    column_mapping = {
        'Departure': 'Departure City',
        'Facility': 'Office/Factory/Site/\nLocation(Optional)',
        'Arrival': 'Arrival City',
        'Start Date': 'Start Date (DD/MM/YYYY Format)',
        "End Date": "End Date (DD/MM/YYYY Format)",
        "Cabin Class": "Class of Travel"
    }
    
    # Load the template workbook and get the specified sheet
    template_df = pd.read_excel(template_workbook_path, sheet_name=None)
    template_sheet_name = 'Import data file_Manufacturing'
    template_data = template_df[template_sheet_name]
    
    # Preserve the first row (header) of the template
    preserved_header = template_data.iloc[:0, :]
    
    # Create a DataFrame with the template columns
    matched_data = pd.DataFrame(columns=template_data.columns)
    
    # Map and copy data based on column_mapping
    for template_col, client_col in column_mapping.items():
        if client_col in merged_data.columns:
            matched_data[template_col] = merged_data[client_col]
        else:
            st.write(f"Column '{client_col}' not found in merged_data")
    
    # Combine header and matched data
    final_data = pd.concat([preserved_header, matched_data], ignore_index=True)
    final_data['CF Standard'] = "IATA"
    final_data['Res_Date'] = "30/03/2024"
    final_data['Activity Unit'] = "kWh"
    final_data['Round Trip'] = "No"
    final_data['Gas'] = "CO2"
    final_data['Res_Date'] = pd.to_datetime(final_data['Res_Date']).dt.date
    
    # Strip any leading/trailing whitespace and replace any non-breaking spaces
    final_data['Facility'] = final_data['Facility'].str.strip().replace('\xa0', ' ', regex=True)
    
    # Attempt to drop NaNs
    final_data = final_data.dropna(subset=['Facility'])
    
    # Save final data to a buffer and return it
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, mode='xlsx') as writer:
        final_data.to_excel(writer, sheet_name='Import data file_Manufacturing', index=False)
    
    buffer.seek(0)
    return buffer

# Streamlit UI
st.title('Excel Data Processing App')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    processed_file = process_excel(uploaded_file)
    st.download_button(
        label="Download Processed Data",
        data=processed_file,
        file_name="processed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
