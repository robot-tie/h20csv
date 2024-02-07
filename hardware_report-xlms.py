import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
import os
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime

# Keywords for unsupported OS versions (simplified to substrings to look for)
unsupported_os_keywords = [' 7', 'XP', ' 8', '2012', '2008']  # Spaces help avoid false positives, e.g., "8" in "Windows 10"

# Function to add recommendations based on OS substring
def add_recommendations(row):
    # Check if any of the unsupported OS keywords is in the 'os' column value
    for keyword in unsupported_os_keywords:
        if keyword in row['os']:
            return "replace"
    return ""

# Reload the CSV file to ensure the data is available
df_new = pd.read_csv('272024-asset-tracking-dump.csv')

# Apply the recommendation function to each row
df_new['recommendation'] = df_new.apply(add_recommendations, axis=1)

# Ensure the 'recommendation' column is included in the columns to extract
columns_to_extract = ['Client Name', 'site', 'chassis', 'ip', 'deviceserial', 'name',
                      'manufacturer', 'model', 'os', 'osinstalldate', 'role', 'recommendation']

# Function to create Excel file for each client
def save_to_excel(dataframe, client_names, directory, columns):
    os.makedirs(directory, exist_ok=True)
    for client_name in client_names:
        # Filter data for the current client
        client_data = dataframe[dataframe['Client Name'] == client_name][columns]

        # Create a new Excel workbook and select the active worksheet
        wb = Workbook()
        ws = wb.active

        # Append column titles with bold font
        for col_num, column_title in enumerate(columns, start=1):
            ws.cell(row=1, column=col_num, value=column_title).font = Font(bold=True)

        # Append the rows of the dataframe to the worksheet, including recommendations
        for row in dataframe_to_rows(client_data, index=False, header=False):
            ws.append(row)

        # Define file name (replace spaces and special characters for compatibility)
        current_date = datetime.datetime.now()
        date_str = current_date.strftime("%Y-%m-%d-")
        safe_client_name = client_name.replace(' ', '_').replace("'", "").replace('/', '_')
        file_name = f"{date_str}{safe_client_name}.xlsx"
        file_path = os.path.join(directory, file_name)

        # Save the workbook
        wb.save(file_path)

if __name__ == "__main__":
    # Define the output directory for the Excel files
    output_dir_excel = 'individual_client_files_excel'

    # Extract all unique client names
    unique_client_names = df_new['Client Name'].unique()

    # Call the function to create Excel files
    save_to_excel(df_new, unique_client_names, output_dir_excel, columns_to_extract)

    print(f"Excel files created in '{output_dir_excel}' with specified columns and bold headings.")
