import pandas as pd
import xml.etree.ElementTree as ET
import os

# Load and parse the XML file
xml_file_path = '272024-asset-tracking-dump.xml'
tree = ET.parse(xml_file_path)
root = tree.getroot()

# Initialize a list to store each row of data
data_rows = []

# Iterate over each 'client' element in the XML
for client in root.findall('client'):
    client_name = client.attrib['name']
    for asset in client:
        # Create a dictionary for each asset, starting with the client name
        asset_data = {'Client Name': client_name}
        # Add additional attributes and elements from the asset as needed
        for detail in asset:
            asset_data[detail.tag] = detail.text
        data_rows.append(asset_data)

# Convert the list of dictionaries to a DataFrame
df_assets = pd.DataFrame(data_rows)

# Specify the columns you want to include in the final CSV
desired_columns = ['Client Name', 'site', 'chassis', 'ip', 'deviceserial', 'name', 'manufacturer', 'model', 'os', 'osinstalldate', 'role']

# Filter the DataFrame to include only the desired columns
df_filtered = df_assets[desired_columns]

# Save the DataFrame to a CSV file
csv_file_path = '272024-asset-tracking-dump.csv'
df_filtered.to_csv(csv_file_path, index=False)

print(f"CSV file has been created at: {csv_file_path}")
