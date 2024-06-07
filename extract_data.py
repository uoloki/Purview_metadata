import os
import pandas as pd
from azure.identity import ClientSecretCredential
from azure.purview.catalog import PurviewCatalogClient
from azure.mgmt.purview import PurviewManagementClient
from openpyxl import load_workbook

# Function to read credentials from a file
def read_credentials(file_path):
    credentials = {}
    with open(file_path, 'r') as file:
        for line in file:
            key, value = line.strip().split('=')
            credentials[key] = value
    return credentials

# Function to connect to Purview
def connect_to_purview(credentials):
    try:
        credential = ClientSecretCredential(
            tenant_id=credentials['AZURE_TENANT_ID'],
            client_id=credentials['AZURE_CLIENT_ID'],
            client_secret=credentials['AZURE_CLIENT_SECRET']
        )
        
        purview_account_name = credentials['PURVIEW_ACCOUNT_NAME']
        purview_client = PurviewCatalogClient(
            endpoint=f"https://{purview_account_name}.purview.azure.com",
            credential=credential
        )
        
        return purview_client
    except Exception as e:
        print(f"Error connecting to Microsoft Purview: {e}")
        return None

# Function to scan data sources
def scan_data_sources(purview_client):
    try:
        # Retrieve the list of data sources
        response = purview_client.discovery.query("SELECT * FROM sys.databases")
        
        # For each data source, initiate a scan
        scan_results = []
        for data_source in response:
            data_source_name = data_source['name']
            # Create a scan for the data source
            scan_configuration = {
                "properties": {
                    "scanRulesetName": "defaultScanRuleset",
                    "scanRulesetType": "System",
                    "scanTriggerType": "OnDemand"
                }
            }
            scan_response = purview_client.discovery.scan.create_or_update_scan(data_source_name, scan_configuration)
            scan_results.append(scan_response)
        
        return scan_results
    except Exception as e:
        print(f"Error initiating scan: {e}")
        return []

# Function to retrieve metadata
def get_metadata(purview_client):
    try:
        metadata = []
        # Fetch metadata
        response = purview_client.discovery.query("SELECT * FROM sys.databases")
        metadata.extend(response)
        return metadata
    except Exception as e:
        print(f"Error fetching metadata: {e}")
        return []

# Function to retrieve data insights
def get_data_insights(purview_client):
    try:
        insights = []
        # Fetch data insights
        response = purview_client.discovery.query("SELECT * FROM sys.databases")
        insights.extend(response)
        return insights
    except Exception as e:
        print(f"Error fetching data insights: {e}")
        return []

# Function to adjust the column width
def adjust_column_widths(sheet):
    for column_cells in sheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        sheet.column_dimensions[column_cells[0].column_letter].width = length + 2

# Function to add _Y columns and populate with "Y"
def add_y_columns(dataframe):
    for col in dataframe.columns:
        dataframe[f"{col}_Y"] = dataframe[col].apply(lambda x: 'Y' if pd.notna(x) else '')
    return dataframe

if __name__ == "__main__":
    try:
        credentials = read_credentials('credentials.txt')
        purview_client = connect_to_purview(credentials)
        
        if purview_client:
            scan_results = scan_data_sources(purview_client)
            metadata = get_metadata(purview_client)
            insights = get_data_insights(purview_client)
            
            # Convert scan results, metadata, and insights to DataFrames
            scan_results_df = pd.DataFrame(scan_results)
            metadata_df = pd.DataFrame(metadata)
            insights_df = pd.DataFrame(insights)
            
            # Add _Y columns to the right
            scan_results_df = add_y_columns(scan_results_df)
            metadata_df = add_y_columns(metadata_df)
            insights_df = add_y_columns(insights_df)
            
            # Write to Excel with separate sheets
            with pd.ExcelWriter('purview_data.xlsx', engine='openpyxl') as writer:
                scan_results_df.to_excel(writer, sheet_name='Scan Results', index=False)
                metadata_df.to_excel(writer, sheet_name='Metadata', index=False)
                insights_df.to_excel(writer, sheet_name='Insights', index=False)
                
                adjust_column_widths(writer.sheets['Scan Results'])
                adjust_column_widths(writer.sheets['Metadata'])
                adjust_column_widths(writer.sheets['Insights'])
            
            print("Data has been written to 'purview_data.xlsx'")
    except Exception as e:
        print(f"An error occurred: {e}")
