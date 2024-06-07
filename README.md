# Purview_metadata

### Prerequisites

1. **Azure Account**: Ensure you have an Azure account with the necessary permissions to create and manage Microsoft Purview resources.
2. **Microsoft Purview**: Set up Microsoft Purview in your Azure environment.
3. **Azure CLI**: Ensure you have Azure CLI installed and configured on your machine.
4. **Azure SDK for Python**: Install the necessary Python packages.

```sh
pip install azure-identity azure-purview-catalog azure-mgmt-purview pandas openpyxl
```

### Step-by-Step Guide

#### 1. Set Up Credentials

Create a `credentials.txt` file to store your Azure credentials. This file should contain:

```plaintext
AZURE_CLIENT_ID=your_client_id
AZURE_TENANT_ID=your_tenant_id
AZURE_CLIENT_SECRET=your_client_secret
AZURE_SUBSCRIPTION_ID=your_subscription_id
PURVIEW_ACCOUNT_NAME=your_purview_account_name
```

Replace `your_client_id`, `your_tenant_id`, `your_client_secret`, `your_subscription_id`, and `your_purview_account_name` with your actual Azure credentials.

### Script 1: Fetch Metadata and Insights, Save to Excel

This script connects to Microsoft Purview, scans data sources, retrieves metadata and insights, and saves them into an Excel file named `purview_data.xlsx` with additional `_Y` columns.

#### Functions

- **read_credentials(file_path)**: Reads the credentials from the `credentials.txt` file.
- **connect_to_purview(credentials)**: Authenticates and connects to Microsoft Purview using the `azure-identity` and `azure-purview-catalog` libraries.
- **scan_data_sources(purview_client)**: Retrieves the list of data sources and initiates scans on each one.
- **get_metadata(purview_client)**: Fetches metadata from Microsoft Purview.
- **get_data_insights(purview_client)**: Fetches data insights from Microsoft Purview.
- **adjust_column_widths(sheet)**: Adjusts the column widths in an Excel sheet to fit the text.
- **add_y_columns(dataframe)**: Adds `_Y` columns to a DataFrame for each existing column, populating with "Y" if the original column has a non-null value.

#### Usage

1. Ensure the `credentials.txt` file is in the same directory as the script.
2. Run the script to fetch metadata and insights, and save them into an Excel file:

```python
python fetch_purview_metadata.py
```

This will create a file called `purview_data.xlsx` containing the metadata and insights with additional `_Y` columns.

### Script 2: Filter Data Based on `_Y` Columns

This script filters the data from the `purview_data.xlsx` file, keeping only the columns where the corresponding `_Y` column contains 'Y'. It then writes the filtered data to a new Excel file, discarding all `_Y` columns.

#### Functions

- **filter_columns_with_Y(df)**: Filters the DataFrame by keeping only those columns where the corresponding `_Y` column contains 'Y', discarding the `_Y` columns.
- **adjust_column_widths(sheet)**: Adjusts the column widths in an Excel sheet to fit the text.
- **process_excel_file(input_file, output_file)**: Reads the original Excel file, processes each sheet to filter columns, and writes the filtered data to a new Excel file. Adjusts the column widths in the output file.

#### Usage

1. Ensure the `purview_data.xlsx` file created by the first script is in the same directory as this script.
2. Run the script to filter the data:

```python
python filter_purview_data.py
```

This will create a new file called `filtered_purview_data.xlsx` containing only the columns where the corresponding `_Y` column has 'Y' in it, and discarding all `_Y` columns.

### Summary

1. **Set Up Credentials**: Store your Azure credentials in a `credentials.txt` file.
2. **Fetch Metadata and Insights**: Use the first script to connect to Microsoft Purview, scan data sources, retrieve metadata and insights, and save them into an Excel file (`purview_data.xlsx`).
3. **Filter Data**: Use the second script to filter the data based on the `_Y` columns, keeping only the columns where the corresponding `_Y` column contains 'Y'. Save the filtered data into a new Excel file (`filtered_purview_data.xlsx`).

