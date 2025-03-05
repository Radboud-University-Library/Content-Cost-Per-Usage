import pandas as pd
import openpyxl
import os

"""
This script creates a cost per use per journal title for data analysis by 
adding Total_Item_Requests from a aggregated Counter 5 Title Master Report to 
a custom WMS Report Designer report containing journal titles, ISSNs and 
invoice data.

An error occurs when the script is run. You can ignore this warning: 
    UserWarning: Workbook contains no default style, apply openpyxl's default
    warn("Workbook contains no default style, apply openpyxl's default")
    
Define READ and WRITE folder paths.
"""

# Declare READ and WRITE folder paths where files are placed.
READ_FOLDER_PATH = r"C:\Users\Documents\Python_read_folder"
WRITE_FOLDER_PATH = r"C:\Users\Documents\Python_write_folder"

# Define Counter file name which contains usage data
counter_file_name = "Example.xlsx"

# Define report file name which contains content data
wms_file_name = "Example.xlsx"

"""
Edit Counter df.
"""

# Read aggregated Counter 5 Title Master Report.
counter_df = pd.read_excel(f"{READ_FOLDER_PATH}\\{counter_file_name}")

# Filter df on Metric_Type Total_Item_Requests.
counter_df = counter_df.loc[counter_df["Metric_Type"] == "Total_Item_Requests"]

# Add new column with sum of all monthly columns. Edit column names to represent the correct year.
counter_df["Reporting_Period_Total"] = counter_df[['2024-01', '2024-02', '2024-03', '2024-04', '2024-05', '2024-06',
    '2024-07', '2024-08', '2024-09', '2024-10', '2024-11', '2024-12']].sum(axis=1)

# Convert Title column type to string
counter_df["Title"] = counter_df["Title"].astype(str)

# Group columns by Title, Print_ISSN, Online_ISSN and sum Reporting_Period_Total.
counter_df = counter_df.groupby("Title").agg({
    "Print_ISSN": "first",
    "Online_ISSN": "first",
    "Reporting_Period_Total": "sum"
}).reset_index()

# Create a new column that contains lists of values from Print_ISSN and Online_ISSN
counter_df['Combined_ISSNs'] = counter_df.apply(lambda row: [row['Print_ISSN'], row['Online_ISSN']], axis=1)

# Create Combined ISSN string column to group by.
counter_df['Combined_ISSNs_str'] = counter_df['Combined_ISSNs'].astype(str)

# Group by Combined ISSNs string column.
counter_df = counter_df.groupby('Combined_ISSNs_str').agg({
    'Title': 'first',
    'Combined_ISSNs': 'first',
    'Reporting_Period_Total': 'sum'
}).reset_index()

"""
Edit WMS df.
"""

# Read WMS Report from read folder and skip empty rows.
wms_df = pd.read_excel(f"{READ_FOLDER_PATH}\\{wms_file_name}",skiprows=3)

# Remove first empty column.
wms_df.drop(["Unnamed: 0"], axis=1, inplace=True)

# Create new columns in WMS df.
wms_df["Total_Item_Requests"] = None
wms_df["Comment"] = None

# Replace comma separator with pipe separator.
wms_df["ISSN"] = wms_df["ISSN"].str.replace(",", "|")

# Iterate over WMS ISSN column.
for index, issn in wms_df.iterrows():
    # Split ISSN column to create lists of ISSNs.
    issn_list = issn["ISSN"].split("|")

    # Set default variables.
    total_requests = 0
    match_found = False
    matched_indices = set()

    # Iterate over lists of ISSNs.
    for issn in issn_list:
        # Search for matches with Counter ISSNs.
        matches = counter_df[counter_df["Combined_ISSNs"].apply(lambda x: issn in x)]

        # Iterate over WMS ISSN and Counter ISSN matches.
        for match_index, match_row in matches.iterrows():
            # Sum Requests for unique match indices.
            if match_index not in matched_indices:
                match_found = True
                total_requests += match_row["Reporting_Period_Total"]
                matched_indices.add(match_index)

    # Add sum of Requests to WMS df.
    if match_found:
        wms_df.at[index, "Total_Item_Requests"] = total_requests
    # Add comment to WMS df.
    else:
        wms_df.at[index, "Comment"] = "ISSN not found in Counter data."

# Group invoices to account for multiple invoices for one title.
invoice_sum_dict = wms_df.groupby('Title')['Invoice Amount (Institution Currency)'].sum().to_dict()
# Update the 'Invoice Amount (Institution Currency)' column with the summed values.
wms_df['Invoice Amount (Institution Currency)'] = wms_df['Title'].map(invoice_sum_dict)

# Calculate cost per use.
wms_df["Cost per use"] = wms_df["Invoice Amount (Institution Currency)"] / wms_df["Total_Item_Requests"]
# Round cost per use to two decimal numbers.
wms_df["Cost per use"] = wms_df["Cost per use"].apply(lambda item: f"{item:.2f}".replace('.', ',') if pd.notnull(item) else '')

# Reorder columns.
wms_df = wms_df[["Fund Name Level 1","Title","Invoice Amount (Vendor Currency)","Invoice Currency","Invoice Exchange Rate",
                "Invoice Amount (Institution Currency)","ISSN","Total_Item_Requests","Cost per use","Comment"]]

# Write file to write folder.
new_file_path = os.path.join(WRITE_FOLDER_PATH, "Content_Cost_Per_Usage.xlsx")
wms_df.to_excel(new_file_path, index=False)
