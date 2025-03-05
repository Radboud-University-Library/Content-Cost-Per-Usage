# README

## Overview
This script calculates the **cost per use per journal title** by merging **Total_Item_Requests** from an aggregated **Counter 5 Title Master Report** with a **custom WMS Report Designer report** containing journal titles, ISSNs, and invoice data. The output is a processed file that provides insights for data analysis.

### **Important Note**
When running the script, you may encounter the following warning:
```
UserWarning: Workbook contains no default style, apply openpyxl's default
warn("Workbook contains no default style, apply openpyxl's default")
```
This warning does not affect the script's functionality and can be ignored.

---
## Folder Paths
The script processes files stored in specific folders. These paths are defined in the script:

- **READ_FOLDER_PATH**: Directory containing the input files.
- **WRITE_FOLDER_PATH**: Directory where the processed output file is saved.

```python
READ_FOLDER_PATH = r"C:\Users\Documents\Python_read_folder"
WRITE_FOLDER_PATH = r"C:\Users\Documents\Python_write_folder"
```

---
## Input Files
The script processes the following input files:
1. **Counter 5 Report** (`Example.xlsx`): Contains journal usage data.
2. **WMS Report** (`Example.xlsx`): Contains journal titles, ISSNs, and invoice data.

---
## Key Processing Steps
### **1. Process Counter 5 Data**
- Reads the **Counter 5 Title Master Report**.
- Filters data for `Metric_Type == 'Total_Item_Requests'`.
- Computes the total number of item requests per title.
- Groups data by **Title**, **Print_ISSN**, and **Online_ISSN**.
- Creates a **Combined_ISSNs** column for ISSN matching.

### **2. Process WMS Report**
- Reads the **WMS Report** and skips empty rows.
- Cleans up data by removing unnecessary columns.
- Adds **Total_Item_Requests** and **Comment** columns.
- Formats the **ISSN** column by replacing commas with pipe separators (`|`).
- Iterates through ISSNs to find matches in the Counter 5 report.
- Computes total item requests per journal title.

### **3. Calculate Cost per Use**
- Groups invoices to sum multiple invoices for a single journal title.
- Computes **Cost per Use** as:

  ```
  Cost per Use = Invoice Amount (Institution Currency) / Total Item Requests
  ```
- Rounds cost per use values to **two decimal places** and uses `,` as a decimal separator.
- Reorders columns for clarity.

### **4. Save Processed Data**
The final processed data is saved as:
```
Content_Cost_Per_Usage.xlsx
```
in the specified output folder.

---
## Dependencies
Ensure the following Python packages are installed before running the script:

```sh
pip install pandas openpyxl
```

---
## Running the Script
Execute the script using Python:
```sh
python main.py
```

---
## Troubleshooting
- Ensure that input files (`Example.xlsx`) exist in the `READ_FOLDER_PATH`.
- If `NaN` values appear in **Cost per Use**, check for missing **Total_Item_Requests** values.
- If the script encounters an **ISSN not found in Counter data**, review the ISSN formatting in the input files.

---
## License
This project is open-source and free to use under the MIT License.

