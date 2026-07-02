import openpyxl
import pandas as pd
import glob
import os

folder_path = "your_folder_path_here"   # e.g. "C:/Users/you/Documents/data"
cell_address = "B2"                     # the cell you want to extract
output_file = "extracted_data.xlsx"     # name of the new Excel file

results = []

for file_path in sorted(glob.glob(os.path.join(folder_path, "*.xlsx"))):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active                    # or wb["SheetName"] for a specific sheet
    value = sheet[cell_address].value
    results.append({"File Name": os.path.basename(file_path), "Value": value})
    wb.close()

df = pd.DataFrame(results)
df.to_excel(output_file, index=False)

print(f"Saved {len(results)} rows to {output_file}")