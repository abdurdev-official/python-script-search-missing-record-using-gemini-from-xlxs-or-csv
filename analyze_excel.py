"""
Script to examine Excel file and find empty data to fill from Google
"""

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

# Load the Excel file
excel_file = "/Users/admin/pythonnscript/american university data list.xlsx"
wb = openpyxl.load_workbook(excel_file)
ws = wb.active

# Also load with pandas for easier viewing
df = pd.read_excel(excel_file)

print("=" * 80)
print("EXCEL FILE ANALYSIS: american university data list.xlsx")
print("=" * 80)

# Display basic info
print(f"\nShape: {df.shape[0]} rows x {df.shape[1]} columns")
print(f"\nColumns: {list(df.columns)}")

# Show the data
print("\n" + "=" * 80)
print("DATA PREVIEW:")
print("=" * 80)
print(df.to_string())

# Analyze empty cells
print("\n" + "=" * 80)
print("EMPTY CELLS ANALYSIS:")
print("=" * 80)

empty_summary = {}
for col_idx, col_name in enumerate(df.columns, 1):
    empty_count = df[col_name].isna().sum()
    if empty_count > 0:
        empty_summary[col_name] = empty_count
        print(f"\n✗ Column '{col_name}': {empty_count} empty cell(s)")
        
        # Show which rows have empty values
        empty_rows = df[df[col_name].isna()].index.tolist()
        for row_idx in empty_rows:
            print(f"   Row {row_idx + 2}: {df.iloc[row_idx].to_dict()}")

if not empty_summary:
    print("✓ No empty cells found!")
else:
    print(f"\n{'=' * 80}")
    print(f"TOTAL EMPTY CELLS: {sum(empty_summary.values())}")
    print(f"{'=' * 80}")
