import pandas as pd
 
# ===========================
# Configuration
# ===========================
 
FILE_1 = "file1.xlsx"
FILE_2 = "file2.xlsx"
 
SHEET_1 = "Sheet1"
SHEET_2 = "Sheet1"
 
# Column that uniquely identifies each record
KEY_COLUMN = "Employee ID"
 
OUTPUT_FILE = "excel_comparison_result.xlsx"
 
 
# ===========================
# Read Excel files
# ===========================
 
df1 = pd.read_excel(FILE_1, sheet_name=SHEET_1)
df2 = pd.read_excel(FILE_2, sheet_name=SHEET_2)
 
# Remove unwanted spaces from column names
df1.columns = df1.columns.str.strip()
df2.columns = df2.columns.str.strip()
 
 
# ===========================
# Validate column structures
# ===========================
 
columns_file1 = set(df1.columns)
columns_file2 = set(df2.columns)
 
missing_in_file2 = columns_file1 - columns_file2
missing_in_file1 = columns_file2 - columns_file1
 
print("Columns missing in File 2:", missing_in_file2)
print("Columns missing in File 1:", missing_in_file1)
 
if KEY_COLUMN not in df1.columns or KEY_COLUMN not in df2.columns:
    raise ValueError(
        f"Key column '{KEY_COLUMN}' must be present in both Excel files."
    )
 
 
# ===========================
# Set unique key as index
# ===========================
 
df1 = df1.set_index(KEY_COLUMN)
df2 = df2.set_index(KEY_COLUMN)
 
 
# ===========================
# Find missing / new records
# ===========================
 
only_in_file1 = df1.loc[df1.index.difference(df2.index)]
only_in_file2 = df2.loc[df2.index.difference(df1.index)]
 
 
# ===========================
# Compare common records
# ===========================
 
common_keys = df1.index.intersection(df2.index)
common_columns = df1.columns.intersection(df2.columns)
 
df1_common = df1.loc[common_keys, common_columns]
df2_common = df2.loc[common_keys, common_columns]
 
differences = []
 
for key in common_keys:
 
    for column in common_columns:
 
        value1 = df1_common.loc[key, column]
        value2 = df2_common.loc[key, column]
 
        # Treat two blank/NaN values as equal
        if pd.isna(value1) and pd.isna(value2):
            continue
 
        # Compare values
        if str(value1).strip() != str(value2).strip():
 
            differences.append({
                KEY_COLUMN: key,
                "Column": column,
                "File 1 Value": value1,
                "File 2 Value": value2
            })
 
 
difference_df = pd.DataFrame(differences)
 
 
# ===========================
# Summary
# ===========================
 
summary = pd.DataFrame({
    "Validation": [
        "Total records in File 1",
        "Total records in File 2",
        "Records only in File 1",
        "Records only in File 2",
        "Different field values"
    ],
    "Count": [
        len(df1),
        len(df2),
        len(only_in_file1),
        len(only_in_file2),
        len(difference_df)
    ]
})
 
 
# ===========================
# Export results
# ===========================
 
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
 
    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )
 
    difference_df.to_excel(
        writer,
        sheet_name="Differences",
        index=False
    )
 
    only_in_file1.reset_index().to_excel(
        writer,
        sheet_name="Only in File 1",
        index=False
    )
 
    only_in_file2.reset_index().to_excel(
        writer,
        sheet_name="Only in File 2",
        index=False
    )
 
 
print("\nComparison completed successfully.")
print(f"Result created: {OUTPUT_FILE}")