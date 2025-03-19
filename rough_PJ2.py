import pandas as pd
import glob
import os
import xlsxwriter

# 📂 Folder containing CSV files
csv_folder = r"D:\Project2_data_automation\rough_csv"
csv_files = glob.glob(csv_folder + r"\*.csv")  # Get all CSV file paths

# 🎯 Column mappings for standardization
column_mapping = {
    "Product": ["Product", "Item", "Prodcut Name", "Product Name"],
    "Quantity": ["Quantity", "Quantity Sold"],
    "Price": ["Price", "Total", "Price Per Unit"]
}

# 🔄 Function to standardize column names
def standardize_columns(df):
    standardized_cols = {}
    for standard_name, variations in column_mapping.items():
        for col in df.columns:
            if col in variations:
                standardized_cols[col] = standard_name  # Map to standardized name
    df.rename(columns=standardized_cols, inplace=True)  # Apply renaming

    # 🛠 Ensure all required columns exist
    for required_col in column_mapping.keys():
        if required_col not in df.columns:
            df[required_col] = None  # Fill missing columns with NaN

    # 🚀 Keep only the necessary columns
    df = df[list(column_mapping.keys())]
    return df

# 📌 Initialize empty list for storing DataFrames
dataframes = []
crash_entries = []  # Store rows with missing or incorrect values

# 📊 Process each CSV file
for file in csv_files:
    df = pd.read_csv(file)  # Read file
    df = standardize_columns(df)  # Standardize column names

    # 🔍 Convert columns to numeric and handle invalid values
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce")  

    # 🔍 Identify rows with missing or invalid values
    invalid_rows = df[df.isnull().any(axis=1)]
    if not invalid_rows.empty:
        crash_entries.append((file, invalid_rows))

    # 🚀 Drop rows with missing values
    df.dropna(inplace=True)

    dataframes.append(df)  # Append cleaned DataFrame

# 🏆 Merge all DataFrames
if dataframes:
    combined_df = pd.concat(dataframes, ignore_index=True)
else:
    print("⚠️ No valid data found in CSV files. Exiting...")
    exit()

# 📜 Save missing & incorrect rows into a text file
if crash_entries:
    with open("crash_entries.txt", "w", encoding="utf-8") as f:
        for file, entry in crash_entries:
            f.write(f"📂 File: {os.path.basename(file)}\n")  # Add filename
            f.write(entry.to_string(index=False) + "\n\n")  # Append crash entries
            f.write("="*50 + "\n\n")  # Separator for readability


# 📂 Export cleaned and merged data
# 🏆 Calculate Total Sales per product
summary = combined_df.groupby("Product").agg({"Quantity": "sum", "Price": "sum"}).reset_index()
summary["Total Sales"] = summary["Quantity"] * (summary["Price"] / summary["Quantity"])  # Adjusted Calculation

# Get the top 5 selling products
top_5 = summary.nlargest(5, "Total Sales")

# Prepare final DataFrame with two blank rows
empty_rows = pd.DataFrame([["", 0, 0, 0]], columns=top_5.columns)  # One blank row
empty_rows_twice = pd.concat([empty_rows] * 2, ignore_index=True)  # Two blank rows

# Append "Total Sales" row at the bottom
total_sales_row = pd.DataFrame([["Total Sales", summary["Quantity"].sum(), summary["Price"].sum(), summary["Total Sales"].sum()]], columns=top_5.columns)

# Final DataFrame for Excel
final_report = pd.concat([top_5, empty_rows_twice, total_sales_row], ignore_index=True)

# Dynamically set output file path
script_dir = os.path.dirname(os.path.abspath(__file__))  # Get script location
output_file = os.path.join(script_dir, "Top_5_Selling_Products.xlsx")

# Export to Excel
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    final_report.to_excel(writer, sheet_name="Summary", index=False)  # Summary Sheet
    combined_df.to_excel(writer, sheet_name="Raw Data", index=False)  # Raw Data Sheet

print(f"✅ Excel report saved as {output_file} with two sheets.")
print("✅ Data processing complete! Cleaned data saved as 'merged_sales_data.csv'.")
print("⚠️ Missing and incorrect values saved in 'crash_entries.txt'.")
