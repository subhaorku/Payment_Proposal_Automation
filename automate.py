import pandas as pd
from datetime import datetime

# Load the Excel file
input_file = "workings_file.xlsx"  # Make sure this file is in the same directory
df = pd.read_excel(input_file, header=1)

# Clean column names
df.columns = df.columns.str.strip()

# Define filter criteria
exclude_gl_texts = [
    "Intercompany payable", "IOU manager", "IOU staff", "Short Term loan",
    "Trade creditors-Foreign", "Vendors bills of exchange"
]
valid_payment_method = "T"
valid_currency = "NGN"

# Apply all filters
filtered_df = df[
    ~df["G/L Account: Long Text"].isin(exclude_gl_texts) &
    df["Payment block"].isna() &
    (df["Payment Method"] == valid_payment_method) &
    (df["Currency"] == valid_currency) &
    ~df["Diageo"].astype(str).str.contains("NTC- VENDOR", case=False, na=False) &
    df["Net Due Date"].notna() &
    (df["Due/Not"].astype(str).str.strip().str.lower() == "due")
]

# Save to a new Excel file
output_file = "filtered_invoices.xlsx"
filtered_df.to_excel(output_file, index=False)

print(f"Filtered data saved to: {output_file}")

# Group and aggregate by Supplier
grouped_df = filtered_df.groupby(["Supplier"], as_index=False).agg({
    "Name": "first",  # First Name entry per Supplier
    "WHT availability": "first",  # First WHT availability per Supplier
    "Diageo/Tolaram": "first", 
    "Document Currency Value": "sum",
    "Payable after WHT": "sum"
})

# Rename columns to match desired output format
grouped_df.rename(columns={
    "Document Currency Value": "Sum of Document Currency Value",
    "Payable after WHT": "Sum of Payable after WHT"
}, inplace=True)

# Save the grouped summary to a new Excel file
summary_output_file = "grouped_summary.xlsx"
grouped_df.to_excel(summary_output_file, index=False)

print(f"Grouped summary saved to: {summary_output_file}")
