import pandas as pd
from datetime import datetime

# Load the main invoice Excel file
invoice_file = "workings_file.xlsx"
df = pd.read_excel(invoice_file, header=1)
df.columns = df.columns.str.strip()

# Load the supplier balance Excel file
supplier_file = "balance_sheet.xlsx"
supplier_df = pd.read_excel(supplier_file)
supplier_df.columns = supplier_df.columns.str.strip()

# Filter out rows with blank supplier
supplier_df = supplier_df[supplier_df["Supplier"].notna() & (supplier_df["Supplier"].astype(str).str.strip() != "")]

# Calculate Net_value_amount = Clsng Blns Debit + Clsng Blns Credit
supplier_df["Net_value_amount"] = supplier_df["Clsng Blns Debit"].fillna(0) + supplier_df["Clsng Blns Credit"].fillna(0)

# Get list of suppliers with Net_value_amount > 0
suppliers_with_balance = supplier_df[supplier_df["Net_value_amount"] > 0]["Supplier"].astype(str).str.strip().unique()

# Define filter criteria for invoice data
exclude_gl_texts = [
    "Intercompany payable", "IOU manager", "IOU staff", "Short Term loan",
    "Trade creditors-Foreign", "Vendors bills of exchange"
]
valid_payment_method = "T"
valid_currency = "NGN"

# Apply filters on invoice file
filtered_df = df[
    ~df["G/L Account: Long Text"].isin(exclude_gl_texts) &
    df["Payment block"].isna() &
    (df["Payment Method"] == valid_payment_method) &
    (df["Currency"] == valid_currency) &
    ~df["Diageo"].astype(str).str.contains("NTC- VENDOR", case=False, na=False) &
    df["Net Due Date"].notna() &
    (df["Due/Not"].astype(str).str.strip().str.lower() == "due")
]

# Exclude suppliers with Net_value_amount > 0
filtered_df = filtered_df[
    ~filtered_df["Supplier"].astype(str).str.strip().isin(suppliers_with_balance)
]

# Save the filtered data
filtered_output_file = "filtered_invoices.xlsx"
filtered_df.to_excel(filtered_output_file, index=False)
print(f"Filtered data saved to: {filtered_output_file}")

# Group and aggregate by Supplier
grouped_df = filtered_df.groupby(["Supplier"], as_index=False).agg({
    "Name": "first",
    "WHT availability": "first",
    "Diageo/Tolaram": "first",
    "Document Currency Value": "sum",
    "Payable after WHT": "sum"
})

grouped_df.rename(columns={
    "Document Currency Value": "Sum of Document Currency Value",
    "Payable after WHT": "Sum of Payable after WHT"
}, inplace=True)

# Save the grouped summary
summary_output_file = "grouped_summary.xlsx"
grouped_df.to_excel(summary_output_file, index=False)
print(f"Grouped summary saved to: {summary_output_file}")
