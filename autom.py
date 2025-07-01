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

# Group by Supplier and calculate sum of balances
grouped_suppliers = supplier_df.groupby("Supplier", as_index=False).agg({
    "Clsng Blns Debit": "sum",
    "Clsng Blns Credit": "sum"
})

# Calculate Net_value_amount = Clsng Blns Debit + Clsng Blns Credit
grouped_suppliers["Net_value_amount"] = grouped_suppliers["Clsng Blns Debit"].fillna(0) + grouped_suppliers["Clsng Blns Credit"].fillna(0)

# Get list of suppliers with Net_value_amount > 0
suppliers_with_balance = grouped_suppliers[grouped_suppliers["Net_value_amount"] > 0]["Supplier"].astype(str).str.strip().unique()

# Define filter criteria for invoice data
exclude_gl_texts = [
    "Intercompany payable", "IOU manager", "IOU staff", "Short Term loan",
    "Trade creditors-Foreign", "Vendors bills of exchange", "Transport Creditors"
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
    (df["Due/Not"].astype(str).str.strip().str.lower() == "due") &
    df["Bank account"].notna() &
    (df["Bank account"].astype(str).str.strip() != "")
]

# Exclude suppliers with total Net_value_amount > 0
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

# List of Supplier codes to remove from final output
# excluded_suppliers = [
#     6004266, 7000241, 7000339, 7000353, 7000354, 7000360,
#     7000363, 7000364, 7000365, 7000366, 7000367, 7000368,
#     7000369, 7000374, 7000382, 7000385, 7000386, 7000388,
#     7000391, 7000398, 7000400, 7000421, 7000423, 7000425,
#     7000430, 7000441, 7000442, 7000443, 7000446
# ]

# Exclude them from the final grouped summary
#grouped_df = grouped_df[~grouped_df["Supplier"].astype(str).isin([str(code) for code in excluded_suppliers])]

# Save the grouped summary
summary_output_file = "grouped_summary.xlsx"
grouped_df.to_excel(summary_output_file, index=False)
print(f"Grouped summary saved to: {summary_output_file}")
