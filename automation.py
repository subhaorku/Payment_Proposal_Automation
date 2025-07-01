import pandas as pd
from datetime import datetime

# Load the Excel file
input_file = "workings_file.xlsx"  # Make sure this file is in the same directory
df = pd.read_excel(input_file, header=1)

# Clean column names
df.columns = df.columns.str.strip()

# Define current date for filtering invoices due
current_date = pd.to_datetime("today").normalize()

# Define filter criteria
exclude_gl_texts = [
    "Intercompany payable", "IOU manager", "IOU staff", "Short-Term loan",
    "Trade creditors- Foreign", "Vendor bills of exchange (SF- Payments)"
]
valid_payment_blocks = ["A", "B"]
valid_payment_method = "T"
valid_currency = "NGN"

# Apply all filters
filtered_df = df[
    ~df["G/L Account: Long Text"].isin(exclude_gl_texts) &
    df["Payment block"].isna() &
    (df["Payment Method"] == valid_payment_method) &
    (df["Currency"] == valid_currency) &
    ~df["Diageo"].astype(str).str.contains("NTC- VENDOR", case=False, na=False) &
    df["Net Due Date"].notna()

    # (pd.to_datetime(df["Net Due Date"], errors="coerce") <= current_date)
]

# Save to a new Excel file
output_file = "filtered_invoices.xlsx"
filtered_df.to_excel(output_file, index=False)

print(f"Filtered data saved to: {output_file}")


