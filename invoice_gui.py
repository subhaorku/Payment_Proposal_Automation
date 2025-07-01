import pandas as pd
from tkinter import Tk, Label, Button, filedialog, messagebox
import os

def process_file(file_path):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path, header=1)
        df.columns = df.columns.str.strip()

        # Filter criteria
        exclude_gl_texts = [
            "Intercompany payable", "IOU manager", "IOU staff", "Short Term loan",
            "Trade creditors-Foreign", "Vendors bills of exchange"
        ]
        valid_payment_method = "T"
        valid_currency = "NGN"

        filtered_df = df[
            ~df["G/L Account: Long Text"].isin(exclude_gl_texts) &
            df["Payment block"].isna() &
            (df["Payment Method"] == valid_payment_method) &
            (df["Currency"] == valid_currency) &
            ~df["Diageo"].astype(str).str.contains("NTC- VENDOR", case=False, na=False) &
            df["Net Due Date"].notna() &
            (df["Due/Not"].astype(str).str.strip().str.lower() == "due")
        ]

        # Save filtered file
        filtered_file = "filtered_invoices.xlsx"
        filtered_df.to_excel(filtered_file, index=False)

        # Group and aggregate
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

        # Save grouped summary
        summary_file = "grouped_summary.xlsx"
        grouped_df.to_excel(summary_file, index=False)

        messagebox.showinfo("Success", f"Files saved:\n{filtered_file}\n{summary_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        process_file(file_path)

# GUI setup
root = Tk()
root.title("Invoice Filter and Summary Tool")
root.geometry("400x200")

Label(root, text="Upload your workings_file.xlsx", font=("Arial", 12)).pack(pady=20)
Button(root, text="Choose File", command=upload_file, font=("Arial", 12)).pack()
Button(root, text="Exit", command=root.quit, font=("Arial", 12)).pack(pady=10)

root.mainloop()
# https://convertico.com/
