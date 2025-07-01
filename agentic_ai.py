import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import yaml
from datetime import datetime
import os

class DynamicInvoiceProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Daily Invoice Processor")
        self.root.geometry("1000x800")  
        
        # Initialize config
        self.config = self.load_default_config()
        
        # GUI Setup
        self.setup_ui()
        
    def load_default_config(self):
        """Load or create default configuration"""
        default_config = {
            "filters": {
                "exclude_gl_texts": [
                    "Intercompany payable", 
                    "IOU manager", 
                    "IOU staff",
                    "Short Term loan",
                    "Trade creditors-Foreign",
                    "Vendors bills of exchange", 
                    "Transport Creditors"
                ],
                "payment_method": "T",
                "currency": "NGN",
                "exclude_suppliers_with_balance": True,
                "exclude_payment_block": True,
                "exclude_ntc_vendor": True,
                "exclude_blank_suppliers": True,
                "exclude_blank_bank_accounts": True,
                "additional_exclusions": []
            },
            "grouping": {
                "by": ["Supplier"],
                "aggregations": {
                    "Name": "first",
                    "WHT availability": "first",
                    "Diageo/Tolaram": "first",
                    "Document Currency Value": "sum",
                    "Payable after WHT": "sum"
                }
            },
            "output": {
                "output_folder": "processed_results",
                "file_prefix": datetime.now().strftime("%Y%m%d")
            }
        }
        
        # Create config file if doesn't exist
        if not os.path.exists("config.yaml"):
            with open("config.yaml", "w") as f:
                yaml.dump(default_config, f)
            return default_config
        else:
            with open("config.yaml", "r") as f:
                return yaml.safe_load(f)
    
    def setup_ui(self):
        """Setup the main user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File Selection Section
        file_frame = ttk.LabelFrame(main_frame, text="1. Upload Daily Files", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        # Invoice File
        ttk.Label(file_frame, text="Invoice Excel File:").grid(row=0, column=0, sticky="e", padx=5)
        self.invoice_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.invoice_path, width=60).grid(row=0, column=1)
        ttk.Button(file_frame, text="Browse", command=lambda: self.browse_file(self.invoice_path)).grid(row=0, column=2)
        
        # Supplier File
        ttk.Label(file_frame, text="Supplier Balance File:").grid(row=1, column=0, sticky="e", padx=5)
        self.supplier_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.supplier_path, width=60).grid(row=1, column=1)
        ttk.Button(file_frame, text="Browse", command=lambda: self.browse_file(self.supplier_path)).grid(row=1, column=2)
        
        # Configuration Section
        config_frame = ttk.LabelFrame(main_frame, text="2. Processing Rules", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        
        # Filter Rules - Row 0
        ttk.Label(config_frame, text="Texts to Exclude (comma separated):").grid(row=0, column=0, sticky="e")
        self.exclude_texts = tk.StringVar(value=", ".join(self.config["filters"]["exclude_gl_texts"]))
        ttk.Entry(config_frame, textvariable=self.exclude_texts, width=60).grid(row=0, column=1)
        
        # Payment Method - Row 1
        ttk.Label(config_frame, text="Payment Method:").grid(row=1, column=0, sticky="e")
        self.payment_method = tk.StringVar(value=self.config["filters"]["payment_method"])
        ttk.Entry(config_frame, textvariable=self.payment_method, width=10).grid(row=1, column=1, sticky="w")
        
        # Currency - Row 2
        ttk.Label(config_frame, text="Currency:").grid(row=2, column=0, sticky="e")
        self.currency = tk.StringVar(value=self.config["filters"]["currency"])
        ttk.Entry(config_frame, textvariable=self.currency, width=10).grid(row=2, column=1, sticky="w")
        
        # Checkboxes for additional filters - Row 3
        self.exclude_balance_var = tk.BooleanVar(value=self.config["filters"]["exclude_suppliers_with_balance"])
        ttk.Checkbutton(config_frame, text="Exclude suppliers with balance", variable=self.exclude_balance_var).grid(row=3, column=0, sticky="w")
        
        self.exclude_payment_block_var = tk.BooleanVar(value=self.config["filters"]["exclude_payment_block"])
        ttk.Checkbutton(config_frame, text="Exclude payment blocked items", variable=self.exclude_payment_block_var).grid(row=3, column=1, sticky="w")
        
        # Checkboxes for additional filters - Row 4
        self.exclude_ntc_var = tk.BooleanVar(value=self.config["filters"]["exclude_ntc_vendor"])
        ttk.Checkbutton(config_frame, text="Exclude NTC-VENDOR items", variable=self.exclude_ntc_var).grid(row=4, column=0, sticky="w")
        
        self.exclude_blank_suppliers_var = tk.BooleanVar(value=self.config["filters"]["exclude_blank_suppliers"])
        ttk.Checkbutton(config_frame, text="Exclude blank suppliers", variable=self.exclude_blank_suppliers_var).grid(row=4, column=1, sticky="w")
        
        self.exclude_blank_bank_var = tk.BooleanVar(value=self.config["filters"]["exclude_blank_bank_accounts"])
        ttk.Checkbutton(config_frame, text="Exclude blank bank accounts", variable=self.exclude_blank_bank_var).grid(row=4, column=2, sticky="w")
        
        # Output Settings
        output_frame = ttk.LabelFrame(main_frame, text="3. Output Settings", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="Output Folder:").grid(row=0, column=0, sticky="e")
        self.output_folder = tk.StringVar(value=self.config["output"]["output_folder"])
        ttk.Entry(output_frame, textvariable=self.output_folder, width=60).grid(row=0, column=1)
        ttk.Button(output_frame, text="Browse", command=self.browse_output_folder).grid(row=0, column=2)
        
        ttk.Label(output_frame, text="File Prefix:").grid(row=1, column=0, sticky="e")
        self.file_prefix = tk.StringVar(value=self.config["output"]["file_prefix"])
        ttk.Entry(output_frame, textvariable=self.file_prefix, width=30).grid(row=1, column=1, sticky="w")
        
        # Process Button
        process_btn = ttk.Button(main_frame, text="Process Files", command=self.process_files)
        process_btn.pack(pady=10)
        
        # Logging Area
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        # Status Bar
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN).pack(fill=tk.X)
    
    def format_negative_numbers(val):
        """Format negative numbers with brackets"""
        if isinstance(val, (int, float)):
            return f"({abs(val):,.2f})" if val < 0 else f"{val:,.2f}"
        return val    
    def browse_file(self, path_var):
        """Open file dialog and set path"""
        file_path = filedialog.askopenfilename(
            title="Select file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            path_var.set(file_path)
    
    def browse_output_folder(self):
        """Open folder dialog for output location"""
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        if folder_path:
            self.output_folder.set(folder_path)
    
    def log_message(self, message):
        """Add message to log area"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_area.see(tk.END)
        self.root.update()
    
    def update_status(self, message):
        """Update status bar"""
        self.status_var.set(message)
        self.root.update()
    
    def validate_inputs(self):
        """Check if required inputs are provided"""
        if not self.invoice_path.get():
            messagebox.showerror("Error", "Please select an invoice file")
            return False
        if not self.supplier_path.get():
            messagebox.showerror("Error", "Please select a supplier file")
            return False
        if not os.path.exists(self.invoice_path.get()):
            messagebox.showerror("Error", "Invoice file does not exist")
            return False
        if not os.path.exists(self.supplier_path.get()):
            messagebox.showerror("Error", "Supplier file does not exist")
            return False
        return True
    
    def update_config(self):
        """Update configuration from UI inputs"""
        self.config["filters"]["exclude_gl_texts"] = [
            x.strip() for x in self.exclude_texts.get().split(",") if x.strip()
        ]
        self.config["filters"]["payment_method"] = self.payment_method.get()
        self.config["filters"]["currency"] = self.currency.get()
        self.config["filters"]["exclude_suppliers_with_balance"] = self.exclude_balance_var.get()
        self.config["filters"]["exclude_payment_block"] = self.exclude_payment_block_var.get()
        self.config["filters"]["exclude_ntc_vendor"] = self.exclude_ntc_var.get()
        self.config["filters"]["exclude_blank_suppliers"] = self.exclude_blank_suppliers_var.get()
        self.config["filters"]["exclude_blank_bank_accounts"] = self.exclude_blank_bank_var.get()
        self.config["output"]["output_folder"] = self.output_folder.get()
        self.config["output"]["file_prefix"] = self.file_prefix.get()
        
        # Save config for future use
        with open("config.yaml", "w") as f:
            yaml.dump(self.config, f)
    

    def process_files(self):
        """Main processing function"""
        if not self.validate_inputs():
            return
        
        try:
            self.update_status("Processing...")
            self.log_message("Starting invoice processing")
            self.update_config()
            
            # Create output folder if not exists
            os.makedirs(self.config["output"]["output_folder"], exist_ok=True)
            
            # Load data
            self.log_message("Loading invoice data...")
            invoice_df = pd.read_excel(self.invoice_path.get(), header=1)
            invoice_df.columns = invoice_df.columns.str.strip()
            
            self.log_message("Loading supplier data...")
            supplier_df = pd.read_excel(self.supplier_path.get())
            supplier_df.columns = supplier_df.columns.str.strip()
            
            # Apply filters
            self.log_message("Applying filters...")
            filtered_df = self.apply_filters(invoice_df, supplier_df)
            
            # Group data
            self.log_message("Grouping data...")
            grouped_df = self.apply_grouping(filtered_df)
            
            # Save outputs
            prefix = self.config["output"]["file_prefix"]
            output_folder = self.config["output"]["output_folder"]
            
            filtered_path = os.path.join(output_folder, f"{prefix}_filtered.xlsx")
            summary_path = os.path.join(output_folder, f"{prefix}_summary.xlsx")

            # Function to save with proper number formatting
            def save_with_accounting_format(df, file_path):
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    
                    # Create accounting number format
                    accounting_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                    
                    # Apply to all numeric columns
                    for col in df.select_dtypes(include=['int64', 'float64']).columns:
                        col_idx = df.columns.get_loc(col) + 1  # +1 because Excel is 1-based
                        
                        # Apply to all cells in column
                        for row in range(2, len(df)+2):  # +2 for header and 1-based index
                            cell = worksheet.cell(row=row, column=col_idx)
                            if pd.api.types.is_numeric_dtype(df[col]):
                                cell.number_format = accounting_format
            
            self.log_message(f"Saving filtered data to: {filtered_path}")
            save_with_accounting_format(filtered_df, filtered_path)
            
            self.log_message(f"Saving summary data to: {summary_path}")
            save_with_accounting_format(grouped_df, summary_path)
            
            self.log_message("Processing completed successfully!")
            self.update_status("Ready")
            messagebox.showinfo("Success", "Files processed successfully!")
            
            # Open output folder
            os.startfile(output_folder)
            
        except Exception as e:
            self.log_message(f"ERROR: {str(e)}")
            self.update_status("Error occurred")
            messagebox.showerror("Processing Error", str(e))

            
            # Custom Excel writer function with formatting
            def save_with_formatting(df, path):
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                    
                    # Get the workbook and worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    
                    # Define accounting format (negative numbers in brackets)
                    accounting_format = '#,##0.00_);(#,##0.00)'
                    
                    # Apply formatting to all numeric columns
                    for col_num, col_name in enumerate(df.columns, 1):
                        if pd.api.types.is_numeric_dtype(df[col_name]):
                            for row_num in range(2, len(df) + 2):  # +2 because Excel is 1-based and we have header
                                worksheet.cell(row=row_num, column=col_num).number_format = accounting_format
            
            self.log_message(f"Saving filtered data to: {filtered_path}")
            save_with_formatting(filtered_df, filtered_path)
            
            self.log_message(f"Saving summary data to: {summary_path}")
            save_with_formatting(grouped_df, summary_path)
            
            self.log_message("Processing completed successfully!")
            self.update_status("Ready")
            messagebox.showinfo("Success", "Files processed successfully!")
            
            # Open output folder
            os.startfile(output_folder)
            
        except Exception as e:
            self.log_message(f"ERROR: {str(e)}")
            self.update_status("Error occurred")
            messagebox.showerror("Processing Error", str(e))    
    def apply_filters(self, invoice_df, supplier_df):
        """Apply all configured filters"""
        filters = self.config["filters"]
        text_columns = ["G/L Account: Long Text", "Payment Method", "Currency", 
                    "Payment block", "Diageo", "Supplier", "Bank account", "Due/Not"]
        for col in text_columns:
            if col in invoice_df.columns:
                invoice_df[col] = invoice_df[col].astype(str).str.strip()
        # Convert columns to string and strip whitespace for consistent comparison
        #invoice_df = invoice_df.apply(lambda x: x.astype(str)).apply(lambda x: x.str.strip())
        
        # Exclude GL texts
        if "exclude_gl_texts" in filters and filters["exclude_gl_texts"]:
            self.log_message(f"Excluding GL texts: {', '.join(filters['exclude_gl_texts'])}")
            invoice_df = invoice_df[~invoice_df["G/L Account: Long Text"].isin(filters["exclude_gl_texts"])]
        
        # Payment method filter
        if "payment_method" in filters:
            self.log_message(f"Filtering for payment method: {filters['payment_method']}")
            invoice_df = invoice_df[invoice_df["Payment Method"] == filters["payment_method"]]
        
        # Currency filter
        if "currency" in filters:
            self.log_message(f"Filtering for currency: {filters['currency']}")
            invoice_df = invoice_df[invoice_df["Currency"] == filters["currency"]]
        
        # Exclude payment blocked items
        if filters.get("exclude_payment_block", False):
            self.log_message("Excluding payment blocked items (A, B, R, V)")
            invoice_df = invoice_df[~invoice_df["Payment block"].isin(["A", "B", "R", "V"])]
        
        # Exclude NTC-VENDOR items
        if filters.get("exclude_ntc_vendor", False):
            self.log_message("Excluding NTC-VENDOR items")
            invoice_df = invoice_df[~invoice_df["Diageo"].str.contains("NTC- VENDOR", case=False, na=False)]
        
        # Exclude blank suppliers
        if filters.get("exclude_blank_suppliers", False):
            self.log_message("Excluding blank suppliers")
            invoice_df = invoice_df[invoice_df["Supplier"].notna() & (invoice_df["Supplier"] != "")]
        
        # Exclude blank bank accounts
        if filters.get("exclude_blank_bank_accounts", False):
            self.log_message("Excluding blank bank accounts")
            invoice_df = invoice_df[~invoice_df["Bank account"].isin(["", "nan", "None"]) & invoice_df["Bank account"].notna()]
        
        # Additional validation checks
        self.log_message("Applying additional validations...")
        invoice_df = invoice_df[
            (invoice_df["Net Due Date"].notna()) &
            (invoice_df["Due/Not"].str.strip().str.lower() == "due")
        ]
        
        # Exclude suppliers with balance
        if filters.get("exclude_suppliers_with_balance", False):
            self.log_message("Excluding suppliers with outstanding balances")
            suppliers_with_balance = self.get_suppliers_with_balance(supplier_df)
            invoice_df = invoice_df[~invoice_df["Supplier"].isin(suppliers_with_balance)]
        
        return invoice_df



    def apply_grouping(self, df):
        """Apply grouping and aggregation"""
        grouping = self.config["grouping"]
        self.log_message(f"Grouping by: {', '.join(grouping['by'])}")
        return df.groupby(grouping["by"], as_index=False).agg(grouping["aggregations"])
    



    def get_suppliers_with_balance(self, supplier_df):
        """Identify suppliers where sum of (Debit + Credit) > 0"""
        supplier_df = supplier_df[supplier_df["Supplier"].notna()].copy()
        
        # Convert to numeric (if not already)
        supplier_df["Clsng Blns Debit"] = pd.to_numeric(supplier_df["Clsng Blns Debit"], errors="coerce").fillna(0)
        supplier_df["Clsng Blns Credit"] = pd.to_numeric(supplier_df["Clsng Blns Credit"], errors="coerce").fillna(0)
        
        # Calculate Net_value (Debit + Credit)
        supplier_df["Net_value"] = supplier_df["Clsng Blns Debit"] + supplier_df["Clsng Blns Credit"]
        
        # Group by Supplier and sum Net_value
        grouped = supplier_df.groupby("Supplier", as_index=False)["Net_value"].sum()
        
        # Filter suppliers where sum(Net_value) > 0
        suppliers_with_balance = grouped[grouped["Net_value"] > 0]["Supplier"].astype(str).str.strip().unique()
        return suppliers_with_balance       
    

if __name__ == "__main__":
    root = tk.Tk()
    app = DynamicInvoiceProcessor(root)
    root.mainloop()