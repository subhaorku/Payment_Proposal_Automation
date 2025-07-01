import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import yaml
from datetime import datetime
import os

class DynamicInvoiceProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Daily Invoice Processor")
        self.root.geometry("1000x1000")
        ctk.set_appearance_mode("system")         # “dark”, “light”, or “system”
        ctk.set_default_color_theme("blue")       # “blue”, “green”, etc.

        # Load or create config.yaml
        self.config = self.load_default_config()

        # Build the UI
        self.setup_ui()

    def load_default_config(self):
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

        if not os.path.exists("config.yaml"):
            with open("config.yaml", "w") as f:
                yaml.dump(default_config, f)
            return default_config
        else:
            with open("config.yaml", "r") as f:
                return yaml.safe_load(f)

    def setup_ui(self):
        # ------------------------------
        # MAIN CONTAINER
        # ------------------------------
        main_frame = ctk.CTkFrame(self.root, corner_radius=10)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header = ctk.CTkLabel(
            main_frame,
            text="📑 Daily Invoice Processor",
            font=ctk.CTkFont(size=20, weight="bold"),
            pady=10
        )
        header.pack(anchor="n")

        # ------------------------------
        # 1) FILE SELECTION SECTION
        # ------------------------------
        section1_label = ctk.CTkLabel(
            main_frame,
            text="1. Upload Daily Files",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w"
        )
        section1_label.pack(fill="x", pady=(10, 0), padx=10)

        file_frame = ctk.CTkFrame(main_frame, corner_radius=8)
        file_frame.pack(fill="x", pady=8, padx=10)

        # Configure grid weights so entries expand
        file_frame.columnconfigure(1, weight=1)

        # Invoice Excel File
        lbl_invoice = ctk.CTkLabel(file_frame, text="Invoice Excel File:")
        lbl_invoice.grid(row=0, column=0, sticky="e", padx=(10, 5), pady=8)

        self.invoice_path = ctk.StringVar()
        entry_invoice = ctk.CTkEntry(file_frame, textvariable=self.invoice_path)
        entry_invoice.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=8)

        btn_browse_invoice = ctk.CTkButton(
            file_frame,
            text="Browse",
            width=80,
            command=lambda: self.browse_file(self.invoice_path)
        )
        btn_browse_invoice.grid(row=0, column=2, padx=(5, 10), pady=8)

        # Supplier Balance File
        lbl_supplier = ctk.CTkLabel(file_frame, text="Supplier Balance File:")
        lbl_supplier.grid(row=1, column=0, sticky="e", padx=(10, 5), pady=8)

        self.supplier_path = ctk.StringVar()
        entry_supplier = ctk.CTkEntry(file_frame, textvariable=self.supplier_path)
        entry_supplier.grid(row=1, column=1, sticky="ew", padx=(0, 5), pady=8)

        btn_browse_supplier = ctk.CTkButton(
            file_frame,
            text="Browse",
            width=80,
            command=lambda: self.browse_file(self.supplier_path)
        )
        btn_browse_supplier.grid(row=1, column=2, padx=(5, 10), pady=8)

        # ------------------------------
        # 2) PROCESSING RULES SECTION
        # ------------------------------
        section2_label = ctk.CTkLabel(
            main_frame,
            text="2. Processing Rules",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w"
        )
        section2_label.pack(fill="x", pady=(15, 0), padx=10)

        config_frame = ctk.CTkFrame(main_frame, corner_radius=8)
        config_frame.pack(fill="x", pady=8, padx=10)

        # Configure columns so entry widgets can expand
        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(2, weight=1)

        # a) Texts to Exclude (Comma Separated)
        lbl_exclude_texts = ctk.CTkLabel(config_frame, text="Texts to Exclude (comma separated):")
        lbl_exclude_texts.grid(row=0, column=0, sticky="e", padx=(10, 5), pady=8)

        self.exclude_texts = ctk.StringVar(value=", ".join(self.config["filters"]["exclude_gl_texts"]))
        entry_exclude_texts = ctk.CTkEntry(config_frame, textvariable=self.exclude_texts)
        entry_exclude_texts.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(0, 10), pady=8)

        # b) Payment Method
        lbl_payment_method = ctk.CTkLabel(config_frame, text="Payment Method:")
        lbl_payment_method.grid(row=1, column=0, sticky="e", padx=(10, 5), pady=8)

        self.payment_method = ctk.StringVar(value=self.config["filters"]["payment_method"])
        entry_payment_method = ctk.CTkEntry(config_frame, textvariable=self.payment_method, width=80)
        entry_payment_method.grid(row=1, column=1, sticky="w", padx=(0, 10), pady=8)

        # c) Currency
        lbl_currency = ctk.CTkLabel(config_frame, text="Currency:")
        lbl_currency.grid(row=2, column=0, sticky="e", padx=(10, 5), pady=8)

        self.currency = ctk.StringVar(value=self.config["filters"]["currency"])
        entry_currency = ctk.CTkEntry(config_frame, textvariable=self.currency, width=80)
        entry_currency.grid(row=2, column=1, sticky="w", padx=(0, 10), pady=8)

        # d) Checkboxes (grouped into two rows)
        # Row 3
        self.exclude_balance_var = ctk.BooleanVar(value=self.config["filters"]["exclude_suppliers_with_balance"])
        chk_excl_balance = ctk.CTkCheckBox(
            config_frame,
            text="Exclude suppliers with balance",
            variable=self.exclude_balance_var,
            onvalue=True,
            offvalue=False
        )
        chk_excl_balance.grid(row=3, column=0, columnspan=2, sticky="w", padx=(10, 5), pady=5)

        self.exclude_payment_block_var = ctk.BooleanVar(value=self.config["filters"]["exclude_payment_block"])
        chk_excl_payment_block = ctk.CTkCheckBox(
            config_frame,
            text="Exclude payment blocked items",
            variable=self.exclude_payment_block_var,
            onvalue=True,
            offvalue=False
        )
        chk_excl_payment_block.grid(row=3, column=2, columnspan=1, sticky="w", padx=(5, 10), pady=5)

        # Row 4
        self.exclude_ntc_var = ctk.BooleanVar(value=self.config["filters"]["exclude_ntc_vendor"])
        chk_excl_ntc = ctk.CTkCheckBox(
            config_frame,
            text="Exclude NTC-VENDOR items",
            variable=self.exclude_ntc_var,
            onvalue=True,
            offvalue=False
        )
        chk_excl_ntc.grid(row=4, column=0, columnspan=1, sticky="w", padx=(10, 5), pady=5)

        self.exclude_blank_suppliers_var = ctk.BooleanVar(value=self.config["filters"]["exclude_blank_suppliers"])
        chk_excl_blank_suppliers = ctk.CTkCheckBox(
            config_frame,
            text="Exclude blank suppliers",
            variable=self.exclude_blank_suppliers_var,
            onvalue=True,
            offvalue=False
        )
        chk_excl_blank_suppliers.grid(row=4, column=1, columnspan=1, sticky="w", padx=(5, 5), pady=5)

        self.exclude_blank_bank_var = ctk.BooleanVar(value=self.config["filters"]["exclude_blank_bank_accounts"])
        chk_excl_blank_bank = ctk.CTkCheckBox(
            config_frame,
            text="Exclude blank bank accounts",
            variable=self.exclude_blank_bank_var,
            onvalue=True,
            offvalue=False
        )
        chk_excl_blank_bank.grid(row=4, column=2, columnspan=1, sticky="w", padx=(5, 10), pady=5)

        # ------------------------------
        # 3) OUTPUT SETTINGS SECTION
        # ------------------------------
        section3_label = ctk.CTkLabel(
            main_frame,
            text="3. Output Settings",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w"
        )
        section3_label.pack(fill="x", pady=(15, 0), padx=10)

        output_frame = ctk.CTkFrame(main_frame, corner_radius=8)
        output_frame.pack(fill="x", pady=8, padx=10)

        output_frame.columnconfigure(1, weight=1)

        # a) Output Folder
        lbl_output_folder = ctk.CTkLabel(output_frame, text="Output Folder:")
        lbl_output_folder.grid(row=0, column=0, sticky="e", padx=(10, 5), pady=8)

        self.output_folder = ctk.StringVar(value=self.config["output"]["output_folder"])
        entry_output_folder = ctk.CTkEntry(output_frame, textvariable=self.output_folder)
        entry_output_folder.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=8)

        btn_browse_output = ctk.CTkButton(
            output_frame,
            text="Browse",
            width=80,
            command=self.browse_output_folder
        )
        btn_browse_output.grid(row=0, column=2, padx=(5, 10), pady=8)

        # b) File Prefix
        lbl_file_prefix = ctk.CTkLabel(output_frame, text="File Prefix:")
        lbl_file_prefix.grid(row=1, column=0, sticky="e", padx=(10, 5), pady=8)

        self.file_prefix = ctk.StringVar(value=self.config["output"]["file_prefix"])
        entry_file_prefix = ctk.CTkEntry(output_frame, textvariable=self.file_prefix, width=120)
        entry_file_prefix.grid(row=1, column=1, sticky="w", padx=(0, 5), pady=8)

        # ------------------------------
        # PROCESS BUTTON
        # ------------------------------
        btn_frame = ctk.CTkFrame(main_frame, corner_radius=5)
        btn_frame.pack(fill="x", pady=(15, 0), padx=10)

        process_btn = ctk.CTkButton(
            btn_frame,
            text="Process Files",
            fg_color="#26ba4b",
            hover_color="#1e9b3f",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.process_files,
            height=40
        )
        process_btn.pack(pady=10, padx=10)

        # ------------------------------
        # 4) LOGGING AREA
        # ------------------------------
        section4_label = ctk.CTkLabel(
            main_frame,
            text="4. Processing Log",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w"
        )
        section4_label.pack(fill="x", pady=(15, 0), padx=10)

        log_frame = ctk.CTkFrame(main_frame, corner_radius=8)
        log_frame.pack(fill="both", expand=True, pady=8, padx=10)

        # Use CTkTextbox for a built-in dark-mode text area
        self.log_area = ctk.CTkTextbox(
            log_frame,
            width=0,        # let “pack(fill='both')” handle widths
            height=0,       # let “pack(fill='both')” handle heights
            fg_color="#1e1e1e",
            text_color="#ffffff",
            corner_radius=5,
            font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.log_area.pack(fill="both", expand=True, padx=5, pady=5)

        # ------------------------------
        # STATUS BAR
        # ------------------------------
        self.status_var = ctk.StringVar(value="Ready")
        status_bar = ctk.CTkLabel(
            main_frame,
            textvariable=self.status_var,
            fg_color="#292c31",
            text_color="#26ba4b",
            anchor="w",
            corner_radius=5,
            height=30
        )
        status_bar.pack(fill="x", pady=(0, 5), padx=10)

    # -------------------------------------------------
    # Helper methods (browse dialogs, logging, status)
    # -------------------------------------------------
    def browse_file(self, path_var):
        file_path = filedialog.askopenfilename(
            title="Select a file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            path_var.set(file_path)

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        if folder_path:
            self.output_folder.set(folder_path)

    def log_message(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert("end", f"[{timestamp}] {message}\n")
        self.log_area.see("end")
        self.root.update()

    def update_status(self, message: str):
        self.status_var.set(message)
        self.root.update()

    # -------------------------------------------------
    # Validation & Config update
    # -------------------------------------------------
    def validate_inputs(self) -> bool:
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
        with open("config.yaml", "w") as f:
            yaml.dump(self.config, f)

    # -------------------------------------------------
    # Main Processing Logic
    # -------------------------------------------------
    def process_files(self):
        if not self.validate_inputs():
            return

        try:
            self.update_status("Processing...")
            self.log_message("🔄 Starting invoice processing...")
            self.update_config()
            os.makedirs(self.config["output"]["output_folder"], exist_ok=True)

            # Load invoice data
            self.log_message("📥 Loading invoice data...")
            invoice_df = pd.read_excel(self.invoice_path.get(), header=1)
            invoice_df.columns = invoice_df.columns.str.strip()

            # Load supplier data
            self.log_message("📥 Loading supplier data...")
            supplier_df = pd.read_excel(self.supplier_path.get())
            supplier_df.columns = supplier_df.columns.str.strip()

            # Apply filters
            self.log_message("🔀 Applying filters...")
            filtered_df = self.apply_filters(invoice_df, supplier_df)

            # Group/aggregate
            self.log_message("📊 Grouping data...")
            grouped_df = self.apply_grouping(filtered_df)

            # Prepare output paths
            prefix = self.config["output"]["file_prefix"]
            output_folder = self.config["output"]["output_folder"]
            filtered_path = os.path.join(output_folder, f"{prefix}_filtered.xlsx")
            summary_path = os.path.join(output_folder, f"{prefix}_summary.xlsx")

            # Helper to save with accounting format
            def save_with_accounting_format(df: pd.DataFrame, file_path: str):
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Sheet1")
                    workbook = writer.book
                    worksheet = writer.sheets["Sheet1"]
                    accounting_fmt = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
                    for col_name in df.select_dtypes(include=["int64", "float64"]).columns:
                        col_idx = df.columns.get_loc(col_name) + 1
                        for row in range(2, len(df) + 2):
                            cell = worksheet.cell(row=row, column=col_idx)
                            if pd.api.types.is_numeric_dtype(df[col_name]):
                                cell.number_format = accounting_fmt

            # Save filtered data
            self.log_message(f"💾 Saving filtered data → {filtered_path}")
            save_with_accounting_format(filtered_df, filtered_path)

            # Save summary data
            self.log_message(f"💾 Saving summary data → {summary_path}")
            save_with_accounting_format(grouped_df, summary_path)

            self.log_message("✅ Processing completed successfully!")
            self.update_status("Ready")
            messagebox.showinfo("Success", "Files processed successfully!")

            # Attempt to open the output folder (Windows only; safely ignore errors elsewhere)
            try:
                os.startfile(output_folder)
            except Exception:
                pass

        except Exception as e:
            self.log_message(f"❌ ERROR: {str(e)}")
            self.update_status("Error occurred")
            messagebox.showerror("Processing Error", str(e))

    # -------------------------------------------------
    # Filtering & Grouping Functions
    # -------------------------------------------------
    def apply_filters(self, invoice_df: pd.DataFrame, supplier_df: pd.DataFrame) -> pd.DataFrame:
        filters = self.config["filters"]
        text_columns = [
            "G/L Account: Long Text", "Payment Method", "Currency",
            "Payment block", "Diageo", "Supplier", "Bank account", "Due/Not"
        ]
        for col in text_columns:
            if col in invoice_df.columns:
                invoice_df[col] = invoice_df[col].astype(str).str.strip()

        # 1) Exclude GL texts
        if filters["exclude_gl_texts"]:
            self.log_message(f"✂ Excluding GL texts: {', '.join(filters['exclude_gl_texts'])}")
            invoice_df = invoice_df[
                ~invoice_df["G/L Account: Long Text"].isin(filters["exclude_gl_texts"])
            ]

        # 2) Payment method filter
        if filters.get("payment_method"):
            self.log_message(f"🔍 Filtering for payment method: {filters['payment_method']}")
            invoice_df = invoice_df[invoice_df["Payment Method"] == filters["payment_method"]]

        # 3) Currency filter
        if filters.get("currency"):
            self.log_message(f"🔍 Filtering for currency: {filters['currency']}")
            invoice_df = invoice_df[invoice_df["Currency"] == filters["currency"]]

        # 4) Exclude payment blocked items (A, B, R, V)
        if filters.get("exclude_payment_block"):
            self.log_message("✂ Excluding payment blocked items (A, B, R, V)")
            invoice_df = invoice_df[~invoice_df["Payment block"].isin(["A", "B", "R", "V"])]

        # 5) Exclude NTC-VENDOR
        if filters.get("exclude_ntc_vendor"):
            self.log_message("✂ Excluding NTC-VENDOR items")
            invoice_df = invoice_df[
                ~invoice_df["Diageo"].str.contains("NTC- VENDOR", case=False, na=False)
            ]

        # 6) Exclude blank suppliers
        if filters.get("exclude_blank_suppliers"):
            self.log_message("✂ Excluding blank suppliers")
            invoice_df = invoice_df[
                invoice_df["Supplier"].notna() & (invoice_df["Supplier"] != "")
            ]

        # 7) Exclude blank bank accounts
        if filters.get("exclude_blank_bank_accounts"):
            self.log_message("✂ Excluding blank bank accounts")
            invoice_df = invoice_df[
                ~invoice_df["Bank account"].isin(["", "nan", "None"]) &
                invoice_df["Bank account"].notna()
            ]

        # 8) Additional validations: Net Due Date present & Due/Not == “due”
        self.log_message("🔬 Applying additional validations")
        invoice_df = invoice_df[
            (invoice_df["Net Due Date"].notna()) &
            (invoice_df["Due/Not"].str.strip().str.lower() == "due")
        ]

        # 9) Exclude suppliers with outstanding balance
        if filters.get("exclude_suppliers_with_balance"):
            self.log_message("✂ Excluding suppliers with outstanding balances")
            suppliers_with_balance = self.get_suppliers_with_balance(supplier_df)
            invoice_df = invoice_df[~invoice_df["Supplier"].isin(suppliers_with_balance)]

        return invoice_df

    def apply_grouping(self, df: pd.DataFrame) -> pd.DataFrame:
        grouping = self.config["grouping"]
        self.log_message(f"📑 Grouping by: {', '.join(grouping['by'])}")
        return df.groupby(grouping["by"], as_index=False).agg(grouping["aggregations"])

    def get_suppliers_with_balance(self, supplier_df: pd.DataFrame):
        supplier_df = supplier_df[supplier_df["Supplier"].notna()].copy()
        supplier_df["Clsng Blns Debit"] = pd.to_numeric(
            supplier_df["Clsng Blns Debit"], errors="coerce"
        ).fillna(0)
        supplier_df["Clsng Blns Credit"] = pd.to_numeric(
            supplier_df["Clsng Blns Credit"], errors="coerce"
        ).fillna(0)
        supplier_df["Net_value"] = (
            supplier_df["Clsng Blns Debit"] + supplier_df["Clsng Blns Credit"]
        )
        grouped = supplier_df.groupby("Supplier", as_index=False)["Net_value"].sum()
        suppliers_with_balance = grouped[grouped["Net_value"] > 0]["Supplier"]\
            .astype(str).str.strip().unique()
        return suppliers_with_balance


if __name__ == "__main__":
    root = ctk.CTk()
    app = DynamicInvoiceProcessor(root)
    root.mainloop()
