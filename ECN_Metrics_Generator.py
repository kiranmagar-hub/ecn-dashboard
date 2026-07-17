"""
BEF ECN Metrics Generator - GUI Application
Generates CT Metrics Dashboard and PowerPoint from Excel file

Usage:
  1. Run this application
  2. Select Excel file
  3. Choose output location
  4. Click Generate
  5. Reports created in date-stamped folder
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import json
import os
import shutil
from datetime import datetime
from pathlib import Path
import subprocess
import sys
try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False

class ECNMetricsGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("BEF ECN Metrics Generator")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        # Variables
        self.data_source = tk.StringVar(value="excel")  # "excel" or "database"
        self.excel_file = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))

        # Database connection variables
        self.db_server = tk.StringVar(value="wilmatom1f")
        self.db_database = tk.StringVar(value="WWECNRequests")
        self.db_username = tk.StringVar(value="ECNRequestData")
        self.db_password = tk.StringVar()
        self.db_table = tk.StringVar(value="Qry_EcnRequestGeneralInfoUpdate")

        # Default date range: start = 365 days ago, end = today
        from datetime import timedelta
        start_date = datetime.now() - timedelta(days=365)
        self.date_range_start = tk.StringVar(value=start_date.strftime("%Y-%m-%d"))
        self.date_range_end = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

        self.setup_ui()

    def setup_ui(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#667eea", height=100)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text="BEF ECN Metrics Generator",
            font=("Segoe UI", 24, "bold"),
            bg="#667eea",
            fg="white"
        )
        title_label.pack(pady=20)

        subtitle_label = tk.Label(
            header_frame,
            text="Generate CT Metrics Dashboard & PowerPoint from Excel or Database",
            font=("Segoe UI", 11),
            bg="#667eea",
            fg="white"
        )
        subtitle_label.pack()

        # Main content
        main_frame = tk.Frame(self.root, padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Data source selection
        self.source_frame = tk.LabelFrame(main_frame, text="1. Select Data Source", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        self.source_frame.pack(fill=tk.X, pady=10)

        tk.Radiobutton(self.source_frame, text="Excel File", variable=self.data_source, value="excel",
                      command=self.toggle_data_source, font=("Segoe UI", 10)).pack(anchor=tk.W)
        tk.Radiobutton(self.source_frame, text="SQL Server Database", variable=self.data_source, value="database",
                      command=self.toggle_data_source, font=("Segoe UI", 10)).pack(anchor=tk.W)

        # Excel file selection
        self.excel_frame = tk.LabelFrame(main_frame, text="Excel File", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        self.excel_frame.pack(fill=tk.X, pady=10)

        tk.Entry(self.excel_frame, textvariable=self.excel_file, font=("Segoe UI", 10), width=60).pack(side=tk.LEFT, padx=5)
        tk.Button(self.excel_frame, text="Browse...", command=self.browse_excel, bg="#667eea", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)

        # Database connection frame
        self.db_frame = tk.LabelFrame(main_frame, text="Database Connection", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        self.db_frame.pack(fill=tk.X, pady=10)

        db_grid = tk.Frame(self.db_frame)
        db_grid.pack(fill=tk.X)

        tk.Label(db_grid, text="Server:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky=tk.W, padx=5, pady=3)
        tk.Entry(db_grid, textvariable=self.db_server, font=("Segoe UI", 10), width=30).grid(row=0, column=1, padx=5, pady=3)

        tk.Label(db_grid, text="Database:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky=tk.W, padx=5, pady=3)
        tk.Entry(db_grid, textvariable=self.db_database, font=("Segoe UI", 10), width=30).grid(row=1, column=1, padx=5, pady=3)

        tk.Label(db_grid, text="Username:", font=("Segoe UI", 9)).grid(row=2, column=0, sticky=tk.W, padx=5, pady=3)
        tk.Entry(db_grid, textvariable=self.db_username, font=("Segoe UI", 10), width=30).grid(row=2, column=1, padx=5, pady=3)

        tk.Label(db_grid, text="Password:", font=("Segoe UI", 9)).grid(row=3, column=0, sticky=tk.W, padx=5, pady=3)
        tk.Entry(db_grid, textvariable=self.db_password, font=("Segoe UI", 10), width=30, show="*").grid(row=3, column=1, padx=5, pady=3)

        tk.Label(db_grid, text="Table/View:", font=("Segoe UI", 9)).grid(row=4, column=0, sticky=tk.W, padx=5, pady=3)
        tk.Entry(db_grid, textvariable=self.db_table, font=("Segoe UI", 10), width=30).grid(row=4, column=1, padx=5, pady=3)

        tk.Button(db_grid, text="Test Connection", command=self.test_db_connection, bg="#48bb78", fg="white",
                 font=("Segoe UI", 9, "bold")).grid(row=5, column=1, padx=5, pady=10, sticky=tk.W)

        # Output directory selection
        output_frame = tk.LabelFrame(main_frame, text="2. Output Location", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        output_frame.pack(fill=tk.X, pady=10)

        tk.Entry(output_frame, textvariable=self.output_dir, font=("Segoe UI", 10), width=60).pack(side=tk.LEFT, padx=5)
        tk.Button(output_frame, text="Browse...", command=self.browse_output, bg="#667eea", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)

        # Date range selection
        date_frame = tk.LabelFrame(main_frame, text="3. Date Range (Optional)", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        date_frame.pack(fill=tk.X, pady=10)

        date_inner = tk.Frame(date_frame)
        date_inner.pack()

        tk.Label(date_inner, text="Start Date (YYYY-MM-DD):", font=("Segoe UI", 9)).grid(row=0, column=0, padx=5, sticky=tk.W)
        tk.Entry(date_inner, textvariable=self.date_range_start, font=("Segoe UI", 10), width=15).grid(row=0, column=1, padx=5)

        tk.Label(date_inner, text="End Date (YYYY-MM-DD):", font=("Segoe UI", 9)).grid(row=0, column=2, padx=5, sticky=tk.W)
        tk.Entry(date_inner, textvariable=self.date_range_end, font=("Segoe UI", 10), width=15).grid(row=0, column=3, padx=5)

        tk.Button(date_inner, text="Clear Dates (Use All Data)", command=self.clear_dates, font=("Segoe UI", 8)).grid(row=0, column=4, padx=10)

        # Generate button
        button_frame = tk.Frame(main_frame, pady=20)
        button_frame.pack(fill=tk.X)

        self.generate_btn = tk.Button(
            button_frame,
            text="Generate Reports",
            command=self.generate_reports,
            bg="#48bb78",
            fg="white",
            font=("Segoe UI", 14, "bold"),
            height=2,
            cursor="hand2"
        )
        self.generate_btn.pack(fill=tk.X)

        # Progress log
        log_frame = tk.LabelFrame(main_frame, text="Progress Log", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, font=("Consolas", 9), bg="#f8f9fa")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("Segoe UI", 9))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Initialize UI state
        self.toggle_data_source()

    def toggle_data_source(self):
        """Show/hide Excel or Database frames based on selection"""
        if self.data_source.get() == "excel":
            self.excel_frame.pack(fill=tk.X, pady=10, after=self.source_frame)
            self.db_frame.pack_forget()
        else:
            self.db_frame.pack(fill=tk.X, pady=10, after=self.source_frame)
            self.excel_frame.pack_forget()

    def test_db_connection(self):
        """Test database connection"""
        if not PYODBC_AVAILABLE:
            messagebox.showerror("Missing Library",
                               "pyodbc is not installed.\n\n"
                               "Install it with:\n"
                               "pip install pyodbc")
            return

        try:
            self.log("Testing database connection...")
            conn_str = (
                f"DRIVER={{SQL Server}};"
                f"SERVER={self.db_server.get()};"
                f"DATABASE={self.db_database.get()};"
                f"UID={self.db_username.get()};"
                f"PWD={self.db_password.get()}"
            )

            conn = pyodbc.connect(conn_str, timeout=5)
            cursor = conn.cursor()

            # Test query
            cursor.execute(f"SELECT COUNT(*) FROM {self.db_table.get()}")
            count = cursor.fetchone()[0]

            conn.close()

            messagebox.showinfo("Success",
                              f"Connection successful!\n\n"
                              f"Table '{self.db_table.get()}' found.\n"
                              f"Record count: {count:,}")
            self.log(f"Database connection successful - {count:,} records found")

        except Exception as e:
            messagebox.showerror("Connection Failed",
                               f"Could not connect to database:\n\n{str(e)}")
            self.log(f"Database connection failed: {str(e)}")

    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file.set(filename)
            self.log(f"Selected Excel file: {filename}")

    def browse_output(self):
        dirname = filedialog.askdirectory(title="Select Output Location")
        if dirname:
            self.output_dir.set(dirname)
            self.log(f"Output location: {dirname}")

    def clear_dates(self):
        self.date_range_start.set("")
        self.date_range_end.set("")
        self.log("Date range cleared - will use all data")

    def fetch_data_from_database(self, output_excel_path):
        """Fetch ECN data from SQL Server database and save as Excel"""
        try:
            # Build connection string
            conn_str = (
                f"DRIVER={{SQL Server}};"
                f"SERVER={self.db_server.get()};"
                f"DATABASE={self.db_database.get()};"
                f"UID={self.db_username.get()};"
                f"PWD={self.db_password.get()}"
            )

            self.log(f"Connecting to {self.db_server.get()}\\{self.db_database.get()}...")

            # Connect to database
            conn = pyodbc.connect(conn_str, timeout=30)

            # Build query with date filter if specified
            start_date = self.date_range_start.get()
            end_date = self.date_range_end.get()

            if start_date and end_date:
                query = f"""
                SELECT *
                FROM {self.db_table.get()}
                WHERE SubmitDate >= '{start_date}' AND SubmitDate <= '{end_date}'
                """
                self.log(f"Applying date filter: {start_date} to {end_date}")
            else:
                query = f"SELECT * FROM {self.db_table.get()}"
                self.log("Fetching all records (no date filter)")

            # Fetch data
            self.log("Executing query...")
            df = pd.read_sql(query, conn)
            conn.close()

            self.log(f"Retrieved {len(df):,} records from database")

            # Save to Excel with the required sheet name
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Document_TB11', index=False)

            self.log(f"Saved data to Excel: {output_excel_path.name}")

        except Exception as e:
            raise Exception(f"Database fetch failed: {str(e)}")

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()

    def update_status(self, message):
        self.status_var.set(message)
        self.root.update()

    def generate_reports(self):
        # Validate inputs based on data source
        if self.data_source.get() == "excel":
            if not self.excel_file.get():
                messagebox.showerror("Error", "Please select an Excel file")
                return
            if not os.path.exists(self.excel_file.get()):
                messagebox.showerror("Error", "Excel file not found")
                return
        else:
            # Database validation
            if not PYODBC_AVAILABLE:
                messagebox.showerror("Error", "pyodbc is not installed.\n\nInstall with: pip install pyodbc")
                return
            if not all([self.db_server.get(), self.db_database.get(), self.db_username.get(),
                       self.db_password.get(), self.db_table.get()]):
                messagebox.showerror("Error", "Please fill in all database connection fields")
                return

        # Disable button during generation
        self.generate_btn.config(state=tk.DISABLED)
        self.log_text.delete(1.0, tk.END)

        try:
            self.log("="*60)
            self.log("BEF ECN METRICS GENERATION STARTED")
            self.log("="*60)

            # Create output folder with timestamp
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            output_folder = Path(self.output_dir.get()) / f"ECN_Metrics_{timestamp}"
            output_folder.mkdir(parents=True, exist_ok=True)

            self.log(f"Created output folder: {output_folder}")
            self.update_status("Processing data...")

            # Get data from source (Excel or Database)
            excel_dest = output_folder / "source_data.xlsx"

            if self.data_source.get() == "database":
                self.log("Fetching data from SQL Server database...")
                self.fetch_data_from_database(excel_dest)
            else:
                # Copy Excel file to output folder
                shutil.copy2(self.excel_file.get(), excel_dest)
                self.log(f"Copied Excel file to output folder")

            # Copy logo if exists
            logo_path = Path(__file__).parent / "ADI-Logo-RGB-FullColor.png"
            if logo_path.exists():
                shutil.copy2(logo_path, output_folder / "ADI-Logo-RGB-FullColor.png")
                self.log("Copied logo file")

            # Copy ECN Coordinators mapping file if exists
            coord_path = Path(__file__).parent / "ECN Coordinators.xlsx"
            if coord_path.exists():
                shutil.copy2(coord_path, output_folder / "ECN Coordinators.xlsx")
                self.log("Copied ECN Coordinators mapping file")
            else:
                self.log("Note: ECN Coordinators.xlsx not found - will use user IDs")

            # Copy ADI PowerPoint template if exists
            template_paths = [
                Path(__file__).parent / "ADI-Brand-PowerPoint-Template-Standard-2026-v1.potx",
                Path.home() / "Downloads" / "ADI-Brand-PowerPoint-Template-Standard-2026-v1.potx"
            ]
            for template_path in template_paths:
                if template_path.exists():
                    shutil.copy2(template_path, output_folder / "ADI-Brand-PowerPoint-Template-Standard-2026-v1.potx")
                    self.log("Copied ADI PowerPoint template")
                    break

            # Process data
            self.log("\nStep 1/4: Processing Excel data...")
            data = self.process_excel_data(excel_dest, output_folder)

            # Generate HTML dashboard
            self.log("\nStep 2/4: Generating HTML dashboard...")
            self.update_status("Generating HTML dashboard...")
            self.generate_html_dashboard(data, output_folder)

            # Generate PowerPoint
            self.log("\nStep 3/4: Generating PowerPoint presentation...")
            self.update_status("Generating PowerPoint...")
            self.generate_powerpoint(output_folder)

            # Create summary
            self.log("\nStep 4/4: Creating summary document...")
            self.create_summary(data, output_folder)

            self.log("\n" + "="*60)
            self.log("GENERATION COMPLETE!")
            self.log("="*60)
            self.log(f"\nOutput Location: {output_folder}")
            self.log("\nGenerated Files:")
            self.log("  - dashboard.html (Interactive web dashboard)")
            self.log("  - BEF_ECN_Metrics.pptx (PowerPoint presentation)")
            self.log("  - data.json (Processed metrics data)")
            self.log("  - SUMMARY.txt (Report summary)")
            self.log("  - source_data.xlsx (Copy of your Excel file)")

            self.update_status("Complete!")

            # Ask to open folder
            if messagebox.askyesno("Success", f"Reports generated successfully!\n\nOpen output folder?"):
                os.startfile(output_folder)

        except Exception as e:
            self.log(f"\nERROR: {str(e)}")
            self.update_status("Error occurred")
            messagebox.showerror("Error", f"An error occurred:\n\n{str(e)}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.generate_btn.config(state=tk.NORMAL)

    def process_excel_data(self, excel_file, output_folder):
        """Process Excel file and generate metrics data using process_data.py"""
        self.log("Processing Excel file using process_data.py...")

        # Use the existing process_data.py script
        script_dir = Path(__file__).parent
        process_script = script_dir / 'process_data.py'

        if not process_script.exists():
            self.log("Warning: process_data.py not found, using simplified processing")
            return self.process_excel_data_simple(excel_file, output_folder)

        # Run process_data.py in the output folder
        import subprocess
        original_dir = os.getcwd()
        try:
            os.chdir(output_folder)

            # Copy process_data.py to output folder and modify it
            with open(process_script, 'r', encoding='utf-8') as f:
                script_content = f.read()

            # Replace the hardcoded Excel filename
            script_content = script_content.replace(
                "df = pd.read_excel('BEF ECN FY27.xlsx', sheet_name='Document_TB11')",
                "df = pd.read_excel('source_data.xlsx', sheet_name='Document_TB11')"
            )

            # Replace the hardcoded date filter
            start_date = self.date_range_start.get()
            end_date = self.date_range_end.get()

            if start_date and end_date:
                script_content = script_content.replace(
                    "(df_clean['SubmitDate'] >= '2025-11-01') &\n    (df_clean['SubmitDate'] <= '2026-06-24')",
                    f"(df_clean['SubmitDate'] >= '{start_date}') &\n    (df_clean['SubmitDate'] <= '{end_date}')"
                )
                self.log(f"Using date range: {start_date} to {end_date}")
            else:
                # Remove the date filter entirely
                script_content = script_content.replace(
                    """# Filter data to November 2025 - June 24, 2026
df_clean = df_clean[
    (df_clean['SubmitDate'] >= '2025-11-01') &
    (df_clean['SubmitDate'] <= '2026-06-24')
].copy()

print(f"Filtered to Nov 2025 - Jun 2026: {len(df_clean)} records")""",
                    """# Using all data (no date filter from GUI)
print(f"Using all data: {len(df_clean)} records")"""
                )
                self.log("Using all available data (no date filter)")

            # Write modified script
            with open(output_folder / 'process_data.py', 'w', encoding='utf-8') as f:
                f.write(script_content)

            # Run the script
            result = subprocess.run(
                [sys.executable, 'process_data.py'],
                capture_output=True,
                text=True,
                encoding='utf-8'
            )

            if result.returncode == 0:
                self.log("Data processing completed successfully")
                # Filter out checkmark characters from output
                for line in result.stdout.strip().split('\n'):
                    if line:
                        clean_line = line.replace('✓', '+').replace('•', '-')
                        self.log(f"  {clean_line}")
            else:
                self.log(f"Warning: process_data.py had some issues:")
                for line in result.stderr.strip().split('\n'):
                    if line and 'charmap' not in line:  # Skip encoding warnings
                        self.log(f"  {line}")

            # Read the generated data.json
            json_path = output_folder / 'data.json'
            if json_path.exists():
                with open(json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.log(f"Loaded data.json successfully")
                return data
            else:
                raise Exception("data.json was not created by process_data.py")

        finally:
            os.chdir(original_dir)

    def process_excel_data_simple(self, excel_file, output_folder):
        """Simplified Excel processing (fallback)"""
        self.log("Reading Excel file...")

        # Read Excel
        df = pd.read_excel(excel_file, sheet_name='Document_TB11')
        self.log(f"Loaded {len(df):,} records")

        # Load coordinator name mapping
        coord_dict = {}
        coord_file = os.path.join(os.path.dirname(excel_file), 'ECN Coordinators.xlsx')
        if os.path.exists(coord_file):
            try:
                coord_mapping = pd.read_excel(coord_file)
                coord_dict = dict(zip(coord_mapping['Username'], coord_mapping['Coordinator']))
                self.log(f"Loaded {len(coord_dict)} coordinator name mappings")
            except Exception as e:
                self.log(f"Warning: Could not load coordinator mappings: {str(e)}")
        else:
            self.log("Note: ECN Coordinators.xlsx not found in same folder - using user IDs")

        # Convert dates
        date_columns = ['SubmitDate', 'ProcEndTime', 'FinalDate', 'FinalOrHoldDate', 'HoldDate']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        # Remove missing data
        df_clean = df.dropna(subset=['SubmitDate']).copy()

        # Apply date filter if specified
        start_date = self.date_range_start.get()
        end_date = self.date_range_end.get()

        if start_date and end_date:
            df_clean = df_clean[
                (df_clean['SubmitDate'] >= start_date) &
                (df_clean['SubmitDate'] <= end_date)
            ].copy()
            self.log(f"Filtered to {start_date} - {end_date}: {len(df_clean):,} records")
        else:
            self.log(f"Using all data: {len(df_clean):,} records")

        # Map coordinator user IDs to names
        if coord_dict:
            df_clean['CoordinatorName'] = df_clean['ECNCoordinator'].map(coord_dict).fillna(df_clean['ECNCoordinator'])
        else:
            df_clean['CoordinatorName'] = df_clean['ECNCoordinator']

        # Parse ECN Topics
        df_clean['ECN Topic'] = df_clean['EcnTopic'].str.split('~').str[0].str.strip().str.upper()
        df_clean.loc[~df_clean['EcnTopic'].str.contains('~', na=False), 'ECN Topic'] = \
            df_clean.loc[~df_clean['EcnTopic'].str.contains('~', na=False), 'EcnTopic'].str.strip().str.upper()

        # Filter out ECN Type 9 - matches any format: "9", "(9)", "(9) Restricted ECN", etc.
        df_before_filter = len(df_clean)
        df_clean = df_clean[~df_clean['ECN Topic'].str.contains(r'^\(?\s*9[\)\s\-]|^9$', case=False, na=False, regex=True)].copy()
        if df_before_filter > len(df_clean):
            self.log(f"Filtered out {df_before_filter - len(df_clean)} ECN Type 9 records")

        # Filter for Closed ECNs
        df_closed = df_clean[df_clean['State'] == 'Closed'].copy()
        self.log(f"Closed ECNs: {len(df_closed):,}")
        self.log(f"Void ECNs: {len(df_clean[df_clean['State'] == 'Void']):,}")

        # Calculate percentiles
        p50 = df_closed['ProcCT(days)'].quantile(0.50)
        p75 = df_closed['ProcCT(days)'].quantile(0.75)
        p90 = df_closed['ProcCT(days)'].quantile(0.90)

        self.log(f"50th Percentile: {p50:.2f} days")
        self.log(f"75th Percentile: {p75:.2f} days")
        self.log(f"90th Percentile: {p90:.2f} days")

        # Calculate top categories by percentile range
        fast_range = df_closed[df_closed['ProcCT(days)'] <= p50]
        medium_range = df_closed[(df_closed['ProcCT(days)'] > p50) & (df_closed['ProcCT(days)'] <= p75)]
        slower_range = df_closed[(df_closed['ProcCT(days)'] > p75) & (df_closed['ProcCT(days)'] <= p90)]

        top_50th = fast_range.groupby('ECN Topic').agg({
            'RequestNum': 'count', 'ProcCT(days)': 'mean'
        }).reset_index().sort_values('RequestNum', ascending=False).head(3)

        top_75th = medium_range.groupby('ECN Topic').agg({
            'RequestNum': 'count', 'ProcCT(days)': 'mean'
        }).reset_index().sort_values('RequestNum', ascending=False).head(3)

        top_90th = slower_range.groupby('ECN Topic').agg({
            'RequestNum': 'count', 'ProcCT(days)': 'mean'
        }).reset_index().sort_values('RequestNum', ascending=False).head(3)

        # Calculate monthly trends
        df_clean['YearMonth'] = df_clean['SubmitDate'].dt.to_period('M')
        monthly_trends = df_clean.groupby('YearMonth').agg({
            'ProcCT(days)': 'mean',
            'TotalCT(Days)': 'mean',
            'RequestNum': 'count'
        }).reset_index()
        monthly_trends['YearMonth'] = monthly_trends['YearMonth'].astype(str)

        # Overall statistics
        overall_stats = {
            'total_requests': int(len(df_clean)),
            'total_closed': int(len(df_closed)),
            'total_void': int(len(df_clean[df_clean['State'] == 'Void'])),
            'avg_proc_ct': float(df_clean['ProcCT(days)'].mean()),
            'avg_total_ct': float(df_clean['TotalCT(Days)'].mean()),
            'median_proc_ct': float(df_clean['ProcCT(days)'].median()),
            'median_total_ct': float(df_clean['TotalCT(Days)'].median()),
            'percentile_50_proc_ct': float(p50),
            'percentile_75_proc_ct': float(p75),
            'percentile_90_proc_ct': float(p90),
            'top_categories_50th': top_50th.to_dict('records'),
            'top_categories_75th': top_75th.to_dict('records'),
            'top_categories_90th': top_90th.to_dict('records'),
            'date_range': {
                'start': df_clean['SubmitDate'].min().strftime('%Y-%m-%d'),
                'end': df_clean['SubmitDate'].max().strftime('%Y-%m-%d')
            }
        }

        # Prepare export data
        data_export = {
            'overall_stats': overall_stats,
            'monthly_trends': monthly_trends.to_dict('records')
        }

        # Save to JSON
        json_path = output_folder / 'data.json'
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data_export, f, indent=2, default=str)

        self.log(f"Saved data.json ({len(json.dumps(data_export)) // 1024} KB)")

        return data_export

    def generate_html_dashboard(self, data, output_folder):
        """Generate HTML dashboard - simplified version"""
        html_path = output_folder / 'dashboard.html'

        # Copy and modify the existing dashboard creation
        script_dir = Path(__file__).parent
        dashboard_script = script_dir / 'create_enhanced_dashboard.py'

        if dashboard_script.exists():
            # Run the dashboard generation script in the output folder
            import sys
            import io
            original_dir = os.getcwd()
            original_stdout = sys.stdout
            try:
                os.chdir(output_folder)
                # Redirect stdout to capture print statements with special characters
                sys.stdout = io.StringIO()
                # Execute dashboard creation with UTF-8 encoding
                exec(open(dashboard_script, encoding='utf-8').read())
                self.log("Generated dashboard.html")
            finally:
                sys.stdout = original_stdout
                os.chdir(original_dir)
        else:
            self.log("Warning: Dashboard generation script not found, creating basic HTML")
            self.create_basic_html(data, html_path)

    def create_basic_html(self, data, html_path):
        """Create a basic HTML dashboard if full script not available"""
        stats = data['overall_stats']

        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>BEF ECN Metrics Dashboard</title>
    <meta charset="UTF-8">
    <style>
        body {{
            font-family: 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #2563eb 0%, #1e3a8a 100%);
            padding: 20px;
            margin: 0;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            padding: 30px;
        }}
        h1 {{
            color: #333;
            text-align: center;
        }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }}
        .stat-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }}
        .stat-value {{
            font-size: 2.5em;
            font-weight: bold;
        }}
        .stat-label {{
            font-size: 0.9em;
            margin-top: 10px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>BEF ECN Cycle Time Metrics</h1>
        <p style="text-align: center; color: #666;">
            Period: {stats['date_range']['start']} to {stats['date_range']['end']}
        </p>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">Total Requests</div>
                <div class="stat-value">{stats['total_requests']:,}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Avg Processing CT</div>
                <div class="stat-value">{stats['avg_proc_ct']:.2f}</div>
                <div class="stat-label">days</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">50th Percentile</div>
                <div class="stat-value">{stats['percentile_50_proc_ct']:.2f}</div>
                <div class="stat-label">days</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">75th Percentile</div>
                <div class="stat-value">{stats['percentile_75_proc_ct']:.2f}</div>
                <div class="stat-label">days</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">90th Percentile</div>
                <div class="stat-value">{stats['percentile_90_proc_ct']:.2f}</div>
                <div class="stat-label">days</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Closed ECNs</div>
                <div class="stat-value">{stats['total_closed']:,}</div>
            </div>
        </div>

        <p style="text-align: center; color: #999; margin-top: 50px;">
            Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}
        </p>
    </div>
</body>
</html>"""

        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

    def generate_powerpoint(self, output_folder):
        """Generate PowerPoint presentation using create_powerpoint.py"""
        script_dir = Path(__file__).parent
        pptx_script = script_dir / 'create_powerpoint.py'

        if not pptx_script.exists():
            self.log("Warning: create_powerpoint.py not found, skipping PowerPoint generation")
            return

        import subprocess
        original_dir = os.getcwd()
        try:
            os.chdir(output_folder)

            # Copy the script
            shutil.copy2(pptx_script, output_folder / 'create_powerpoint.py')

            # Run the script
            result = subprocess.run(
                [sys.executable, 'create_powerpoint.py'],
                capture_output=True,
                text=True,
                encoding='utf-8'
            )

            if result.returncode == 0:
                self.log("PowerPoint created successfully")
                for line in result.stdout.strip().split('\n'):
                    if line:
                        clean_line = line.replace('✓', '+').replace('•', '-')
                        self.log(f"  {clean_line}")
            else:
                self.log(f"Warning: PowerPoint generation had issues:")
                for line in result.stderr.strip().split('\n'):
                    if line and 'charmap' not in line:
                        self.log(f"  {line}")

        except Exception as e:
            self.log(f"Error generating PowerPoint: {str(e)}")
        finally:
            os.chdir(original_dir)

    def create_summary(self, data, output_folder):
        """Create summary text file"""
        stats = data['overall_stats']

        summary = f"""================================================================================
BEF ECN CYCLE TIME METRICS REPORT
Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}
================================================================================

DATA PERIOD:
  Start Date: {stats['date_range']['start']}
  End Date:   {stats['date_range']['end']}

SUMMARY STATISTICS:
  Total Requests:       {stats['total_requests']:,}
  Closed ECNs:          {stats['total_closed']:,} ({stats['total_closed']/stats['total_requests']*100:.1f}%)
  Void ECNs:            {stats['total_void']:,} ({stats['total_void']/stats['total_requests']*100:.1f}%)

CYCLE TIME METRICS:
  Average Processing CT: {stats['avg_proc_ct']:.2f} days
  Average Total CT:      {stats['avg_total_ct']:.2f} days
  Median Processing CT:  {stats['median_proc_ct']:.2f} days
  Median Total CT:       {stats['median_total_ct']:.2f} days

PERCENTILES (CLOSED ECNs ONLY):
  50th Percentile:      {stats['percentile_50_proc_ct']:.2f} days
  75th Percentile:      {stats['percentile_75_proc_ct']:.2f} days
  90th Percentile:      {stats['percentile_90_proc_ct']:.2f} days

INTERPRETATION:
  • 50% of closed requests are processed within {stats['percentile_50_proc_ct']:.2f} days
  • 75% of closed requests are processed within {stats['percentile_75_proc_ct']:.2f} days
  • 90% of closed requests are processed within {stats['percentile_90_proc_ct']:.2f} days

GENERATED FILES:
  • dashboard.html - Interactive web dashboard
  • BEF_ECN_Metrics.pptx - PowerPoint presentation
  • data.json - Processed metrics data
  • source_data.xlsx - Copy of Excel source file
  • SUMMARY.txt - This summary file

================================================================================
"""

        summary_path = output_folder / 'SUMMARY.txt'
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write(summary)

        self.log("Generated SUMMARY.txt")

def main():
    root = tk.Tk()
    app = ECNMetricsGenerator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
