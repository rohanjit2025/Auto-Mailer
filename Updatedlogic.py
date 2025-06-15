import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import numpy as np
import win32com.client as win32
import os
import requests
import json
from urllib.parse import urljoin
from datetime import datetime, timedelta


class ETLDataFetcher:
    """Custom ETL data fetcher"""

    def __init__(self):
        self.base_url = "https://datacentral.a2z.com/dw-platform/servlet/dwp/template/"
        self.session = requests.Session()
        self.session.headers.update({
            'Content-Type': 'application/json',
        })

    def get_etl_data(self, profile_id, days_back=30):
        """Fetch data from ETL profile"""
        try:
            # For local testing, check if dump file exists
            local_file = f"etl_dump_{profile_id}.xlsx"
            if os.path.exists(local_file):
                return pd.read_excel(local_file)

            # If no local file, try to fetch from ETL
            url = urljoin(self.base_url, f"EtlViewExtractJobs.vm/job_profile_id/{profile_id}")
            response = self.session.get(url)
            response.raise_for_status()
            return pd.DataFrame(response.json())

        except Exception as e:
            raise Exception(f"ETL data fetch failed: {str(e)}")


def load_productivity_from_etl():
    """Fetch productivity data from ETL"""
    try:
        etl = ETLDataFetcher()
        df = etl.get_etl_data("13404076")  # Productivity profile ID

        required_cols = ['useralias', 'pipeline', 'p_week',
                         'processed_volume', 'processed_volume_tr',
                         'processed_time', 'processed_time_tr',
                         'precision_correction']

        df = df[required_cols].astype({
            'useralias': 'string',
            'pipeline': 'string',
            'p_week': 'int32',
            'processed_volume': 'float32',
            'processed_volume_tr': 'float32',
            'processed_time': 'float32',
            'processed_time_tr': 'float32',
            'precision_correction': 'float32'
        })

        return df

    except Exception as e:
        raise Exception(f"Failed to load productivity data: {str(e)}")


def load_timesheet_from_etl():
    """Fetch timesheet data from ETL"""
    try:
        etl = ETLDataFetcher()
        df = etl.get_etl_data("13417956")  # Timesheet profile ID

        required_cols = ['work_date', 'week', 'timesheet_missing', 'loginid']
        df = df[required_cols]
        df['work_date'] = pd.to_datetime(df['work_date'])

        return df

    except Exception as e:
        raise Exception(f"Failed to load timesheet data: {str(e)}")


def load_productivity(data_path):
    """Load productivity data from Excel"""
    try:
        cols = ['useralias', 'pipeline', 'p_week', 'processed_volume', 'processed_volume_tr',
                'processed_time', 'processed_time_tr', 'precision_correction']
        dtype = {'useralias': 'string', 'pipeline': 'string', 'p_week': 'int32',
                 'processed_volume': 'float32', 'processed_volume_tr': 'float32',
                 'processed_time': 'float32', 'processed_time_tr': 'float32',
                 'precision_correction': 'float32'}
        df = pd.read_excel(data_path, usecols=cols, dtype=dtype)
        for c in ['processed_volume', 'processed_volume_tr', 'processed_time', 'processed_time_tr',
                  'precision_correction']:
            df[c] = df[c].fillna(0)
        return df
    except ValueError:
        cols = ['useralias', 'p_week', 'precision_correction']
        df = pd.read_excel(data_path, usecols=cols)
        df['precision_correction'] = df['precision_correction'].fillna(0)
        return df


def load_quality(qual_path):
    """Load quality data from Excel"""
    cols = ['week', 'program', 'auditor_login', 'usecase', 'qc2_judgement',
            'qc2_subreason', 'auditor_correction_type']
    df = pd.read_excel(qual_path, usecols=cols)
    return df[df['program'] == 'RDR']


def user_prod_dict(df, user, weeks=None):
    """Calculate user productivity metrics"""
    dfu = df[df['useralias'] == user]
    if weeks:
        dfu = dfu[dfu['p_week'].isin(weeks)]
    dfu['total_volume'] = dfu['processed_volume'] + dfu['processed_volume_tr']
    dfu['total_time'] = dfu['processed_time'] + dfu['processed_time_tr']
    g = dfu.groupby(['pipeline', 'p_week'], as_index=False)[['total_volume', 'total_time']].sum()
    g['productivity'] = np.where(g['total_time'] > 0, g['total_volume'] / g['total_time'], 0)
    return g.pivot(index='pipeline', columns='p_week', values='productivity').fillna(0).to_dict('index')


def all_prod_dict(df, weeks=None):
    """Calculate productivity metrics for all users"""
    if weeks:
        df = df[df['p_week'].isin(weeks)]
    df['total_volume'] = df['processed_volume'] + df['processed_volume_tr']
    df['total_time'] = df['processed_time'] + df['processed_time_tr']
    g = df.groupby(['useralias', 'pipeline', 'p_week'], as_index=False)[['total_volume', 'total_time']].sum()
    g['productivity'] = np.where(g['total_time'] > 0, g['total_volume'] / g['total_time'], 0)
    d = {}
    for row in g.itertuples(index=False):
        d.setdefault(row.useralias, {}).setdefault(row.pipeline, {})[row.p_week] = row.productivity
    return d


def productivity_percentiles(allprod, weeks):
    """Calculate productivity percentiles"""
    out = {}
    for week in weeks:
        pipeline_vals = {}
        for auditor in allprod.values():
            for pipe, d in auditor.items():
                if week in d:
                    pipeline_vals.setdefault(pipe, []).append(d[week])
        out[week] = {
            pipe: {
                'p30': np.percentile(vals, 30) if vals else 0,
                'p50': np.percentile(vals, 50) if vals else 0,
                'p75': np.percentile(vals, 75) if vals else 0,
                'p90': np.percentile(vals, 90) if vals else 0
            }
            for pipe, vals in pipeline_vals.items()
        }
    return out


class ReporterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📊 RDR Report Generator")
        self.geometry("1200x850")
        self.configure(bg='#2b2d42')

        # Initialize variables
        self.data_source = tk.StringVar(value="both")
        self.report_type = tk.StringVar(value="RDR Report")
        self.mode = tk.StringVar(value="preview")

        # Build UI
        self.setup_styles()
        self.build_ui()

        # Initialize data source tracking
        self.data_sources = {
            "etl": {"status": "not_checked"},
            "excel": {"status": "not_loaded"}
        }

        # Welcome message
        self.log("🚀 Welcome to RDR Report Generator!")
        self.log("📝 Select data sources and configure report settings to begin.")

    def setup_styles(self):
        """Configure UI styles"""
        self.style = ttk.Style()
        self.style.theme_use('alt')

        # Configure styles
        style_configs = {
            'Modern.TLabel': {
                'background': '#2b2d42',
                'foreground': '#edf2f4',
                'font': ('Segoe UI', 10)
            },
            'Title.TLabel': {
                'background': '#2b2d42',
                'foreground': '#8ecae6',
                'font': ('Segoe UI', 18, 'bold')
            },
            'Modern.TButton': {
                'background': '#8ecae6',
                'foreground': '#2b2d42',
                'font': ('Segoe UI', 9, 'bold'),
                'padding': (10, 5)
            },
            'Action.TButton': {
                'background': '#fb8500',
                'foreground': '#2b2d42',
                'font': ('Segoe UI', 10, 'bold'),
                'padding': (15, 8)
            }
        }

        for style, config in style_configs.items():
            self.style.configure(style, **config)

        # Configure button hover effects
        self.style.map('Modern.TButton',
                       background=[('active', '#219ebc'), ('pressed', '#023047')])
        self.style.map('Action.TButton',
                       background=[('active', '#ffb703'), ('pressed', '#8ecae6')])

    def log(self, msg):
        """Add message to status box with timestamp"""
        try:
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_msg = f"[{timestamp}] {msg}\n"

            if hasattr(self, 'status'):
                self.status.config(state='normal')
                self.status.insert(tk.END, formatted_msg)
                self.status.config(state='disabled')
                self.status.see(tk.END)
                self.update_idletasks()
            else:
                print(formatted_msg)  # Fallback if status widget isn't ready
        except Exception as e:
            print(f"Logging error: {str(e)}")

    def _toggle_data_source(self):
        """Toggle data source options"""
        source = self.data_source.get()
        excel_state = "normal" if source in ["both", "excel"] else "disabled"

        # Update UI elements based on selected source
        for frame in [self.data_frame, self.qual_frame, self.timesheet_frame]:
            for child in frame.winfo_children():
                if isinstance(child, (tk.Entry, ttk.Button)):
                    child.config(state=excel_state)

    def _browse(self, var):
        """Open file dialog to browse files"""
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xlsm")],
                title="Select Excel File"
            )
            if filename:
                var.set(filename)
                self.update_status("excel", "not_checked")
                return True
            return False
        except Exception as exc:
            self.log(f"❌ Error browsing file: {str(exc)}")
            return False

    def update_status(self, source, status):
        """Update status indicators"""
        if source not in self.data_sources:
            return

        status_configs = {
            "success": {
                "text": f"✅ {source.upper()}: Connected",
                "fg": '#90EE90'
            },
            "failed": {
                "text": f"❌ {source.upper()}: Failed",
                "fg": '#FF6B6B'
            },
            "not_checked": {
                "text": f"⭕ {source.upper()}: Not checked",
                "fg": '#edf2f4'
            }
        }

        self.data_sources[source]["status"] = status
        if hasattr(self, 'status_labels') and source in self.status_labels:
            config = status_configs.get(status, status_configs["not_checked"])
            self.status_labels[source].config(**config)

    def build_ui(self):
        """Build main UI components"""
        # Main scrollable container
        self.main_container = tk.Frame(self, bg='#2b2d42')
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # Title
        title_label = tk.Label(self.main_container,
                               text="📊 RDR Report Generator",
                               bg='#2b2d42',
                               fg='#8ecae6',
                               font=('Segoe UI', 18, 'bold'))
        title_label.pack(pady=(0, 20))

        # Build sections
        self.build_data_source_section()
        self.build_file_section()
        self.build_params_section()
        self.build_actions_section()
        self.build_status_section()

    def build_data_source_section(self):
        """Build data source selection section"""
        source_frame = self.create_section("Data Sources")

        # Data source selection
        tk.Label(source_frame,
                 text="Select Data Source:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).pack(anchor="w", pady=(0, 5))

        sources_frame = tk.Frame(source_frame, bg='#3d5a80')
        sources_frame.pack(fill="x", pady=(0, 10))

        for value, text in [("both", "Both"), ("etl", "ETL Only"), ("excel", "Excel Only")]:
            tk.Radiobutton(sources_frame,
                           text=text,
                           variable=self.data_source,
                           value=value,
                           bg='#3d5a80',
                           fg='#edf2f4',
                           selectcolor='#2b2d42',
                           command=self._toggle_data_source).pack(side="left", padx=10)

    def build_file_section(self):
        """Build file input section"""
        files_frame = self.create_section("File Selection")

        self.data_frame = self.create_file_input(files_frame, "Productivity Data", "data")
        self.qual_frame = self.create_file_input(files_frame, "Quality Data", "qual")
        self.timesheet_frame = self.create_file_input(files_frame, "Timesheet Data", "timesheet")
        self.mgr_frame = self.create_file_input(files_frame, "Manager Mapping", "mgr")

    def build_params_section(self):
        """Build parameters section"""
        params_frame = self.create_section("Report Parameters")

        # Weeks input
        week_frame = tk.Frame(params_frame, bg='#3d5a80')
        week_frame.pack(fill="x", pady=(0, 10))

        tk.Label(week_frame,
                 text="Weeks (comma-separated):",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).pack(side="left")

        self.wks = tk.StringVar()
        tk.Entry(week_frame,
                 textvariable=self.wks,
                 width=25,
                 font=('Segoe UI', 9),
                 bg='#edf2f4').pack(side="left", padx=10)

        # User input
        user_frame = tk.Frame(params_frame, bg='#3d5a80')
        user_frame.pack(fill="x")

        tk.Label(user_frame,
                 text="User Login:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).pack(side="left")

        self.user = tk.StringVar()
        tk.Entry(user_frame,
                 textvariable=self.user,
                 width=25,
                 font=('Segoe UI', 9),
                 bg='#edf2f4').pack(side="left", padx=10)

    def build_actions_section(self):
        """Build actions section"""
        actions_frame = tk.Frame(self.main_container, bg='#2b2d42', pady=20)
        actions_frame.pack(fill="x")

        buttons_frame = tk.Frame(actions_frame, bg='#2b2d42')
        buttons_frame.pack()

        ttk.Button(buttons_frame,
                   text="Generate Single Report",
                   command=self.send_single,
                   style='Action.TButton').pack(side="left", padx=5)

        ttk.Button(buttons_frame,
                   text="Generate Bulk Reports",
                   command=self.send_bulk,
                   style='Action.TButton').pack(side="left", padx=5)

        # Mode selection
        mode_frame = tk.Frame(buttons_frame, bg='#2b2d42')
        mode_frame.pack(side="left", padx=20)

        tk.Label(mode_frame,
                 text="Mode:",
                 bg='#2b2d42',
                 fg='#edf2f4',
                 font=('Segoe UI', 9, 'bold')).pack(side="left")

        for value, text in [("preview", "Preview"), ("send", "Send")]:
            tk.Radiobutton(mode_frame,
                           text=text,
                           variable=self.mode,
                           value=value,
                           bg='#2b2d42',
                           fg='#edf2f4',
                           selectcolor='#3d5a80').pack(side="left", padx=5)

    def build_status_section(self):
        """Build status section"""
        status_frame = self.create_section("Activity Log")

        # Status text with scrollbar
        text_frame = tk.Frame(status_frame, bg='#3d5a80')
        text_frame.pack(fill="both", expand=True, padx=5, pady=5)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        self.status = tk.Text(text_frame,
                              height=12,
                              width=140,
                              font=('Consolas', 9),
                              bg='#1e2124',
                              fg='#dcddde',
                              wrap=tk.WORD,
                              yscrollcommand=scrollbar.set)
        self.status.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.status.yview)

    def create_section(self, title):
        """Create a section frame with title"""
        frame = tk.LabelFrame(self.main_container,
                              text=title,
                              bg='#3d5a80',
                              fg='#edf2f4',
                              font=('Segoe UI', 11, 'bold'),
                              padx=20,
                              pady=15)
        frame.pack(fill="x", pady=10)
        return frame

    def create_file_input(self, parent, label, var_name):
        """Create a file input row"""
        frame = tk.Frame(parent, bg='#3d5a80')
        frame.pack(fill="x", pady=5)

        tk.Label(frame,
                 text=f"{label}:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).pack(side="left")

        setattr(self, var_name, tk.StringVar())
        entry = tk.Entry(frame,
                         textvariable=getattr(self, var_name),
                         width=60,
                         font=('Segoe UI', 9),
                         bg='#edf2f4')
        entry.pack(side="left", padx=10)

        ttk.Button(frame,
                   text="Browse",
                   command=lambda: self._browse(getattr(self, var_name)),
                   style='Modern.TButton').pack(side="left")

        return frame

def show_loading(self):
    """Show loading indicator"""
    self.loading_window = tk.Toplevel(self)
    self.loading_window.title("Processing")
    self.loading_window.geometry("300x150")
    self.loading_window.configure(bg='#2b2d42')

    frame = tk.Frame(self.loading_window, bg='#2b2d42')
    frame.pack(expand=True)

    tk.Label(frame,
             text="Processing...\nPlease wait",
             bg='#2b2d42',
             fg='#edf2f4',
             font=('Segoe UI', 12)).pack(pady=20)

    self.progress = ttk.Progressbar(frame,
                                    mode='indeterminate',
                                    length=200)
    self.progress.pack(pady=10)
    self.progress.start(10)

    self.loading_window.transient(self)
    self.loading_window.grab_set()
    self.loading_window.update()


def hide_loading(self):
    """Hide loading indicator"""
    if hasattr(self, 'loading_window'):
        self.progress.stop()
        self.loading_window.destroy()

    def send_email(self, to_email, subject, html_content, cc_email=""):
        """Send email using Outlook"""
        try:
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.To = to_email
            mail.CC = cc_email
            mail.Subject = subject
            mail.HTMLBody = html_content

            if self.mode.get() == "preview":
                mail.Display()
                self.log(f"📧 Email previewed for {to_email}")
            else:
                mail.Send()
                self.log(f"✅ Email sent to {to_email}")

        except Exception as e:
            self.log(f"❌ Error sending email to {to_email}: {str(e)}")
            raise

    def generate_html_report(self, user, prod_data, qual_data=None, timesheet_data=None):
        """Generate HTML report content"""
        html = f"""
        <html>
        <body style="margin:0;padding:20px;background:#f9f9fb;font-family:Segoe UI,Arial,sans-serif;">
            <h2 style="font-size:18px;margin-bottom:10px;">RDR Performance Report</h2>
            <p style="font-size:13px;">Hi {user}, here is your performance report:</p>
        """

        # Add productivity section
        if prod_data:
            html += self.generate_productivity_section(prod_data)

        # Add quality section if data available
        if qual_data:
            html += self.generate_quality_section(qual_data)

        # Add timesheet section if data available
        if timesheet_data:
            html += self.generate_timesheet_section(timesheet_data)

        # Add footer
        html += """
            <hr style="margin:24px 0;">
            <p style="font-size:11px;color:#555;">
                <i>For questions about this report, please contact your manager.</i>
            </p>
        </body>
        </html>
        """
        return html

    def generate_productivity_section(self, data):
        """Generate productivity section of report"""
        html = """
        <div style="margin-top:20px;">
            <h3 style="font-size:14px;margin-bottom:10px;">Productivity Metrics</h3>
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
                <tr style="background:#2C3E50;color:white;">
                    <th style="padding:8px;border:1px solid #ddd;">Pipeline</th>
                    <th style="padding:8px;border:1px solid #ddd;">Value</th>
                    <th style="padding:8px;border:1px solid #ddd;">Benchmark</th>
                </tr>
        """

        for row in data:
            html += f"""
                <tr>
                    <td style="padding:8px;border:1px solid #ddd;">{row['pipeline']}</td>
                    <td style="padding:8px;border:1px solid #ddd;">{row['value']:.2f}</td>
                    <td style="padding:8px;border:1px solid #ddd;{self.get_benchmark_style(row['benchmark'])}">
                        {row['benchmark']}
                    </td>
                </tr>
            """

        html += "</table></div>"
        return html

    def get_benchmark_style(self, benchmark):
        """Get CSS style for benchmark cell"""
        styles = {
            'P90+': 'background:#006400;color:white;',
            'P75-P90': 'background:#90EE90;color:black;',
            'P50-P75': 'background:#FFD700;color:black;',
            'P30-P50': 'background:#ffc107;color:black;',
            '<P30': 'background:#dc3545;color:white;'
        }
        return styles.get(benchmark, '')

    def send_single(self):
        """Handle single report generation"""
        try:
            user = self.user.get().strip()
            if not user:
                self.log("❌ User is required")
                return

            weeks = self._weeks()
            if not weeks:
                self.log("❌ No valid weeks specified")
                return

            self.show_loading()
            self.log(f"🚀 Generating report for {user}...")

            # Load data
            prod_data = self.load_data_source("productivity")
            if prod_data is None:
                self.log("❌ No productivity data available")
                return

            # Process data
            processed_data = self.process_user_data(user, prod_data, weeks)

            # Generate and send report
            html = self.generate_html_report(user, processed_data)
            self.send_email(
                f"{user}@amazon.com",
                "RDR Performance Report",
                html
            )

        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
        finally:
            self.hide_loading()

    def send_bulk(self):
        """Handle bulk report generation"""
        try:
            if not self.mgr.get():
                self.log("❌ Manager mapping file is required")
                return

            mgr_df = pd.read_excel(self.mgr.get())
            total_users = len(mgr_df['loginname'].unique())
            self.log(f"🚀 Processing {total_users} users...")

            # Load data once
            prod_data = self.load_data_source("productivity")
            if prod_data is None:
                self.log("❌ No productivity data available")
                return

            for user in mgr_df['loginname'].unique():
                try:
                    # Get manager email for CC
                    manager = mgr_df[mgr_df['loginname'] == user]['supervisorloginname'].iloc[0]
                    cc_email = f"{manager}@amazon.com" if pd.notna(manager) else ""

                    # Process user data
                    processed_data = self.process_user_data(user, prod_data, self._weeks())

                    # Generate and send report
                    html = self.generate_html_report(user, processed_data)
                    self.send_email(
                        f"{user}@amazon.com",
                        "RDR Performance Report",
                        html,
                        cc_email
                    )

                except Exception as e:
                    self.log(f"⚠️ Error processing {user}: {str(e)}")
                    continue

        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
        finally:
            self.hide_loading()

    def process_user_data(self, user, prod_data, weeks):
        """Process user data for reporting"""
        user_prod = user_prod_dict(prod_data, user, weeks)
        all_prod = all_prod_dict(prod_data, weeks)
        prod_percentiles = productivity_percentiles(all_prod, weeks)

        processed_data = []
        for pipeline, data in user_prod.items():
            latest_week = weeks[-1]
            latest_value = data.get(latest_week, 0)

            processed_data.append({
                'pipeline': pipeline,
                'value': latest_value,
                'benchmark': self.get_benchmark(
                    latest_value,
                    prod_percentiles[latest_week].get(pipeline, {})
                )
            })

        return processed_data

    def get_benchmark(self, value, percentiles):
        """Get benchmark label based on percentiles"""
        if not percentiles or value == 0:
            return "-"
        if value >= percentiles.get('p90', 0): return "P90+"
        if value >= percentiles.get('p75', 0): return "P75-P90"
        if value >= percentiles.get('p50', 0): return "P50-P75"
        if value >= percentiles.get('p30', 0): return "P30-P50"
        return "<P30"


# Main execution
if __name__ == "__main__":
    try:
        # Set up logging
        import logging

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('app.log'),
                logging.StreamHandler()
            ]
        )

        # Start application
        app = ReporterApp()
        app.mainloop()

    except Exception as e:
        logging.error(f"Application failed to start: {str(e)}")
        print(f"Critical Error: {str(e)}")