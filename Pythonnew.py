import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import numpy as np
import win32com.client as win32
import os

# ----------------- Data Processing -------------------
def load_productivity(data_path):
    cols = ['useralias', 'pipeline', 'p_week', 'processed_volume', 'processed_volume_tr', 'processed_time', 'processed_time_tr']
    dtype = {'useralias': 'string', 'pipeline': 'string', 'p_week': 'int32',
             'processed_volume': 'float32', 'processed_volume_tr': 'float32',
             'processed_time': 'float32', 'processed_time_tr': 'float32'}
    df = pd.read_excel(data_path, usecols=cols, dtype=dtype)
    for c in ['processed_volume', 'processed_volume_tr', 'processed_time', 'processed_time_tr']:
        df[c] = df[c].fillna(0)
    return df

def load_quality(qual_path):
    cols = ['week', 'program', 'auditor_login', 'usecase', 'qc2_judgement', 'qc2_subreason']
    df = pd.read_excel(qual_path, usecols=cols)
    return df[df['program'] == 'RDR'].copy()

def user_prod_dict(df, user, weeks=None):
    dfu = df[df['useralias'] == user].copy()
    if weeks: dfu = dfu[dfu['p_week'].isin(weeks)].copy()
    dfu['total_volume'] = dfu['processed_volume'] + dfu['processed_volume_tr']
    dfu['total_time'] = dfu['processed_time'] + dfu['processed_time_tr']
    g = dfu.groupby(['pipeline', 'p_week'], as_index=False)[['total_volume','total_time']].sum()
    g['productivity'] = np.where(g['total_time'] > 0, g['total_volume']/g['total_time'], 0)
    return g.pivot(index='pipeline', columns='p_week', values='productivity').fillna(0).to_dict('index')

def all_prod_dict(df, weeks=None):
    if weeks: df = df[df['p_week'].isin(weeks)].copy()
    df['total_volume'] = df['processed_volume'] + df['processed_volume_tr']
    df['total_time'] = df['processed_time'] + df['processed_time_tr']
    g = df.groupby(['useralias','pipeline','p_week'], as_index=False)[['total_volume','total_time']].sum()
    g['productivity'] = np.where(g['total_time'] > 0, g['total_volume']/g['total_time'], 0)
    d = {}
    for row in g.itertuples(index=False):
        d.setdefault(row.useralias, {}).setdefault(row.pipeline, {})[row.p_week] = row.productivity
    return d

def productivity_percentiles(allprod, weeks):
    out = {}
    for week in weeks:
        pipeline_vals = {}
        for auditor in allprod.values():
            for pipe, d in auditor.items():
                if week in d: pipeline_vals.setdefault(pipe,[]).append(d[week])
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

def user_quality_dict(df, user, weeks=None):
    dfu = df[df['auditor_login'] == user].copy()
    if weeks: dfu = dfu[dfu['week'].isin(weeks)].copy()
    d = {}
    for (pipe, week), grp in dfu.groupby(['usecase','week']):
        total = len(grp)
        err = grp['qc2_judgement'].isin(['AUDITOR_INCORRECT','BOTH_INCORRECT']).sum()
        score = (total-err)/total if total else 0
        d.setdefault(pipe, {})[week] = {'score': score, 'err': err, 'total': total}
    return d

def all_quality_dict(df, weeks=None):
    if weeks: df = df[df['week'].isin(weeks)].copy()
    d = {}
    for (aud, pipe, week), grp in df.groupby(['auditor_login','usecase','week']):
        total = len(grp)
        err = grp['qc2_judgement'].isin(['AUDITOR_INCORRECT','BOTH_INCORRECT']).sum()
        score = (total-err)/total if total else 0
        d.setdefault(aud, {}).setdefault(pipe, {})[week] = {'score': score, 'err': err, 'total': total}
    return d

def quality_percentiles(allqual, weeks):
    out = {}
    for week in weeks:
        pipeline_vals = {}
        for auditor in allqual.values():
            for pipe, d in auditor.items():
                if week in d: pipeline_vals.setdefault(pipe,[]).append(d[week]['score'])
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

def qc2_subreason_analysis(df, user, weeks=None):
    dfu = df[df['auditor_login'] == user].copy()
    if weeks: dfu = dfu[dfu['week'].isin(weeks)].copy()
    out = {}
    for (pipe, week), grp in dfu.groupby(['usecase','week']):
        incorrect = grp[grp['qc2_judgement'].isin(['AUDITOR_INCORRECT','BOTH_INCORRECT'])]
        total_incorrect = len(incorrect)
        if not total_incorrect: continue
        counts = incorrect['qc2_subreason'].value_counts().to_dict()
        percentages = (incorrect['qc2_subreason'].value_counts(normalize=True)*100).round(1).to_dict()
        out.setdefault(pipe, {})[week] = {
            'counts': counts, 'percentages': percentages, 'total': total_incorrect
        }
    return out

# ------------- Table/Color Utilities -------------
def percentile_label(val, percentiles):
    if percentiles is None or val == 0:
        return "-"
    if val >= percentiles.get('p90', 0): return "P90 +"
    if val >= percentiles.get('p75', 0): return "P75-P90"
    if val >= percentiles.get('p50', 0): return "P50-P75"
    if val >= percentiles.get('p30', 0): return "P30-P50"
    return "<P30"
def pct_bg_fg(bench):
    if bench == "P90 +":
        return "#28a745", "white"  # green
    elif bench == "P75-P90":
        return "#17a2b8", "white"  # teal
    elif bench == "P50-P75":
        return "#ffc107", "black"  # yellow
    elif bench == "P30-P50":
        return "#fd7e14", "white"  # orange
    elif bench == "<P30":
        return "#dc3545", "white"  # red
    else:
        return "#ffffff", "black"  # white default

# ------- Table with Latest Week Benchmark Column -------
def html_metric_value_table_with_latest(data, weeks, percentiles, section="Quality", is_quality=False):
    latest_week = weeks[-1]
    ths = "".join(
        f"<th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>{w}</th>"
        for w in weeks
    )
    ths += "<th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Latest</th>"
    rows = ""
    for pipe in sorted(data):
        tds = ""
        for w in weeks:
            if is_quality:
                val_dict = data[pipe].get(w, None)
                val = val_dict['score'] if isinstance(val_dict, dict) and 'score' in val_dict else 0
                err = val_dict['err'] if isinstance(val_dict, dict) and 'err' in val_dict else 0
                total = val_dict['total'] if isinstance(val_dict, dict) and 'total' in val_dict else 0
                disp = f"{val:.2%} ({err}/{total})" if total else "-"
            else:
                val = data[pipe].get(w, 0)
                disp = f"{val:.2f}" if val else "-"
            color = "#222222" if disp != "-" else "#999"
            tds += f"<td style='padding:6px;border:1px solid #ddd;text-align:center;font-size:13px;color:{color};'>{disp}</td>"

        val = data[pipe].get(latest_week, 0)
        val_scalar = val['score'] if is_quality and isinstance(val, dict) and 'score' in val else val
        pctls = percentiles[latest_week].get(pipe, None) if percentiles and latest_week in percentiles else None
        bench = percentile_label(val_scalar, pctls)
        if bench is None:
            bench = "-"
        bg_color, text_color = pct_bg_fg(bench)
        tds += f"<td style='padding:6px;border:1px solid #ddd;text-align:center;font-size:13px;background-color:{bg_color};color:{text_color};'>{bench if bench != '-' else '-'}</td>"
        rows += f"<tr><td style='padding:6px;border:1px solid #ddd;font-size:13px;text-align:left;'>{pipe}</td>{tds}</tr>"

    return f"""
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:1px solid #ccc;'>
      <tr>
        <th colspan="{len(weeks)+2}" style='background:#000;color:white;text-align:center;font-size:13px;padding:8px;border:1px solid #000;'>
          Weekly {section} Metrics
        </th>
      </tr>
      <tr>
        <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Pipeline</th>
        {ths}
      </tr>
      {rows}
    </table>"""

def html_metric_pct_table(data, weeks, percentiles, section="Quality"):
    ths = "".join(
        f"<th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>{w}</th>"
        for w in weeks
    )
    rows = ""
    for pipe in sorted(data):
        tds = ""
        for w in weeks:
            val = data[pipe].get(w, 0)
            val_scalar = val['score'] if isinstance(val, dict) and 'score' in val else val
            pctls = percentiles[w].get(pipe, None) if percentiles and w in percentiles else None
            bench = percentile_label(val_scalar, pctls)
            if bench is None:
                bench = "-"
            bg_color, text_color = pct_bg_fg(bench)
            tds += f"<td style='padding:6px;border:1px solid #ccc;text-align:center;font-size:13px;background-color:{bg_color};color:{text_color};'>{bench}</td>"
        rows += f"<tr><td style='padding:6px;border:1px solid #ccc;font-size:13px;text-align:left;'>{pipe}</td>{tds}</tr>"

    return f"""
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:1px solid #ccc;'>
      <tr>
        <th colspan="{len(weeks)+1}" style='background:#000;color:white;text-align:center;font-size:13px;padding:8px;border:1px solid #000;'>
          Weekly {section} Benchmarks
        </th>
      </tr>
      <tr>
        <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Pipeline</th>
        {ths}
      </tr>
      {rows}
    </table>"""

def html_qc2_reason_value_table(subreason_data, weeks):
    all_subs = set()
    for pipe in subreason_data:
        for w in subreason_data[pipe]:
            all_subs.update(subreason_data[pipe][w]['counts'].keys())
    ths = "".join(
        f"<th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>{w}</th>"
        for w in weeks
    )
    rows = ""
    for sub in sorted(all_subs):
        tds = ""
        for w in weeks:
            count = sum(
                subreason_data[pipe][w]['counts'].get(sub, 0)
                for pipe in subreason_data if w in subreason_data[pipe]
            )
            disp = str(count) if count else "-"
            color = "#222222" if disp != "-" else "#999"
            tds += f"<td style='padding:6px;border:1px solid #ccc;text-align:center;font-size:13px;color:{color};'>{disp}</td>"
        rows += f"<tr><td style='padding:6px;border:1px solid #ccc;font-size:13px;text-align:left;'>{sub}</td>{tds}</tr>"
    return f"""
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:1px solid #ccc;'>
      <tr>
        <th colspan="{len(weeks)+1}" style='background:#000;color:white;text-align:center;font-size:13px;padding:8px;border:1px solid #000;'>
          QC2 Subreason Counts
        </th>
      </tr>
      <tr>
        <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Subreason</th>
        {ths}
      </tr>
      {rows}
    </table>"""

def html_qc2_reason_pct_table(subreason_data, weeks):
    all_subs = set()
    for pipe in subreason_data:
        for w in subreason_data[pipe]:
            all_subs.update(subreason_data[pipe][w]['counts'].keys())
    ths = "".join(
        f"<th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>{w}</th>"
        for w in weeks
    )
    rows = ""
    for sub in sorted(all_subs):
        tds = ""
        for w in weeks:
            total = sum(subreason_data[pipe][w]['total'] for pipe in subreason_data if w in subreason_data[pipe])
            count = sum(subreason_data[pipe][w]['counts'].get(sub, 0) for pipe in subreason_data if w in subreason_data[pipe])
            perc = (count / total * 100) if total else 0
            disp = f"{perc:.1f}%" if total else "-"
            color = "#222222" if disp != "-" else "#999"
            tds += f"<td style='padding:6px;border:1px solid #ccc;text-align:center;font-size:13px;color:{color};'>{disp}</td>"
        rows += f"<tr><td style='padding:6px;border:1px solid #ccc;font-size:13px;text-align:left;'>{sub}</td>{tds}</tr>"
    return f"""
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:1px solid #ccc;'>
      <tr>
        <th colspan="{len(weeks)+1}" style='background:#000;color:white;text-align:center;font-size:13px;padding:8px;border:1px solid #000;'>
          QC2 Subreason Percentages
        </th>
      </tr>
      <tr>
        <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Subreason</th>
        {ths}
      </tr>
      {rows}
    </table>"""

def html_side_by_side_tables(left, right):
    return f"""
    <table style="width:100%;border:none;">
      <tr>
        <td style="width:49%;vertical-align:top;padding-right:8px;">{left}</td>
        <td style="width:2%"></td>
        <td style="width:49%;vertical-align:top;padding-left:8px;">{right}</td>
      </tr>
    </table>
    """

def compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right):
    return f"""
    <html>
    <body style="margin:0;padding:20px;background:#f9f9fb;font-family:Segoe UI,Arial,sans-serif;color:#333;line-height:1.4;">
        <h2 style="font-size:18px;margin-bottom:10px;font-weight:normal;">RDR Productivity & Quality Metrics Report</h2>
        <p style="margin-top:0;font-size:13px;">Hi {user}, here are your weekly metrics:</p>

        <h3 style="font-size:14px;margin-top:20px;font-weight:normal;">Productivity</h3>
        <table style="width:100%;border:none;margin-bottom:24px;">
            <tr>
                <td style="width:49%;vertical-align:top;padding-right:8px;">{prod_table}</td>
                <td style="width:2%"></td>
                <td style="width:49%;vertical-align:top;padding-left:8px;">{prod_pct_table}</td>
            </tr>
        </table>

        <h3 style="font-size:14px;margin-top:20px;font-weight:normal;">Quality</h3>
        <table style="width:100%;border:none;margin-bottom:24px;">
            <tr>
                <td style="width:49%;vertical-align:top;padding-right:8px;">{qual_table}</td>
                <td style="width:2%"></td>
                <td style="width:49%;vertical-align:top;padding-left:8px;">{qual_pct_table}</td>
            </tr>
        </table>

        <h3 style="font-size:14px;margin-top:20px;font-weight:normal;">QC2 Subreason Analysis</h3>
        <table style="width:100%;border:none;margin-bottom:24px;">
            <tr>
                <td style="width:49%;vertical-align:top;padding-right:8px;">{qc2_left}</td>
                <td style="width:2%"></td>
                <td style="width:49%;vertical-align:top;padding-left:8px;">{qc2_right}</td>
            </tr>
        </table>

        <hr style="margin:24px 0;">
        <p style="font-size:11px;color:#555;">
            <i>Benchmarks: "P90 +" (top 10%), "P75-P90", "P50-P75", "P30-P50", "&lt;P30"<br>
            Questions? Contact your manager.</i>
        </p>
    </body>
    </html>
    """

def send_mail_html(to, cc, subject, html, preview=True):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To, mail.CC, mail.Subject, mail.HTMLBody = to, cc, subject, html
    if preview: mail.Display()
    else: mail.Send()

# ---------------- GUI ----------------
class ReporterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RDR Productivity/Quality Mailer")
        self.geometry("1100x800")
        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self)
        frm.pack(fill="both",expand=True)
        ttk.Label(frm, text="Data Excel:").grid(row=0,column=0,sticky="w"); self.data = tk.StringVar()
        ttk.Entry(frm, textvariable=self.data, width=40).grid(row=0,column=1); ttk.Button(frm, text="Browse", command=lambda:self._browse(self.data)).grid(row=0,column=2)
        ttk.Label(frm, text="Quality Excel:").grid(row=1,column=0,sticky="w"); self.qual = tk.StringVar()
        ttk.Entry(frm, textvariable=self.qual, width=40).grid(row=1,column=1); ttk.Button(frm, text="Browse", command=lambda:self._browse(self.qual)).grid(row=1,column=2)
        ttk.Label(frm, text="Manager Map Excel:").grid(row=2,column=0,sticky="w"); self.mgr = tk.StringVar()
        ttk.Entry(frm, textvariable=self.mgr, width=40).grid(row=2,column=1); ttk.Button(frm, text="Browse", command=lambda:self._browse(self.mgr)).grid(row=2,column=2)
        ttk.Label(frm, text="Weeks (comma):").grid(row=3,column=0,sticky="w"); self.wks = tk.StringVar()
        ttk.Entry(frm, textvariable=self.wks, width=20).grid(row=3,column=1,sticky="w")
        ttk.Label(frm, text="User Login:").grid(row=4,column=0,sticky="w"); self.user = tk.StringVar()
        ttk.Entry(frm, textvariable=self.user, width=20).grid(row=4,column=1,sticky="w")
        self.mode = tk.StringVar(value="preview")
        ttk.Radiobutton(frm, text="Preview", variable=self.mode, value="preview").grid(row=5,column=0)
        ttk.Radiobutton(frm, text="Send", variable=self.mode, value="send").grid(row=5,column=1)
        ttk.Button(frm, text="Send Single", command=self.send_single).grid(row=6,column=0)
        ttk.Button(frm, text="Send Bulk", command=self.send_bulk).grid(row=6,column=1)
        self.status = tk.Text(frm, height=10, width=140); self.status.grid(row=7,column=0,columnspan=3)

    def _browse(self, var):
        var.set(filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xlsm")]))
    def log(self, msg):
        self.status.insert(tk.END, msg+"\n"); self.status.see(tk.END); self.update()
    def _weeks(self):
        try: return [int(w) for w in self.wks.get().split(",") if w.strip()]
        except: self.log("Invalid weeks"); return []
    def send_single(self):
        user = self.user.get().strip()
        if not user or not self.data.get(): return self.log("User and data required.")
        weeks = self._weeks()
        df = load_productivity(self.data.get())
        prod = user_prod_dict(df, user, weeks)
        allprod = all_prod_dict(df, weeks)
        prod_pct = productivity_percentiles(allprod, weeks)
        prod_table = html_metric_value_table_with_latest(prod, weeks, prod_pct, section="Productivity", is_quality=False)
        prod_pct_table = html_metric_pct_table(prod, weeks, prod_pct, section="Productivity")
        qual_table = qual_pct_table = qc2_left = qc2_right = ""
        if self.qual.get() and os.path.exists(self.qual.get()):
            qdf = load_quality(self.qual.get())
            qual = user_quality_dict(qdf, user, weeks)
            allqual = all_quality_dict(qdf, weeks)
            qual_pct = quality_percentiles(allqual, weeks)
            qual_table = html_metric_value_table_with_latest(qual, weeks, qual_pct, section="Quality", is_quality=True)
            qual_pct_table = html_metric_pct_table(qual, weeks, qual_pct, section="Quality")
            subreason = qc2_subreason_analysis(qdf, user, weeks)
            qc2_left = html_qc2_reason_value_table(subreason, weeks)
            qc2_right = html_qc2_reason_pct_table(subreason, weeks)
        html = compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right)
        send_mail_html(f"{user}@amazon.com", "", "RDR Productivity & Quality Metrics Report", html, preview=self.mode.get()=="preview")
        self.log(f"Mail for {user} {'previewed' if self.mode.get()=='preview' else 'sent'}")
    def send_bulk(self):
        if not self.data.get() or not self.mgr.get(): return self.log("Data and manager map required.")
        weeks = self._weeks()
        mgr = pd.read_excel(self.mgr.get())
        df = load_productivity(self.data.get())
        allprod = all_prod_dict(df, weeks)
        prod_pct = productivity_percentiles(allprod, weeks)
        allqual = qual_pct = {}
        if self.qual.get() and os.path.exists(self.qual.get()):
            qdf = load_quality(self.qual.get())
            allqual = all_quality_dict(qdf, weeks)
            qual_pct = quality_percentiles(allqual, weeks)
        for user in mgr['loginname'].dropna().unique():
            prod = allprod.get(user, {})
            prod_table = html_metric_value_table_with_latest(prod, weeks, prod_pct, section="Productivity", is_quality=False)
            prod_pct_table = html_metric_pct_table(prod, weeks, prod_pct, section="Productivity")
            qual = allqual.get(user, {}) if allqual else ""
            qual_table = qual_pct_table = qc2_left = qc2_right = ""
            if allqual:
                qual_table = html_metric_value_table_with_latest(qual, weeks, qual_pct, section="Quality", is_quality=True)
                qual_pct_table = html_metric_pct_table(qual, weeks, qual_pct, section="Quality")
            html = compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right)
            cc = ""
            if 'supervisorloginname' in mgr.columns:
                cc_val = mgr[mgr['loginname']==user]['supervisorloginname'].iloc[0]
                if pd.notna(cc_val): cc = f"{cc_val}@amazon.com"
            send_mail_html(f"{user}@amazon.com", cc, "RDR Productivity & Quality Metrics Report", html, preview=self.mode.get() == "preview")
            self.log(f"Mail for {user} {'previewed' if self.mode.get()=='preview' else 'sent'}")

if __name__ == "__main__":
    ReporterApp().mainloop()
