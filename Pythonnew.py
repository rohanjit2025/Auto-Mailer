import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import numpy as np
import win32com.client as win32
import os

# ----------------- Data Processing -------------------
def load_productivity(data_path):
    try:
        # Try original columns first
        cols = ['useralias', 'pipeline', 'p_week', 'processed_volume', 'processed_volume_tr',
                'processed_time', 'processed_time_tr', 'precision_correction']  # Added precision_correction
        dtype = {'useralias': 'string', 'pipeline': 'string', 'p_week': 'int32',
                'processed_volume': 'float32', 'processed_volume_tr': 'float32',
                'processed_time': 'float32', 'processed_time_tr': 'float32',
                'precision_correction': 'float32'}  # Added precision_correction dtype
        df = pd.read_excel(data_path, usecols=cols, dtype=dtype)
        for c in ['processed_volume', 'processed_volume_tr', 'processed_time', 'processed_time_tr', 'precision_correction']:
            df[c] = df[c].fillna(0)
    except ValueError:
        # If original columns not found, try loading with minimal required columns for precision report
        cols = ['useralias', 'p_week', 'precision_correction']
        df = pd.read_excel(data_path, usecols=cols)
        df['precision_correction'] = df['precision_correction'].fillna(0)
    return df


def calculate_precision_corrections(df, weeks=None):
    """Calculate weekly precision corrections for each auditor"""
    if weeks:
        df = df[df['p_week'].isin(weeks)].copy()

    # Group by user and week, sum precision_correction
    precision_data = df.groupby(['useralias', 'p_week'])['precision_correction'].sum().reset_index()

    # Filter users with more than 20 corrections in a week
    high_precision = precision_data[precision_data['precision_correction'] > 20]

    # Convert to dictionary format
    result = {}
    for _, row in high_precision.iterrows():
        result.setdefault(row['useralias'], {})[row['p_week']] = row['precision_correction']

    return result

def load_quality(qual_path):
    cols = ['week', 'program', 'auditor_login', 'usecase', 'qc2_judgement', 'qc2_subreason','auditor_correction_type']
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


def analyze_timesheet_missing(timesheet_path, week, user):  # Added user parameter
    """Analyze timesheet missing data for a specific week and user"""
    try:
        # Read Excel file
        df = pd.read_excel(timesheet_path)

        # Make sure the required columns exist
        required_columns = ['work_date', 'week', 'timesheet_missing', 'loginid']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

        # Filter for the specified week and user
        weekly_data = df[(df['week'] == week) & (df['loginid'] == user)].copy()

        # Convert work_date to datetime if it's not already
        weekly_data['work_date'] = pd.to_datetime(weekly_data['work_date'])

        # Filter for missing time > 35 minutes and select only required columns
        daily_missing = weekly_data[weekly_data['timesheet_missing'] > 35][['work_date', 'timesheet_missing']]

        # Sort by date
        daily_missing = daily_missing.sort_values('work_date')

        return daily_missing if not daily_missing.empty else None

    except Exception as exc:
        print(f"Error processing timesheet: {str(exc)}")
        return None

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
        return "#006400", "white"  # strong/dark green
    elif bench == "P75-P90":
        return "#90EE90", "black"  # light green
    elif bench == "P50-P75":
        return "#FFD700", "black"  # bright yellow
    elif bench == "P30-P50":
        return "#ffc107", "black"  # yellow
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
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:2px solid #000000;'>
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


def html_precision_correction_table(data, weeks):
    ths = "".join(
        f"<th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>{w}</th>"
        for w in weeks
    )

    rows = ""
    for user in sorted(data):
        tds = ""
        for w in weeks:
            val = data[user].get(w, 0)
            disp = f"{val:.0f}" if val else "-"
            color = "#222222" if disp != "-" else "#999"
            tds += f"<td style='padding:6px;border:1px solid #ddd;text-align:center;font-size:13px;color:{color};'>{disp}</td>"

        rows += f"<tr><td style='padding:6px;border:1px solid #ddd;font-size:13px;text-align:left;'>{user}</td>{tds}</tr>"

    return f"""
    <table style='border-collapse:collapse;width:40%;font-family:Segoe UI,Arial,sans-serif;border:2px solid #000000;'>
      <tr>
        <th colspan="{len(weeks) + 1}" style='background:#000;color:white;text-align:center;font-size:13px;padding:8px;border:1px solid #000;'>
          Weekly Precision Corrections (>20)
        </th>
      </tr>
      <tr>
        <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Auditor</th>
        {ths}
      </tr>
      {rows}
    </table>"""


def html_timesheet_missing_table(timesheet_data):
    """Generate HTML table for timesheet missing data"""
    if timesheet_data is None or timesheet_data.empty:
        return "<p>No timesheet missing data found exceeding 35 minutes.</p>"

    try:
        rows = ""
        for _, row in timesheet_data.iterrows():
            work_date = row['work_date'].strftime('%Y-%m-%d')
            missing = f"{row['timesheet_missing']:.2f}"
            rows += f"<tr><td style='padding:6px;border:1px solid #ddd;font-size:13px;text-align:left;'>{work_date}</td>"
            rows += f"<td style='padding:6px;border:1px solid #ddd;text-align:center;font-size:13px;'>{missing}</td></tr>"

        return f"""
        <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:2px solid #000000;'>
          <tr>
            <th colspan="2" style='background:#000;color:white;text-align:center;font-size:13px;padding:8px;border:1px solid #000;'>
              Timesheet Missing Report (>35 minutes)
            </th>
          </tr>
          <tr>
            <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Date</th>
            <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Missing Minutes</th>
          </tr>
          {rows}
        </table>"""
    except Exception as e:
        print(f"Error generating HTML table: {str(e)}")
        return "<p>Error generating timesheet report.</p>"

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
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:2px solid #000000;'>
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


def calculate_correction_type_data(df, user, weeks=None):
    """Calculate correction type metrics for a user"""
    dfu = df[df['auditor_login'] == user].copy()
    if weeks:
        dfu = dfu[dfu['week'].isin(weeks)].copy()

    correction_data = {}
    for (correction_type, week), grp in dfu.groupby(['auditor_correction_type', 'week']):
        total = len(grp)
        err = grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']).sum()
        score = (total - err) / total if total else 0
        correction_data.setdefault(correction_type, {})[week] = {
            'score': score,
            'err': err,
            'total': total
        }
    return correction_data

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
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:2px solid #000000;'>
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
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:2px solid #000000;'>
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


def html_correction_type_quality_table(correction_data, weeks):
    ths = "".join(
        f"<th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>{w}</th>"
        for w in weeks
    )

    rows = ""
    for correction_type in sorted(correction_data):
        tds = ""
        for w in weeks:
            val_dict = correction_data[correction_type].get(w, {})
            val = val_dict.get('score', 0)
            err = val_dict.get('err', 0)
            total = val_dict.get('total', 0)
            disp = f"{val:.2%} ({err}/{total})" if total else "-"
            color = "#222222" if disp != "-" else "#999"
            tds += f"<td style='padding:6px;border:1px solid #ddd;text-align:center;font-size:13px;color:{color};'>{disp}</td>"

        rows += f"<tr><td style='padding:6px;border:1px solid #ddd;font-size:13px;text-align:left;'>{correction_type}</td>{tds}</tr>"

    return f"""
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:2px solid #000000;'>
      <tr>
        <th colspan="{len(weeks) + 1}" style='background:#000;color:white;text-align:center;font-size:13px;padding:8px;border:1px solid #000;'>
          Weekly Correction Type Quality
        </th>
      </tr>
      <tr>
        <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Correction Type</th>
        {ths}
      </tr>
      {rows}
    </table>"""


def html_correction_type_count_table(correction_data, weeks):
    ths = "".join(
        f"<th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>{w}</th>"
        for w in weeks
    )

    rows = ""
    for correction_type in sorted(correction_data):
        tds = ""
        for w in weeks:
            val_dict = correction_data[correction_type].get(w, {})
            count = val_dict.get('total', 0)
            disp = str(count) if count else "-"
            color = "#222222" if disp != "-" else "#999"
            tds += f"<td style='padding:6px;border:1px solid #ddd;text-align:center;font-size:13px;color:{color};'>{disp}</td>"

        rows += f"<tr><td style='padding:6px;border:1px solid #ddd;font-size:13px;text-align:left;'>{correction_type}</td>{tds}</tr>"

    return f"""
    <table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;border:2px solid #000000;'>
      <tr>
        <th colspan="{len(weeks) + 1}" style='background:#000;color:white;text-align:center;font-size:13px;padding:8px;border:1px solid #000;'>
          Weekly Correction Type Case Counts
        </th>
      </tr>
      <tr>
        <th style='padding:6px;border:1px solid #ccc;background:#2C3E50;color:white;text-align:center;font-size:13px;'>Correction Type</th>
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

def compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right,
                correction_quality_table, correction_count_table,timesheet_table=""):
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

        <h3 style="font-size:14px;margin-top:20px;font-weight:normal;">Correction Type Analysis</h3>
        <table style="width:100%;border:none;margin-bottom:24px;">
            <tr>
                <td style="width:49%;vertical-align:top;padding-right:8px;">{correction_count_table}</td>
                <td style="width:2%"></td>
                <td style="width:49%;vertical-align:top;padding-left:8px;">{correction_quality_table}</td>
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
        self.title("📊 Email Report Generator")
        self.geometry("1200x850")
        # Modern dark background
        self.configure(bg='#2b2d42')

        # Configure style - Use a theme that works with dark colors
        self.style = ttk.Style()
        self.style.theme_use('alt')  # 'alt' theme works better for custom colors

        # Configure the styles that actually work
        self.style.configure('Modern.TLabel',
                             background='#2b2d42',
                             foreground='#edf2f4',
                             font=('Segoe UI', 10))

        self.style.configure('Title.TLabel',
                             background='#2b2d42',
                             foreground='#8ecae6',
                             font=('Segoe UI', 18, 'bold'))

        self.style.configure('Heading.TLabel',
                             background='#2b2d42',
                             foreground='#ffb3c6',
                             font=('Segoe UI', 11, 'bold'))

        # Buttons that will actually show up
        self.style.configure('Modern.TButton',
                             background='#8ecae6',
                             foreground='#2b2d42',
                             font=('Segoe UI', 9, 'bold'),
                             padding=(10, 5))

        self.style.map('Modern.TButton',
                       background=[('active', '#219ebc'),
                                   ('pressed', '#023047')])

        self.style.configure('Action.TButton',
                             background='#fb8500',
                             foreground='#2b2d42',
                             font=('Segoe UI', 10, 'bold'),
                             padding=(15, 8))

        self.style.map('Action.TButton',
                       background=[('active', '#ffb703'),
                                   ('pressed', '#8ecae6')])

        self._build_ui()
        self._update_fields()

    def _weeks(self):
        """Convert comma-separated weeks string to list of integers"""
        try:
            weeks = [int(w.strip()) for w in self.wks.get().split(',') if w.strip()]
            if not weeks:
                self.log("❌ No valid weeks specified")
                return []
            return weeks
        except ValueError:
            self.log("❌ Invalid week format. Please use comma-separated numbers.")
            return []

    def _browse(self, var):
        """Open file dialog to browse files"""
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx *.xlsm")],
                title="Select Excel File"
            )
            if filename:
                var.set(filename)
                return True
            return False
        except Exception as exc:
            self.log(f"❌ Error browsing file: {str(exc)}")
            return False

    def _update_fields(self):
        """Update field visibility based on selected report type"""
        report_type = self.report_type.get()

        # Hide all fields first
        for widget in [self.data_frame, self.qual_frame, self.mgr_frame, self.timesheet_frame]:
            widget.grid_remove()

        # Show fields based on report type
        if report_type == "RDR Report":
            self.data_frame.grid()
            self.qual_frame.grid()
            self.mgr_frame.grid()
            self.info_label.config(text="📋 RDR Report requires Data Excel, Quality Excel, and Manager Mapping files")

        elif report_type == "Precision Correction Report":
            self.data_frame.grid()
            self.mgr_frame.grid()
            self.info_label.config(text="🎯 Precision Correction Report requires Data Excel and Manager Mapping files")

        elif report_type == "Timesheet Missing Report":
            self.timesheet_frame.grid()
            self.mgr_frame.grid()
            self.info_label.config(text="⏰ Timesheet Missing Report requires Timesheet Excel and Manager Mapping files")

    def _build_ui(self):
        # Main container with modern background
        # Create a canvas and a vertical scrollbar linked to it
        # Scrollable canvas setup
        canvas = tk.Canvas(self, borderwidth=0, background="#2b2d42", highlightthickness=0)
        vsb = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        main_frame = tk.Frame(canvas, bg="#2b2d42")

        main_window = canvas.create_window((0, 0), window=main_frame, anchor="nw", width=self.winfo_width())

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def resize_main(event):
            canvas.itemconfig(main_window, width=event.width)

        main_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", resize_main)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # Linux scroll up
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))  # Linux scroll down
        # Title
        title_label = tk.Label(main_frame,
                               text="📊 Email Report Generator",
                               bg='#2b2d42',
                               fg='#8ecae6',
                               font=('Segoe UI', 18, 'bold'))
        title_label.pack(pady=(0, 25))

        # Content frame
        content_frame = tk.Frame(main_frame, bg='#2b2d42')
        content_frame.pack(fill="both", expand=True)

        # Report Type Selection Section - Modern card style
        report_section = tk.LabelFrame(content_frame,
                                       text="📋 Report Configuration",
                                       bg='#3d5a80',
                                       fg='#edf2f4',
                                       font=('Segoe UI', 11, 'bold'),
                                       padx=20, pady=15,
                                       relief='flat',
                                       bd=2)
        report_section.pack(fill="x", pady=(0, 20))

        tk.Label(report_section,
                 text="Report Type:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.report_type = ttk.Combobox(
            report_section,
            values=["RDR Report", "Precision Correction Report", "Timesheet Missing Report"],
            state="readonly",
            font=('Segoe UI', 10),
            width=35
        )
        self.report_type.set("RDR Report")
        self.report_type.grid(row=1, column=0, sticky="w", pady=(0, 10))
        self.report_type.bind('<<ComboboxSelected>>', lambda e: self._update_fields())

        # Info label
        self.info_label = tk.Label(report_section,
                                   text="",
                                   bg='#3d5a80',
                                   fg='#edf2f4',
                                   font=('Segoe UI', 9),
                                   wraplength=500,
                                   justify='left')
        self.info_label.grid(row=2, column=0, sticky="w")

        # File Input Section
        files_section = tk.LabelFrame(content_frame,
                                      text="📁 File Selection",
                                      bg='#3d5a80',
                                      fg='#edf2f4',
                                      font=('Segoe UI', 11, 'bold'),
                                      padx=20, pady=15,
                                      relief='flat',
                                      bd=2)
        files_section.pack(fill="x", pady=(0, 20))

        # Data Excel Frame
        self.data_frame = tk.Frame(files_section, bg='#3d5a80')
        self.data_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))

        tk.Label(self.data_frame,
                 text="📊 Data Excel:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky="w")

        self.data = tk.StringVar()
        data_entry = tk.Entry(self.data_frame,
                              textvariable=self.data,
                              width=60,
                              font=('Segoe UI', 9),
                              bg='#edf2f4',
                              fg='#2b2d42',
                              relief='flat',
                              bd=5)
        data_entry.grid(row=1, column=0, padx=(0, 10), pady=(5, 0))

        ttk.Button(self.data_frame,
                   text="Browse",
                   command=lambda: self._browse(self.data),
                   style='Modern.TButton').grid(row=1, column=1, pady=(5, 0))

        # Quality Excel Frame
        self.qual_frame = tk.Frame(files_section, bg='#3d5a80')
        self.qual_frame.grid(row=1, column=0, sticky="ew", pady=(0, 15))

        tk.Label(self.qual_frame,
                 text="🎖️ Quality Excel:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky="w")

        self.qual = tk.StringVar()
        qual_entry = tk.Entry(self.qual_frame,
                              textvariable=self.qual,
                              width=60,
                              font=('Segoe UI', 9),
                              bg='#edf2f4',
                              fg='#2b2d42',
                              relief='flat',
                              bd=5)
        qual_entry.grid(row=1, column=0, padx=(0, 10), pady=(5, 0))

        ttk.Button(self.qual_frame,
                   text="Browse",
                   command=lambda: self._browse(self.qual),
                   style='Modern.TButton').grid(row=1, column=1, pady=(5, 0))

        # Manager Map Excel Frame
        self.mgr_frame = tk.Frame(files_section, bg='#3d5a80')
        self.mgr_frame.grid(row=2, column=0, sticky="ew", pady=(0, 15))

        tk.Label(self.mgr_frame,
                 text="👥 Manager Mapping Excel:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky="w")

        self.mgr = tk.StringVar()
        mgr_entry = tk.Entry(self.mgr_frame,
                             textvariable=self.mgr,
                             width=60,
                             font=('Segoe UI', 9),
                             bg='#edf2f4',
                             fg='#2b2d42',
                             relief='flat',
                             bd=5)
        mgr_entry.grid(row=1, column=0, padx=(0, 10), pady=(5, 0))

        ttk.Button(self.mgr_frame,
                   text="Browse",
                   command=lambda: self._browse(self.mgr),
                   style='Modern.TButton').grid(row=1, column=1, pady=(5, 0))

        # Timesheet Excel Frame
        self.timesheet_frame = tk.Frame(files_section, bg='#3d5a80')
        self.timesheet_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        tk.Label(self.timesheet_frame,
                 text="⏰ Timesheet Excel:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky="w")

        self.timesheet = tk.StringVar()
        timesheet_entry = tk.Entry(self.timesheet_frame,
                                   textvariable=self.timesheet,
                                   width=60,
                                   font=('Segoe UI', 9),
                                   bg='#edf2f4',
                                   fg='#2b2d42',
                                   relief='flat',
                                   bd=5)
        timesheet_entry.grid(row=1, column=0, padx=(0, 10), pady=(5, 0))

        ttk.Button(self.timesheet_frame,
                   text="Browse",
                   command=lambda: self._browse(self.timesheet),
                   style='Modern.TButton').grid(row=1, column=1, pady=(5, 0))

        # Configure grid weights
        files_section.columnconfigure(0, weight=1)
        for frame in [self.data_frame, self.qual_frame, self.mgr_frame, self.timesheet_frame]:
            frame.columnconfigure(0, weight=1)

        # Parameters Section
        params_section = tk.LabelFrame(content_frame,
                                       text="⚙️ Parameters",
                                       bg='#3d5a80',
                                       fg='#edf2f4',
                                       font=('Segoe UI', 11, 'bold'),
                                       padx=20, pady=15,
                                       relief='flat',
                                       bd=2)
        params_section.pack(fill="x", pady=(0, 20))

        params_frame = tk.Frame(params_section, bg='#3d5a80')
        params_frame.pack(fill="x")

        # Left side parameters
        left_params = tk.Frame(params_frame, bg='#3d5a80')
        left_params.pack(side="left", fill="x", expand=True)

        tk.Label(left_params,
                 text="📅 Weeks (comma-separated):",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.wks = tk.StringVar()
        weeks_entry = tk.Entry(left_params,
                               textvariable=self.wks,
                               width=25,
                               font=('Segoe UI', 9),
                               bg='#edf2f4',
                               fg='#2b2d42',
                               relief='flat',
                               bd=5)
        weeks_entry.grid(row=1, column=0, sticky="w")

        tk.Label(left_params,
                 text="👤 User Login (for single reports):",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).grid(row=0, column=1, sticky="w", padx=(30, 0), pady=(0, 5))

        self.user = tk.StringVar()
        user_entry = tk.Entry(left_params,
                              textvariable=self.user,
                              width=25,
                              font=('Segoe UI', 9),
                              bg='#edf2f4',
                              fg='#2b2d42',
                              relief='flat',
                              bd=5)
        user_entry.grid(row=1, column=1, sticky="w", padx=(30, 0))

        # Right side - Mode selection
        right_params = tk.Frame(params_frame, bg='#3d5a80')
        right_params.pack(side="right", padx=(20, 0))

        tk.Label(right_params,
                 text="📧 Email Mode:",
                 bg='#3d5a80',
                 fg='#ffb3c6',
                 font=('Segoe UI', 11, 'bold')).pack(anchor="w")

        mode_frame = tk.Frame(right_params, bg='#3d5a80')
        mode_frame.pack(fill="x", pady=(5, 0))

        self.mode = tk.StringVar(value="preview")
        preview_rb = tk.Radiobutton(mode_frame,
                                    text="👁️ Preview",
                                    variable=self.mode,
                                    value="preview",
                                    bg='#3d5a80',
                                    fg='#edf2f4',
                                    selectcolor='#8ecae6',
                                    font=('Segoe UI', 9))
        preview_rb.pack(side="left", padx=(0, 15))

        send_rb = tk.Radiobutton(mode_frame,
                                 text="📤 Send",
                                 variable=self.mode,
                                 value="send",
                                 bg='#3d5a80',
                                 fg='#edf2f4',
                                 selectcolor='#8ecae6',
                                 font=('Segoe UI', 9))
        send_rb.pack(side="left")

        # Action Buttons Section
        actions_section = tk.Frame(content_frame, bg='#2b2d42')
        actions_section.pack(fill="x", pady=(0, 20))

        button_frame = tk.Frame(actions_section, bg='#2b2d42')
        button_frame.pack()

        # Action buttons
        single_btn = ttk.Button(button_frame,
                                text="📧 Send Single Report",
                                command=self.send_single,
                                style='Action.TButton')
        single_btn.pack(side="left", padx=(0, 20))

        bulk_btn = ttk.Button(button_frame,
                              text="📨 Send Bulk Reports",
                              command=self.send_bulk,
                              style='Action.TButton')
        bulk_btn.pack(side="left")

        # Status Section
        status_section = tk.LabelFrame(content_frame,
                                       text="💬 Activity Log",
                                       bg='#3d5a80',
                                       fg='#edf2f4',
                                       font=('Segoe UI', 11, 'bold'),
                                       padx=15, pady=10,
                                       relief='flat',
                                       bd=2)
        status_section.pack(fill="both", expand=True)

        # Status text with scrollbar - FIXED VERSION
        status_frame = tk.Frame(status_section, bg='#3d5a80')
        status_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Create scrollbar first
        scrollbar = tk.Scrollbar(status_frame,
                                 orient="vertical",
                                 bg='#2b2d42',
                                 troughcolor='#1e2124',
                                 activebackground='#8ecae6')
        scrollbar.pack(side="right", fill="y")

        # Create text widget and connect to scrollbar
        self.status = tk.Text(
            status_frame,
            height=12,
            width=140,
            font=('Consolas', 9),
            bg='#1e2124',  # Dark code-like background
            fg='#dcddde',  # Light gray text
            insertbackground='#8ecae6',  # Blue cursor
            selectbackground='#36393f',  # Selection background
            selectforeground='#dcddde',
            wrap=tk.WORD,
            relief='flat',
            bd=0,
            yscrollcommand=scrollbar.set  # Connect text to scrollbar
        )
        self.status.pack(side="left", fill="both", expand=True)

        # Connect scrollbar to text widget
        scrollbar.config(command=self.status.yview)

        # ALSO UPDATE YOUR log METHOD TO ENSURE AUTO-SCROLL:
        def log(self, msg):
            """Add message to status box with timestamp"""
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_msg = f"[{timestamp}] {msg}\n"

            self.status.config(state='normal')  # Enable editing
            self.status.insert(tk.END, formatted_msg)
            self.status.config(state='disabled')  # Make read-only
            self.status.see(tk.END)  # Auto-scroll to bottom
            self.update_idletasks()  # Force update

        # ALTERNATIVE: If you want a more modern scrollbar, use ttk.Scrollbar:
        # Replace the scrollbar creation with this:
        """
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        self.status = tk.Text(
            status_frame,
            height=12,
            width=140,
            font=('Consolas', 9),
            bg='#1e2124',
            fg='#dcddde',
            insertbackground='#8ecae6',
            selectbackground='#36393f',
            selectforeground='#dcddde',
            wrap=tk.WORD,
            relief='flat',
            bd=0,
            yscrollcommand=scrollbar.set
        )
        self.status.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.status.yview)
        """

        # Initial welcome message
        self.log("🚀 Welcome to Email Report Generator!")
        self.log("📝 Select a report type and configure the required files to get started.")

    def log(self, msg):
        """Add message to status box with timestamp"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] {msg}\n"

        self.status.insert(tk.END, formatted_msg)
        self.status.see(tk.END)
        self.update()

    def send_timesheet_single(self):
        """Send timesheet report for single user"""
        user = self.user.get().strip()
        if not user or not self.timesheet.get():
            self.log("❌ User and timesheet data required.")
            return

        weeks = self._weeks()
        if not weeks:
            self.log("❌ No valid weeks specified")
            return

        try:
            if not os.path.exists(self.timesheet.get()):
                self.log("❌ Timesheet file does not exist")
                return

            self.log(f"📊 Processing timesheet report for {user}...")

            # Pass user to analyze_timesheet_missing
            timesheet_data = analyze_timesheet_missing(self.timesheet.get(), weeks[-1], user)
            if timesheet_data is None:
                self.log("ℹ️ No timesheet missing data found exceeding 35 minutes")
                return

            timesheet_table = html_timesheet_missing_table(timesheet_data)
            if not timesheet_table:
                self.log("ℹ️ No data to display in timesheet report")
                return

            html = f"""
            <html>
            <body style="margin:0;padding:20px;background:#f9f9fb;font-family:Segoe UI,Arial,sans-serif;color:#333;line-height:1.4;">
                <h2 style="font-size:18px;margin-bottom:10px;font-weight:normal;">Timesheet Missing Report</h2>
                <p style="margin-top:0;font-size:13px;">Hi {user}, here is your timesheet missing report:</p>
                {timesheet_table}
            </body>
            </html>
            """

            send_mail_html(f"{user}@amazon.com", "", "Timesheet Missing Report", html,
                           preview=self.mode.get() == "preview")

            mode_text = "previewed" if self.mode.get() == "preview" else "sent"
            self.log(f"✅ Timesheet mail for {user} {mode_text} successfully!")
            return

        except Exception as exc:
            self.log(f"❌ Error generating timesheet report: {str(exc)}")
            return

    def send_single(self):
        """Handle single report sending based on selected type"""
        report_type = self.report_type.get()
        self.log(f"🚀 Starting {report_type} (Single)...")

        if report_type == "RDR Report":
            self.send_rdr_single()
        elif report_type == "Precision Correction Report":
            self.send_precision_single()
        elif report_type == "Timesheet Missing Report":
            self.send_timesheet_single()
        else:
            self.log("❌ Invalid report type selected")
            return

    def send_bulk(self):
        """Handle bulk report sending based on selected type"""
        report_type = self.report_type.get()
        self.log(f"🚀 Starting {report_type} (Bulk)...")

        if report_type == "RDR Report":
            self.send_rdr_bulk()
        elif report_type == "Precision Correction Report":
            self.send_precision_bulk()
        elif report_type == "Timesheet Missing Report":
            self.send_timesheet_bulk()
        else:
            self.log("❌ Invalid report type selected")

    def send_rdr_single(self):
        user = self.user.get().strip()
        if not user or not self.data.get():
            return self.log("❌ User and data required for RDR Report.")
        weeks = self._weeks()
        df = load_productivity(self.data.get())
        prod = user_prod_dict(df, user, weeks)
        allprod = all_prod_dict(df, weeks)
        prod_pct = productivity_percentiles(allprod, weeks)
        prod_table = html_metric_value_table_with_latest(prod, weeks, prod_pct, section="Productivity",
                                                         is_quality=False)
        prod_pct_table = html_metric_pct_table(prod, weeks, prod_pct, section="Productivity")
        qual_table = qual_pct_table = qc2_left = qc2_right = ""
        correction_quality_table = ""
        correction_count_table = ""

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

            correction_data = calculate_correction_type_data(qdf, user, weeks)
            correction_quality_table = html_correction_type_quality_table(correction_data, weeks)
            correction_count_table = html_correction_type_count_table(correction_data, weeks)
        # In send_rdr_single method:
        timesheet_table = ""
        if self.timesheet.get() and os.path.exists(self.timesheet.get()):
            timesheet_data = analyze_timesheet_missing(self.timesheet.get(), weeks[-1])  # analyze last week
            timesheet_table = html_timesheet_missing_table(timesheet_data)

        html = compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right,
                            correction_quality_table, correction_count_table,
                            timesheet_table)  # Add timesheet_table here
        send_mail_html(f"{user}@amazon.com", "", "RDR Productivity & Quality Metrics Report", html,
                       preview=self.mode.get() == "preview")
        self.log(f"RDR mail for {user} {'previewed' if self.mode.get() == 'preview' else 'sent'}")

    def send_rdr_bulk(self):
        if not self.data.get() or not self.mgr.get():
            return self.log("❌ Data and manager map required for RDR Bulk Report.")
        weeks = self._weeks()
        mgr = pd.read_excel(self.mgr.get())
        df = load_productivity(self.data.get())
        allprod = all_prod_dict(df, weeks)
        prod_pct = productivity_percentiles(allprod, weeks)
        allqual = qual_pct = {}
        qdf = None

        if self.qual.get() and os.path.exists(self.qual.get()):
            qdf = load_quality(self.qual.get())
            allqual = all_quality_dict(qdf, weeks)
            qual_pct = quality_percentiles(allqual, weeks)

        for user in mgr['loginname'].dropna().unique():
            prod = allprod.get(user, {})
            prod_table = html_metric_value_table_with_latest(prod, weeks, prod_pct, section="Productivity",
                                                             is_quality=False)
            prod_pct_table = html_metric_pct_table(prod, weeks, prod_pct, section="Productivity")
            qual = allqual.get(user, {}) if allqual else ""
            qual_table = qual_pct_table = qc2_left = qc2_right = ""
            correction_quality_table = ""
            correction_count_table = ""

            if allqual:
                qual_table = html_metric_value_table_with_latest(qual, weeks, qual_pct, section="Quality",
                                                                 is_quality=True)
                qual_pct_table = html_metric_pct_table(qual, weeks, qual_pct, section="Quality")
                subreason = qc2_subreason_analysis(qdf, user, weeks)
                qc2_left = html_qc2_reason_value_table(subreason, weeks)
                qc2_right = html_qc2_reason_pct_table(subreason, weeks)

                correction_data = calculate_correction_type_data(qdf, user, weeks)
                correction_quality_table = html_correction_type_quality_table(correction_data, weeks)
                correction_count_table = html_correction_type_count_table(correction_data, weeks)
            # In send_rdr_single method:
            timesheet_table = ""
            if self.timesheet.get() and os.path.exists(self.timesheet.get()):
                timesheet_data = analyze_timesheet_missing(self.timesheet.get(), weeks[-1])  # analyze last week
                timesheet_table = html_timesheet_missing_table(timesheet_data)

            html = compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right,
                                correction_quality_table, correction_count_table,
                                timesheet_table)  # Add timesheet_table here
            html = compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right,
                                correction_quality_table, correction_count_table)
            cc = ""
            if 'supervisorloginname' in mgr.columns:
                cc_val = mgr[mgr['loginname'] == user]['supervisorloginname'].iloc[0]
                if pd.notna(cc_val):
                    cc = f"{cc_val}@amazon.com"
            send_mail_html(f"{user}@amazon.com", cc, "RDR Productivity & Quality Metrics Report", html,
                           preview=self.mode.get() == "preview")
            self.log(f"Mail for {user} {'previewed' if self.mode.get() == 'preview' else 'sent'}")

    def send_precision_single(self):
        user = self.user.get().strip()
        if not user or not self.data.get():
            return self.log("User and data required.")
        weeks = self._weeks()
        df = load_productivity(self.data.get())
        precision_data = calculate_precision_corrections(df, weeks)

        if user in precision_data:
            precision_table = html_precision_correction_table({user: precision_data[user]}, weeks)
            html = f"""
            <html>
            <body style="margin:0;padding:20px;background:#f9f9fb;font-family:Segoe UI,Arial,sans-serif;color:#333;line-height:1.4;">
                <h2 style="font-size:18px;margin-bottom:10px;font-weight:normal;">High Precision Correction Alert</h2>
                <p style="margin-top:0;font-size:13px;">Hi {user}, you have high precision corrections:</p>
                {precision_table}
            </body>
            </html>
            """

            send_mail_html(f"{user}@amazon.com", "", "High Precision Corrections Alert", html,
                           preview=self.mode.get() == "preview")
            self.log(f"Precision mail for {user} {'previewed' if self.mode.get() == 'preview' else 'sent'}")

    def send_precision_bulk(self):
        if not self.data.get():
            return self.log("Data required.")
        weeks = self._weeks()
        df = load_productivity(self.data.get())
        precision_data = calculate_precision_corrections(df, weeks)

        for user in precision_data:
            precision_table = html_precision_correction_table({user: precision_data[user]}, weeks)
            html = f"""
            <html>
            <body style="margin:0;padding:20px;background:#f9f9fb;font-family:Segoe UI,Arial,sans-serif;color:#333;line-height:1.4;">
                <h2 style="font-size:18px;margin-bottom:10px;font-weight:normal;">High Precision Correction Alert</h2>
                <p style="margin-top:0;font-size:13px;">Hi {user}, you have high precision corrections:</p>
                {precision_table}
            </body>
            </html>
            """

            send_mail_html(f"{user}@amazon.com", "", "High Precision Corrections Alert", html,
                           preview=self.mode.get() == "preview")
            self.log(f"Precision mail for {user} {'previewed' if self.mode.get() == 'preview' else 'sent'}")

    def send_timesheet_bulk(self):
        """Send timesheet reports in bulk to multiple users"""
        # Check required inputs
        if not self.timesheet.get() or not self.mgr.get():
            self.log("❌ Timesheet data and manager map required for Timesheet Bulk Report.")
            return

        # Get weeks
        weeks = self._weeks()
        if not weeks:
            self.log("❌ No valid weeks specified")
            return

        try:
            # Read manager mapping file
            mgr = pd.read_excel(self.mgr.get())
            self.log(f"⏰ Processing timesheet reports for {len(mgr)} users...")

            # Process each user
            for user in mgr['loginname'].dropna().unique():
                try:
                    # Get timesheet data for user
                    timesheet_data = analyze_timesheet_missing(self.timesheet.get(), weeks[-1],
                                                               user)  # Added user parameter
                    if timesheet_data is None or timesheet_data.empty:
                        self.log(f"ℹ️ No timesheet data found for {user}")
                        continue

                    # Generate HTML report
                    timesheet_table = html_timesheet_missing_table(timesheet_data)
                    html = f"""
                    <html>
                    <body style="margin:0;padding:20px;background:#f9f9fb;font-family:Segoe UI,Arial,sans-serif;color:#333;line-height:1.4;">
                        <h2 style="font-size:18px;margin-bottom:10px;font-weight:normal;">Timesheet Missing Report</h2>
                        <p style="margin-top:0;font-size:13px;">Hi {user}, here is your timesheet missing report:</p>
                        {timesheet_table}
                    </body>
                    </html>
                    """

                    # Get CC email if supervisor exists
                    cc = ""
                    if 'supervisorloginname' in mgr.columns:
                        cc_val = mgr[mgr['loginname'] == user]['supervisorloginname'].iloc[0]
                        if pd.notna(cc_val):
                            cc = f"{cc_val}@amazon.com"

                    # Send email
                    send_mail_html(
                        to=f"{user}@amazon.com",
                        cc=cc,
                        subject="Timesheet Missing Report",
                        html=html,
                        preview=self.mode.get() == "preview"
                    )

                    # Log success
                    mode_text = "previewed" if self.mode.get() == "preview" else "sent"
                    self.log(f"✅ Timesheet mail for {user} {mode_text}")

                except Exception as user_exc:
                    # Log individual user errors but continue processing
                    self.log(f"⚠️ Error processing user {user}: {str(user_exc)}")
                    continue

            # Log completion
            self.log("🎉 Bulk timesheet processing completed!")
            return

        except Exception as exc:
            # Log critical errors
            self.log(f"❌ Error processing bulk timesheet: {str(exc)}")
            return

if __name__ == "__main__":
    ReporterApp().mainloop()
