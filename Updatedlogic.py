import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import numpy as np
import win32com.client as win32
import os
import requests
import json
from urllib.parse import urljoin
from tkinter import messagebox

class ETLDataFetcher:
    def __init__(self):
        self.base_url = "https://datacentral.a2z.com/dw-platform/servlet/dwp/template/"
        self.session = requests.Session()

        # Initialize authentication
        self.authenticate_midway()

    def authenticate_midway(self):
        """Open Midway authentication in default browser"""
        import webbrowser

        midway_url = "https://midway-auth.amazon.com/login"
        webbrowser.open(midway_url)

        # Don't use input() here - the GUI handles the confirmation

    def get_etl_data(self, profile_id, days_back=30):
        try:
            local_file = f"etl_dump_{profile_id}.xlsx"
            if os.path.exists(local_file):
                return pd.read_excel(local_file)

            url = urljoin(self.base_url, f"EtlViewExtractJobs.vm/job_profile_id/{profile_id}")
            response = self.session.get(url)
            print("response", response)

            if response.status_code == 401:
                raise Exception("Authentication failed or session expired")

            response.raise_for_status()
            return pd.DataFrame(response.json())

        except Exception as e:
            raise Exception(f"ETL data fetch failed: {str(e)}")
def validate_excel_file(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        required_sheets = ['Volume', 'Quality', 'TimeSheet', 'Manager']
        missing_sheets = [sheet for sheet in required_sheets if sheet not in xl.sheet_names]
        if missing_sheets:
            return False, f"Missing required sheets: {', '.join(missing_sheets)}"
        return True, "All required sheets present"
    except Exception as e:
        return False, f"Error validating Excel file: {str(e)}"

def load_productivity_from_etl():
    try:
        etl = ETLDataFetcher()
        df = etl.get_etl_data("13404076")  # Productivity profile ID

        required_cols = ['useralias', 'pipeline', 'p_week',
                         'processed_volume', 'processed_volumetr',
                         'processed_time', 'processed_time_tr',
                         'precision_correction','acl']

        df = df[required_cols].astype({
            'useralias': 'string',
            'pipeline': 'string',
            'p_week': 'int32',
            'processed_volume': 'float32',
            'processed_volumetr': 'float32',
            'processed_time': 'float32',
            'processed_time_tr': 'float32',
            'precision_correction': 'float32',
            'acl': 'string'
        })

        return df

    except Exception as e:
        raise Exception(f"Failed to load productivity data: {str(e)}")

def load_timesheet_from_etl():
    try:
        etl = ETLDataFetcher()
        df = etl.get_etl_data("13417956")  # Timesheet profile ID

        required_cols = ['work_date', 'week', 'timesheet_missing', 'loginid','acl']
        df = df[required_cols]
        df['work_date'] = pd.to_datetime(df['work_date'])

        return df

    except Exception as e:
        raise Exception(f"Failed to load timesheet data: {str(e)}")

# ----------------- Data Processing -------------------
def load_productivity(data_path):
    try:
        cols = ['useralias', 'pipeline', 'p_week', 'processed_volume', 'processed_volumetr',
                'processed_time', 'processed_time_tr', 'precision_correction']
        dtype = {'useralias': 'string', 'pipeline': 'string', 'p_week': 'int32',
                'processed_volume': 'float32', 'processed_volumetr': 'float32',
                'processed_time': 'float32', 'processed_time_tr': 'float32',
                'precision_correction': 'float32'}
        df = pd.read_excel(data_path, sheet_name='Volume', usecols=cols, dtype=dtype)
        for c in ['processed_volume', 'processed_volumetr', 'processed_time', 'processed_time_tr', 'precision_correction']:
            df[c] = df[c].fillna(0)
    except ValueError:
        cols = ['useralias', 'p_week', 'precision_correction']
        df = pd.read_excel(data_path, sheet_name='Volume', usecols=cols)
        df['precision_correction'] = df['precision_correction'].fillna(0)
    return df


def get_common_styles():
    """Common styles for all tables to ensure consistency"""
    return {
        'container_style': '''
            margin: 10px 0;
            font-family: Arial, sans-serif;
        ''',
        'table_style': '''
            border-collapse: separate;
            border-spacing: 0;
            width: 100%;
            border: 3px solid white;
            border-top: 3px solid white;  # Added this line to ensure top border is white
            font-family: Arial, sans-serif;
        ''',
        'title_style': '''
            background-color: #000000;
            color: white;
            text-align: center;
            font-size: 14px;
            font-weight: bold;
            padding: 12px;
            border: none;
        ''',
        'header_style': '''
            padding: 8px 6px;
            border: 1px solid white;
            background-color: #000000;
            color: white;
            text-align: center;
            font-size: 13px;
            font-weight: bold;
            vertical-align: middle;
        ''',
        'cell_style': '''
            padding: 8px 6px;
            border: 1px solid white;
            text-align: center;
            font-size: 12px;
            background-color: #ffffff;
            color: #000000;
            height: 32px;
            vertical-align: middle;
            word-wrap: break-word;
            overflow: hidden;
        ''',
        'pipeline_cell_style': '''
            padding: 8px 12px;
            border: 1px solid white;
            background-color: #ffffff;
            font-size: 13px;
            text-align: left;
            font-weight: bold;
            color: #000000;
            vertical-align: middle;
        '''
    }

def calculate_precision_corrections(df, weeks=None):
    if weeks:
        df = df[df['p_week'].isin(weeks)]
    precision_data = df.groupby(['useralias', 'p_week', 'acl'])['precision_correction'].sum().reset_index()
    high_precision = precision_data[precision_data['precision_correction'] > 20]
    result = {}
    for _, row in high_precision.iterrows():
        result.setdefault(row['useralias'], {}).setdefault(row['p_week'], {
            'value': row['precision_correction'],
            'acl': row['acl']
        })
    return result


def load_quality(qual_path):
    try:
        # Read the Quality sheet
        df = pd.read_excel(qual_path, sheet_name='Quality')
        print("\nDebugging load_quality:")
        print(f"Raw quality data shape: {df.shape}")
        print(f"All available columns: {df.columns.tolist()}")

        # Check if the column name might be different
        possible_correction_columns = [col for col in df.columns if 'correction' in col.lower()]
        print(f"Possible correction-related columns: {possible_correction_columns}")

        # Define required columns with correction type
        required_cols = ['volume', 'auditor_login', 'program', 'usecase',
                         'qc2_judgement', 'qc2_subreason', 'week',
                         'auditor_correction_type']  # Make sure this matches exact column name

        # Verify columns exist
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            print(f"Missing columns: {missing_cols}")
            # Check for case-insensitive matches
            for missing_col in missing_cols:
                matches = [col for col in df.columns if col.lower() == missing_col.lower()]
                if matches:
                    print(f"Found possible match for {missing_col}: {matches}")

        # Select and process columns
        try:
            df = df[required_cols].copy()
            df['volume'] = df['volume'].fillna(0)
            df = df[df['program'] == 'RDR']
            print(f"Processed quality data shape: {df.shape}")
            print(f"Sample processed data:\n{df.head()}")
            return df

        except KeyError as ke:
            print(f"KeyError when selecting columns: {ke}")
            print("Available columns:", df.columns.tolist())
            return pd.DataFrame()

    except Exception as e:
        print(f"Error loading quality data: {str(e)}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def user_prod_dict(df, user, weeks=None):
    dfu = df[df['useralias'] == user]
    if weeks: dfu = dfu[dfu['p_week'].isin(weeks)]
    dfu['total_volume'] = dfu['processed_volume'] + dfu['processed_volumetr']
    dfu['total_time'] = dfu['processed_time'] + dfu['processed_time_tr']
    g = dfu.groupby(['pipeline', 'p_week'], as_index=False)[['total_volume','total_time']].sum()
    g['productivity'] = np.where(g['total_time'] > 0, g['total_volume']/g['total_time'], 0)
    return g.pivot(index='pipeline', columns='p_week', values='productivity').fillna(0).to_dict('index')

def all_prod_dict(df, weeks=None):
    if weeks: df = df[df['p_week'].isin(weeks)]
    df['total_volume'] = df['processed_volume'] + df['processed_volumetr']
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
    try:
        dfu = df[df['auditor_login'] == user]
        print(f"Quality data for user {user}:", len(dfu))  # Debug print

        if weeks:
            dfu = dfu[dfu['week'].isin(weeks)]
            print(f"Quality data for weeks {weeks}:", len(dfu))  # Debug print

        d = {}
        for (pipe, week), grp in dfu.groupby(['usecase', 'week']):
            total_volume = grp['volume'].sum()
            error_volume = grp[grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT'])]['volume'].sum()
            score = (total_volume - error_volume) / total_volume if total_volume else 0
            print(f"Pipe: {pipe}, Week: {week}, Total: {total_volume}, Errors: {error_volume}")  # Debug print
            d.setdefault(pipe, {})[week] = {'score': score, 'err': error_volume, 'total': total_volume}
        return d

    except Exception as e:
        print(f"Error in user_quality_dict: {str(e)}")
        return {}

def all_quality_dict(df, weeks=None):
    if weeks:
        df = df[df['week'].isin(weeks)]
    d = {}
    for (aud, pipe, week), grp in df.groupby(['auditor_login','usecase','week']):
        total_volume = grp['volume'].sum()
        error_volume = grp[grp['qc2_judgement'].isin(['AUDITOR_INCORRECT','BOTH_INCORRECT'])]['volume'].sum()
        score = (total_volume-error_volume)/total_volume if total_volume else 0
        d.setdefault(aud, {}).setdefault(pipe, {})[week] = {'score': score, 'err': error_volume, 'total': total_volume}
    return d

def analyze_timesheet_missing(timesheet_path, week, user):
    try:
        df = pd.read_excel(timesheet_path, sheet_name='TimeSheet')
        required_columns = ['work_date', 'week', 'timesheet_missing', 'loginid']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")
        weekly_data = df[(df['week'] == week) & (df['loginid'] == user)]
        weekly_data['work_date'] = pd.to_datetime(weekly_data['work_date'])
        daily_missing = weekly_data[weekly_data['timesheet_missing'] > 35][['work_date', 'timesheet_missing']]
        return daily_missing.sort_values('work_date') if not daily_missing.empty else None
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
    dfu = df[df['auditor_login'] == user]
    if weeks: dfu = dfu[dfu['week'].isin(weeks)]
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
        return "#006400", "white"
    elif bench == "P75-P90":
        return "#90EE90", "black"
    elif bench == "P50-P75":
        return "#FFD700", "black"
    elif bench == "P30-P50":
        return "#ffc107", "black"
    elif bench == "<P30":
        return "#dc3545", "white"
    else:
        return "#ffffff", "black"


def html_metric_value_table_with_latest(self, data, weeks, percentiles, section="Quality", is_quality=False,
                                        month_data=None):
    from datetime import datetime
    import calendar

    styles = get_common_styles()
    if not weeks:
        weeks = []
    styles = get_common_styles()
    # Get the month name based on current/previous selection
    current_date = datetime.now()
    if self.month_selection.get() == "current":
        month_name = calendar.month_name[current_date.month]
    else:
        # For previous month
        previous_month = 12 if current_date.month == 1 else current_date.month - 1
        month_name = calendar.month_name[previous_month]

    # Calculate column widths
    num_weeks = len(weeks)
    pipeline_width = 25
    week_width = (60 - 15) / num_weeks
    latest_width = 15
    month_width = 15

    # Create header cells
    ths = "".join(
        f"<th style='{styles['header_style']} width: {week_width}%;'>{w}</th>"
        for w in weeks
    )
    ths += f"<th style='{styles['header_style']} width: {latest_width}%;'>Latest</th>"
    ths += f"<th style='{styles['header_style']} width: {month_width}%;'>Month({month_name})</th>"

    # Create rows with monthly data
    rows = ""
    for pipe in sorted(data):
        tds = ""
        for w in weeks:
            if is_quality:
                val_dict = data[pipe].get(w, None)
                val = val_dict['score'] if isinstance(val_dict, dict) and 'score' in val_dict else 0
                err = val_dict['err'] if isinstance(val_dict, dict) and 'err' in val_dict else 0
                total = val_dict['total'] if isinstance(val_dict, dict) and 'total' in val_dict else 0
                disp = f"{val:.1%}<br>({err}/{total})" if total else "-"
            else:
                val = data[pipe].get(w, 0)
                disp = f"{val:.1f}" if val else "-"

            tds += f"<td style='{styles['cell_style']} width: {week_width}%;'>{disp}</td>"

        # Add latest benchmark
        val = data[pipe].get(weeks[-1], 0)
        val_scalar = val['score'] if is_quality and isinstance(val, dict) and 'score' in val else val
        pctls = percentiles[weeks[-1]].get(pipe, None) if percentiles and weeks[-1] in percentiles else None
        bench = percentile_label(val_scalar, pctls)
        bg_color, text_color = pct_bg_fg(bench)

        benchmark_style = f'''
            padding: 8px 6px;
            border: 1px solid white;
            text-align: center;
            font-size: 12px;
            background-color: {bg_color};
            color: {text_color};
            height: 32px;
            vertical-align: middle;
            word-wrap: break-word;
            overflow: hidden;
            width: {latest_width}%;
        '''

        tds += f"<td style='{benchmark_style}'>{bench if bench != '-' else '-'}</td>"

        # Add month data
        if month_data and pipe in month_data:
            month_val = month_data[pipe]
            month_disp = f"{month_val:.1f}" if not is_quality else f"{month_val:.1%}"
        else:
            month_disp = "-"

        tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{month_disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'>{pipe}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 3}" style="{styles['title_style']}">
                    Weekly {section} Metrics
                </td>
            </tr>
            <tr>
                <th style="{styles['header_style']} width: {pipeline_width}%;">Pipeline</th>
                {ths}
            </tr>
            {rows}
        </table>
    </div>
    """


def html_precision_correction_table(precision_data, weeks):
    styles = get_common_styles()

    if not precision_data:
        return f"""<div style='padding:15px;background-color:#fef3c7;border:3px solid black;color:#92400e;font-weight:bold;width:400px;'>
                   No precision corrections found exceeding 20.</div>"""

    rows = ""
    for user in precision_data:
        for week in weeks:
            if week in precision_data[user]:
                week_data = precision_data[user][week]
                acl = week_data['acl']
                value = week_data['value']
                rows += f"""<tr>
                    <td style='{styles['pipeline_cell_style']} width: 150px; border: 1px solid black;'>{week}</td>
                    <td style='{styles['pipeline_cell_style']} width: 150px; border: 1px solid black;'>{acl}</td>
                    <td style='{styles['cell_style']} width: 100px; border: 1px solid black; color: #ff6b6b;'>{value:.0f}</td>
                    </tr>"""

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']} width: 400px; border: 3px solid black;">
            <tr>
                <td colspan="3" style="{styles['title_style']} background-color: #0d47a1;">
                    Precision Corrections Report
                </td>
            </tr>
            <tr>
                <th style="{styles['header_style']} width: 150px; border: 1px solid black; background-color: #0d47a1;">Week</th>
                <th style="{styles['header_style']} width: 150px; border: 1px solid black; background-color: #0d47a1;">ACL</th>
                <th style="{styles['header_style']} width: 100px; border: 1px solid black; background-color: #0d47a1;">Corrections</th>
            </tr>
            {rows}
        </table>
    </div>
    """


def html_timesheet_missing_table(timesheet_data):
    styles = get_common_styles()

    if timesheet_data is None or timesheet_data.empty:
        return f"""<div style='padding:15px;background-color:#fef3c7;border:3px solid black;color:#92400e;font-weight:bold;width:400px;'>
                   No timesheet missing data found exceeding 35 minutes.</div>"""

    rows = ""
    for _, row in timesheet_data.iterrows():
        work_date = row['work_date'].strftime('%Y-%m-%d')
        missing = f"{row['timesheet_missing']:.2f}"
        rows += f"""<tr>
                    <td style='{styles['pipeline_cell_style']} width: 200px; border: 1px solid black;'>{work_date}</td>
                    <td style='{styles['cell_style']} width: 200px; border: 1px solid black; color: #ff6b6b;'>{missing}</td>
                    </tr>"""

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']} width: 400px; border: 3px solid black;">
            <tr>
                <td colspan="2" style="{styles['title_style']} background-color: #0d47a1;">
                    Timesheet Missing Report (>35 minutes)
                </td>
            </tr>
            <tr>
                <th style="{styles['header_style']} width: 200px; border: 1px solid black; background-color: #0d47a1;">Date</th>
                <th style="{styles['header_style']} width: 200px; border: 1px solid black; background-color: #0d47a1;">Missing Minutes</th>
            </tr>
            {rows}
        </table>
    </div>
    """


def html_metric_pct_table(data, weeks, percentiles, section="Quality"):
    styles = get_common_styles()

    # Calculate column widths
    num_weeks = len(weeks)
    pipeline_width = 25
    week_width = 75 / num_weeks

    ths = "".join(
        f"<th style='{styles['header_style']} width: {week_width}%;'>{w}</th>"
        for w in weeks
    )

    rows = ""
    for pipe in sorted(data):
        tds = ""
        for w in weeks:
            val = data[pipe].get(w, 0)
            val_scalar = val['score'] if isinstance(val, dict) and 'score' in val else val
            pctls = percentiles[w].get(pipe, None) if percentiles and w in percentiles else None
            bench = percentile_label(val_scalar, pctls) or "-"
            bg_color, text_color = pct_bg_fg(bench)

            cell_style = f'''
                padding: 8px 6px;
                border: 1px solid white;
                border-top: 1px solid white;
                text-align: center;
                font-size: 12px;
                background-color: {bg_color};
                color: {text_color};
                height: 32px;
                vertical-align: middle;
                word-wrap: break-word;
                overflow: hidden;
                width: {week_width}%;
            '''

            tds += f"<td style='{cell_style}'>{bench}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'>{pipe}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 1}" style="{styles['title_style']}">
                    Weekly {section} Benchmarks
                </td>
            </tr>
            <tr>
                <th style="{styles['header_style']} width: {pipeline_width}%;">Pipeline</th>
                {ths}
            </tr>
            {rows}
        </table>
    </div>
    """


def calculate_monthly_productivity_metrics(df, is_current=True):
    """Calculate monthly productivity metrics"""
    from datetime import datetime, timedelta
    import pandas as pd

    current_date = datetime.now()
    current_year = current_date.year

    # Create a date column based on week numbers
    df['date'] = pd.to_datetime(df['p_week'].astype(str) + str(current_year) + '1', format='%W%Y%w')

    if is_current:
        month_start = current_date.replace(day=1)
    else:
        first_of_month = current_date.replace(day=1)
        month_start = (first_of_month - timedelta(days=1)).replace(day=1)

    month_data = df[df['date'] >= month_start]
    if not is_current:
        next_month = (month_start + timedelta(days=32)).replace(day=1)
        month_data = month_data[month_data['date'] < next_month]

    # Calculate monthly metrics
    month_data['total_volume'] = month_data['processed_volume'] + month_data['processed_volumetr']
    month_data['total_time'] = month_data['processed_time'] + month_data['processed_time_tr']

    monthly_metrics = month_data.groupby('pipeline').agg({
        'total_volume': 'sum',
        'total_time': 'sum'
    }).reset_index()

    monthly_metrics['productivity'] = monthly_metrics['total_volume'] / monthly_metrics['total_time']

    return monthly_metrics.set_index('pipeline')['productivity'].to_dict()


def calculate_monthly_quality_metrics(df, is_current=True):
    """Calculate monthly quality metrics"""
    from datetime import datetime, timedelta
    import pandas as pd

    current_date = datetime.now()
    current_year = current_date.year

    # Create a date column based on week numbers
    df['date'] = pd.to_datetime(df['week'].astype(str) + str(current_year) + '1', format='%W%Y%w')

    if is_current:
        month_start = current_date.replace(day=1)
    else:
        first_of_month = current_date.replace(day=1)
        month_start = (first_of_month - timedelta(days=1)).replace(day=1)

    month_data = df[df['date'] >= month_start]
    if not is_current:
        next_month = (month_start + timedelta(days=32)).replace(day=1)
        month_data = month_data[month_data['date'] < next_month]

    # Calculate quality metrics by usecase
    monthly_metrics = {}
    for (pipe, grp) in month_data.groupby('usecase'):
        total_volume = grp['volume'].sum()
        error_volume = grp[grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT'])]['volume'].sum()
        score = (total_volume - error_volume) / total_volume if total_volume else 0
        monthly_metrics[pipe] = score

    return monthly_metrics


def calculate_monthly_subreason_metrics(df, is_current=True):
    """Calculate monthly subreason metrics"""
    from datetime import datetime, timedelta

    current_date = datetime.now()
    if is_current:
        month_start = current_date.replace(day=1)
    else:
        first_of_month = current_date.replace(day=1)
        month_start = (first_of_month - timedelta(days=1)).replace(day=1)

    df['date'] = pd.to_datetime(df['week'].astype(str) + str(current_date.year) + '1', format='%W%Y%w')
    month_data = df[df['date'] >= month_start]

    if not is_current:
        next_month = (month_start + timedelta(days=32)).replace(day=1)
        month_data = month_data[month_data['date'] < next_month]

    monthly_metrics = {}
    for (usecase, sub_reason), grp in month_data.groupby(['usecase', 'qc2_subreason']):
        total = len(grp)
        monthly_metrics.setdefault(usecase, {})[sub_reason] = {
            'count': total,
            'percentage': total / len(month_data) * 100 if len(month_data) > 0 else 0
        }

    return monthly_metrics


def calculate_monthly_correction_metrics(df, is_current=True):
    """Calculate monthly correction type metrics"""
    from datetime import datetime, timedelta

    current_date = datetime.now()
    if is_current:
        month_start = current_date.replace(day=1)
    else:
        first_of_month = current_date.replace(day=1)
        month_start = (first_of_month - timedelta(days=1)).replace(day=1)

    df['date'] = pd.to_datetime(df['week'].astype(str) + str(current_date.year) + '1', format='%W%Y%w')
    month_data = df[df['date'] >= month_start]

    if not is_current:
        next_month = (month_start + timedelta(days=32)).replace(day=1)
        month_data = month_data[month_data['date'] < next_month]

    monthly_metrics = {}
    for correction_type, grp in month_data.groupby('auditor_correction_type'):
        total = len(grp)
        errors = len(grp[grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT'])])
        monthly_metrics[correction_type] = {
            'total': total,
            'errors': errors,
            'score': (total - errors) / total if total > 0 else 0
        }

    return monthly_metrics
def calculate_correction_type_data(df, user, weeks=None):
    try:
        # Debug prints
        print(f"\nDebugging calculate_correction_type_data:")
        print(f"Available columns: {df.columns.tolist()}")
        print(f"Sample data:\n{df.head()}")

        dfu = df[df['auditor_login'] == user]
        print(f"Data for user {user}: {len(dfu)} rows")

        if weeks:
            dfu = dfu[dfu['week'].isin(weeks)]
            print(f"Data after week filter: {len(dfu)} rows")

        # Check for correction type data
        if 'auditor_correction_type' not in dfu.columns:
            print("Missing auditor_correction_type column")
            return {}

        # Show unique correction types
        print(f"Unique correction types: {dfu['auditor_correction_type'].unique()}")

        correction_data = {}
        for (correction_type, week), grp in dfu.groupby(['auditor_correction_type', 'week']):
            total = len(grp)
            err = grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']).sum()
            score = (total - err) / total if total else 0
            print(f"Correction type: {correction_type}, Week: {week}, Total: {total}, Errors: {err}")
            correction_data.setdefault(correction_type, {})[week] = {
                'score': score,
                'err': err,
                'total': total
            }
        return correction_data

    except Exception as e:
        print(f"Error in calculate_correction_type_data: {str(e)}")
        import traceback
        traceback.print_exc()
        return {}


def html_qc2_reason_value_table(self, subreason_data, weeks, month_data=None):
    styles = get_common_styles()
    if not weeks:
        weeks = []
    if not subreason_data:
        subreason_data = {}
    # Get month name
    current_date = datetime.now()
    if self.month_selection.get() == "current":
        month_name = calendar.month_name[current_date.month]
    else:
        previous_month = 12 if current_date.month == 1 else current_date.month - 1
        month_name = calendar.month_name[previous_month]

    # Calculate column widths
    num_weeks = len(weeks)
    subreason_width = 30
    week_width = (55 - 15) / num_weeks
    month_width = 15

    # Create header
    ths = "".join(
        f"<th style='{styles['header_style']} width: {week_width}%;'>{w}</th>"
        for w in weeks
    )
    ths += f"<th style='{styles['header_style']} width: {month_width}%;'>Month({month_name})</th>"

    # Create rows with monthly data
    rows = ""
    all_subs = sorted({sub for pipe in subreason_data for w in subreason_data[pipe]
                       for sub in subreason_data[pipe][w]['counts']})

    for sub in all_subs:
        tds = ""
        for w in weeks:
            count = sum(subreason_data[pipe][w]['counts'].get(sub, 0)
                        for pipe in subreason_data if w in subreason_data[pipe])
            tds += f"<td style='{styles['cell_style']} width: {week_width}%;'>{count if count else '-'}</td>"

        # Add monthly data
        if month_data:
            month_count = sum(pipe_data.get(sub, {}).get('count', 0)
                              for pipe_data in month_data.values())
            tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{month_count if month_count else '-'}</td>"
        else:
            tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>-</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {subreason_width}%;'>{sub}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr><td colspan="{len(weeks) + 2}" style="{styles['title_style']}">QC2 Subreason Counts</td></tr>
            <tr>
                <th style="{styles['header_style']} width: {subreason_width}%;">Subreason</th>
                {ths}
            </tr>
            {rows}
        </table>
    </div>
    """

def html_qc2_reason_pct_table(subreason_data, weeks):
    styles = get_common_styles()

    all_subs = sorted({sub for pipe in subreason_data for w in subreason_data[pipe] for sub in
                       subreason_data[pipe][w]['percentages']})

    # Calculate column widths
    num_weeks = len(weeks)
    subreason_width = 30
    week_width = 70 / num_weeks

    ths = "".join(
        f"<th style='{styles['header_style']} width: {week_width}%;'>{w}</th>"
        for w in weeks
    )

    rows = ""
    for sub in all_subs:
        tds = ""
        for w in weeks:
            perc = sum(subreason_data[pipe][w]['percentages'].get(sub, 0) for pipe in subreason_data if
                       w in subreason_data[pipe])
            disp = f"{perc:.1f}%" if perc else "-"

            tds += f"<td style='{styles['cell_style']} width: {week_width}%;'>{disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {subreason_width}%;'>{sub}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 1}" style="{styles['title_style']}">
                    QC2 Subreason Percentages
                </td>
            </tr>
            <tr>
                <th style="{styles['header_style']} width: {subreason_width}%;">Subreason</th>
                {ths}
            </tr>
            {rows}
        </table>
    </div>
    """


def html_correction_type_quality_table(self, correction_data, weeks, month_data=None):
    styles = get_common_styles()

    # Get month name
    current_date = datetime.now()
    if self.month_selection.get() == "current":
        month_name = calendar.month_name[current_date.month]
    else:
        previous_month = 12 if current_date.month == 1 else current_date.month - 1
        month_name = calendar.month_name[previous_month]

    # Create header with monthly column
    ths = "".join(
        f"<th style='{styles['header_style']}'>{w}</th>"
        for w in weeks
    )
    ths += f"<th style='{styles['header_style']}'>Month({month_name})</th>"

    rows = ""
    for correction_type in sorted(correction_data):
        tds = ""
        for w in weeks:
            val = correction_data[correction_type].get(w, {})
            if val:
                score = val.get('score', 0)
                err = val.get('err', 0)
                total = val.get('total', 0)
                disp = f"{score:.1%}<br>({err}/{total})"
            else:
                disp = "-"
            tds += f"<td style='{styles['cell_style']}'>{disp}</td>"

        # Add monthly data
        if month_data and correction_type in month_data:
            month_val = month_data[correction_type]
            month_disp = f"{month_val['score']:.1%}<br>({month_val['errors']}/{month_val['total']})"
        else:
            month_disp = "-"
        tds += f"<td style='{styles['cell_style']}'>{month_disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']}'>{correction_type}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr><td colspan="{len(weeks) + 2}" style="{styles['title_style']}">Correction Type Quality</td></tr>
            <tr><th style="{styles['header_style']}">Type</th>{ths}</tr>
            {rows}
        </table>
    </div>
    """

def html_correction_type_count_table(correction_data, weeks):
    styles = get_common_styles()

    # Calculate column widths
    num_weeks = len(weeks)
    correction_width = 30
    week_width = 70 / num_weeks

    ths = "".join(
        f"<th style='{styles['header_style']} width: {week_width}%;'>{w}</th>"
        for w in weeks
    )

    rows = ""
    for correction_type in sorted(correction_data):
        tds = ""
        for w in weeks:
            val = correction_data[correction_type].get(w, {})
            count = val.get('total', 0)
            disp = str(count) if count else "-"

            tds += f"<td style='{styles['cell_style']} width: {week_width}%;'>{disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {correction_width}%;'>{correction_type}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 1}" style="{styles['title_style']}">
                    Weekly Correction Type Case Counts
                </td>
            </tr>
            <tr>
                <th style="{styles['header_style']} width: {correction_width}%;">Correction Type</th>
                {ths}
            </tr>
            {rows}
        </table>
    </div>
    """

def compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right,
                 correction_quality_table, correction_count_table, timesheet_table=""):
    return f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
        * {{
            box-sizing: border-box;
        }}
        @media (max-width: 768px) {{
            .table-container {{
                display: block !important;
                width: 100% !important;
            }}
            .table-container td {{
                display: block !important;
                width: 100% !important;
                padding: 0 !important;
                margin-bottom: 20px !important;
            }}
        }}
        .outer-table {{
            border: 3px solid white;
            border-collapse: separate !important;
            border-spacing: 0;
            background-color: #1f2937;
            table-layout: fixed;
            width: 100%;
        }}
        </style>
    </head>
    <body style="margin:0;padding:20px;background-color:#1f2937;font-family:Arial,sans-serif;color:#111827;">

        <table cellpadding="0" cellspacing="0" class="outer-table" style="margin-bottom:25px;">
            <tr>
                <td style="padding:20px;text-align:center;">
                    <h1 style="font-size:24px;margin:0;font-weight:bold;color:white;">RDR Productivity & Quality Metrics Report</h1>
                    <p style="margin:8px 0 0 0;font-size:16px;color:#d1d5db;">Hi {user}, here are your weekly performance insights</p>
                </td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table" style="margin-bottom:30px;">
            <tr>
                <td style="padding:12px;text-align:center;">
                    <h2 style="font-size:18px;margin:0;font-weight:bold;color:white;">📊 Productivity Metrics</h2>
                </td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table table-container" style="margin-bottom:30px;padding:20px;">
            <tr>
                <td style="width:48%;vertical-align:top;padding-right:10px;">{prod_table}</td>
                <td style="width:4%;"></td>
                <td style="width:48%;vertical-align:top;padding-left:10px;">{prod_pct_table}</td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table" style="margin-bottom:30px;">
            <tr>
                <td style="padding:12px;text-align:center;">
                    <h2 style="font-size:18px;margin:0;font-weight:bold;color:white;">⭐ Quality Metrics</h2>
                </td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table table-container" style="margin-bottom:30px;padding:20px;">
            <tr>
                <td style="width:48%;vertical-align:top;padding-right:10px;">{qual_table}</td>
                <td style="width:4%;"></td>
                <td style="width:48%;vertical-align:top;padding-left:10px;">{qual_pct_table}</td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table" style="margin-bottom:30px;">
            <tr>
                <td style="padding:12px;text-align:center;">
                    <h2 style="font-size:18px;margin:0;font-weight:bold;color:white;">🔍 QC2 Subreason Analysis</h2>
                </td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table table-container" style="margin-bottom:30px;padding:20px;">
            <tr>
                <td style="width:48%;vertical-align:top;padding-right:10px;">{qc2_left}</td>
                <td style="width:4%;"></td>
                <td style="width:48%;vertical-align:top;padding-left:10px;">{qc2_right}</td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table" style="margin-bottom:30px;">
            <tr>
                <td style="padding:12px;text-align:center;">
                    <h2 style="font-size:18px;margin:0;font-weight:bold;color:white;">🔧 Correction Type Analysis</h2>
                </td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table table-container" style="margin-bottom:30px;padding:20px;">
            <tr>
                <td style="width:48%;vertical-align:top;padding-right:10px;">{correction_count_table}</td>
                <td style="width:4%;"></td>
                <td style="width:48%;vertical-align:top;padding-left:10px;">{correction_quality_table}</td>
            </tr>
        </table>

        <div class="outer-table" style="padding:20px;margin-bottom:30px;">
            {timesheet_table}
        </div>

        <table cellpadding="0" cellspacing="0" class="outer-table" style="margin-top:30px;">
            <tr>
                <td style="padding:15px;">
                    <p style="font-size:12px;color:#d1d5db;margin:0;line-height:1.5;">
                        <strong style="color:#ffffff;">📈 Benchmarks Legend:</strong> "P90 +" (top 10%), "P75-P90", "P50-P75", "P30-P50", "&lt;P30"<br>
                        <strong style="color:#ffffff;">❓ Questions?</strong> Contact your manager for detailed insights and improvement strategies.
                    </p>
                </td>
            </tr>
        </table>

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
        self.title("🔍 Email Report Generator")
        self.geometry("1920x1200")
        self.state('zoomed')
        self.configure(bg='#121212')

        # Initialize variables
        self.data_source = tk.StringVar(value="excel")
        self.data = tk.StringVar()
        self.qual = tk.StringVar()
        self.mgr = tk.StringVar()
        self.timesheet = tk.StringVar()
        self.wks = tk.StringVar()
        self.user = tk.StringVar()
        self.mode = tk.StringVar(value="preview")
        self.month_selection = tk.StringVar(value="current")
        self.is_authenticated = False

        # Initialize data variables
        self.qdf = None
        self.subreason = None
        self.correction_data = None
        self.monthly_qual_metrics = None

        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('alt')

        # Dark Material Design color palette
        self.colors = {
            'primary': '#bb86fc',  # Purple 200 (Material Dark Primary)
            'primary_dark': '#985eff',  # Purple 300
            'primary_light': '#d0bcff',  # Purple 100
            'primary_variant': '#3700b3',  # Purple 700
            'secondary': '#03dac6',  # Teal 200 (Material Dark Secondary)
            'secondary_light': '#66fff0',  # Teal 100
            'secondary_dark': '#00a693',  # Teal 700
            'surface': '#1e1e1e',  # Dark surface
            'surface_variant': '#2d2d2d',  # Lighter dark surface
            'background': '#121212',  # Dark background
            'card_background': '#1e1e1e',  # Card background
            'on_surface': '#e1e1e1',  # Light text on dark surface
            'on_surface_variant': '#c1c1c1',  # Medium light text
            'on_surface_disabled': '#757575',  # Disabled text
            'success': '#4caf50',  # Green 500
            'error': '#cf6679',  # Red 300 (Material Dark Error)
            'warning': '#ffb74d',  # Orange 300
            'outline': '#404040',  # Border color
            'outline_variant': '#2d2d2d'  # Subtle border
        }

        # Configure ttk styles
        self._configure_styles()

        # Build UI
        self._build_ui()
        self.update_idletasks()
        self.after(100, self._force_refresh)

    def _force_refresh(self, event=None):
        """Force a refresh of the UI to ensure correct rendering"""
        try:
            # Store current geometry
            current_state = self.state()
            current_geometry = self.geometry()

            # Toggle state to force redraw
            if current_state == 'zoomed':
                self.state('normal')
                self.after(50, lambda: self.state('zoomed'))
            else:
                temp_geometry = f"{self.winfo_width() + 1}x{self.winfo_height() + 1}"
                self.geometry(temp_geometry)
                self.after(50, lambda: self.geometry(current_geometry))

            # Update all child widgets
            self.update_idletasks()

        except Exception as e:
            print(f"Error in _force_refresh: {str(e)}")
    def _on_resize(self, event=None):
        """Handle window resize events"""
        if event.widget == self:
            # Update layout if needed
            self.update_idletasks()

    def _handle_ui_update(self):
        """Handle UI updates after data changes"""
        try:
            self.update_idletasks()
            if hasattr(self, 'status') and self.status:
                self.status.see(tk.END)
        except Exception as e:
            print(f"Error updating UI: {str(e)}")

    def log(self, msg):
        """Add message to status box with timestamp"""
        try:
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_msg = f"[{timestamp}] {msg}\n"

            if hasattr(self, 'status') and self.status:
                self.status.insert(tk.END, formatted_msg)
                self.status.see(tk.END)
                self._handle_ui_update()
        except Exception as e:
            print(f"Error logging message: {str(e)}")

    def load_data(self):
        """Load data from either Excel or ETL based on selected source"""
        try:
            if self.data_source.get() == "etl":
                if not self.is_authenticated:
                    self.log("❌ Please authenticate with Midway first")
                    return None

                try:
                    etl = ETLDataFetcher()
                    return etl.get_etl_data("13404076")  # Productivity profile ID
                except Exception as e:
                    self.log(f"❌ ETL data fetch failed: {str(e)}")
                    return None

            else:  # Excel mode
                if not self.data.get():
                    self.log("❌ Please select a data file")
                    return None

                if not os.path.exists(self.data.get()):
                    self.log("❌ Data file does not exist")
                    return None

                try:
                    return load_productivity(self.data.get())
                except Exception as e:
                    self.log(f"❌ Error loading Excel file: {str(e)}")
                    return None

        except Exception as e:
            self.log(f"❌ Error loading data: {str(e)}")
            return None
    def _configure_styles(self):
        """Configure Dark Material Design-inspired ttk styles"""
        self.style = ttk.Style()
        self.style.theme_use('alt')  # Use alt theme as base

        # Define colors
        self.colors = {
            'background': '#121212',
            'card_background': '#1e1e1e',
            'surface': '#2b2d42',
            'surface_variant': '#3d405b',
            'primary': '#bb86fc',
            'primary_variant': '#3700b3',
            'secondary': '#03dac6',
            'on_surface': '#e1e1e1',
            'on_surface_variant': '#c1c1c1',
            'outline': '#404040',
            'error': '#cf6679',
            'success': '#4caf50',
            'warning': '#ff9800'
        }

        # Primary Button Style
        self.style.configure('Primary.TButton',
                             background=self.colors['primary'],
                             foreground=self.colors['background'],
                             font=('Segoe UI', 10, 'bold'),
                             borderwidth=0,
                             focuscolor='none',
                             padding=(16, 8)
                             )
        self.style.map('Primary.TButton',
                       background=[('active', self.colors['primary_variant']),
                                   ('disabled', self.colors['surface_variant'])],
                       foreground=[('disabled', self.colors['on_surface_variant'])]
                       )

        # Secondary Button Style
        self.style.configure('Secondary.TButton',
                             background=self.colors['surface'],
                             foreground=self.colors['primary'],
                             font=('Segoe UI', 10),
                             borderwidth=1,
                             relief='solid',
                             focuscolor='none',
                             padding=(12, 6)
                             )
        self.style.map('Secondary.TButton',
                       background=[('active', self.colors['surface_variant'])],
                       foreground=[('active', self.colors['primary_variant'])]
                       )

        # Action Button Style (for Send Single/Bulk Reports)
        self.style.configure('Action.TButton',
                             background=self.colors['secondary'],
                             foreground=self.colors['background'],
                             font=('Segoe UI', 11, 'bold'),
                             borderwidth=0,
                             focuscolor='none',
                             padding=(20, 10)
                             )
        self.style.map('Action.TButton',
                       background=[('active', self.colors['secondary']),
                                   ('pressed', self.colors['primary_variant'])]
                       )

        # Label Styles
        self.style.configure('Title.TLabel',
                             background=self.colors['background'],
                             foreground=self.colors['on_surface'],
                             font=('Segoe UI', 18, 'bold')
                             )

        self.style.configure('Heading.TLabel',
                             background=self.colors['card_background'],
                             foreground=self.colors['primary'],
                             font=('Segoe UI', 11, 'bold')
                             )

        self.style.configure('Normal.TLabel',
                             background=self.colors['card_background'],
                             foreground=self.colors['on_surface'],
                             font=('Segoe UI', 10)
                             )

        # Combobox Style
        self.style.configure('TCombobox',
                             background=self.colors['surface_variant'],
                             fieldbackground=self.colors['surface_variant'],
                             foreground=self.colors['on_surface'],
                             arrowcolor=self.colors['primary'],
                             selectbackground=self.colors['primary'],
                             selectforeground=self.colors['background']
                             )
        self.style.map('TCombobox',
                       fieldbackground=[('readonly', self.colors['surface_variant'])],
                       selectbackground=[('readonly', self.colors['primary'])],
                       selectforeground=[('readonly', self.colors['background'])]
                       )

        # Entry Style
        self.style.configure('TEntry',
                             fieldbackground=self.colors['surface_variant'],
                             foreground=self.colors['on_surface'],
                             insertcolor=self.colors['primary'],
                             borderwidth=1,
                             relief='solid'
                             )

        # Frame Style
        self.style.configure('Card.TFrame',
                             background=self.colors['card_background'],
                             relief='flat',
                             borderwidth=0
                             )

        # Scrollbar Style
        self.style.configure('TScrollbar',
                             background=self.colors['surface_variant'],
                             troughcolor=self.colors['surface'],
                             bordercolor=self.colors['outline'],
                             arrowcolor=self.colors['primary'],
                             relief='flat',
                             borderwidth=0
                             )
        self.style.map('TScrollbar',
                       background=[('active', self.colors['primary_variant']),
                                   ('pressed', self.colors['primary'])]
                       )

        # Radio button style (if using ttk.Radiobutton)
        self.style.configure('TRadiobutton',
                             background=self.colors['card_background'],
                             foreground=self.colors['on_surface'],
                             font=('Segoe UI', 10)
                             )
        self.style.map('TRadiobutton',
                       background=[('active', self.colors['surface_variant'])],
                       foreground=[('active', self.colors['primary'])]
                       )

        # Configure padding and spacing
        self.style.configure('.',
                             padding=4,
                             relief='flat',
                             borderwidth=0
                             )

        # LabelFrame Style (for sections)
        self.style.configure('Card.TLabelframe',
                             background=self.colors['card_background'],
                             foreground=self.colors['on_surface'],
                             borderwidth=0,
                             relief='flat'
                             )
        self.style.configure('Card.TLabelframe.Label',
                             background=self.colors['card_background'],
                             foreground=self.colors['primary'],
                             font=('Segoe UI', 11, 'bold')
                             )

    def _create_custom_radio_button(self, parent, text, variable, value, command=None):
        """Create a custom Material Design radio button"""
        # Create main container frame
        radio_frame = tk.Frame(parent, bg=self.colors['card_background'])

        # Create indicator container (for the radio circle)
        indicator_container = tk.Frame(radio_frame, bg=self.colors['card_background'], width=20, height=20)
        indicator_container.pack(side='left', padx=(0, 8))
        indicator_container.pack_propagate(False)  # Maintain fixed size

        # Create outer circle (border)
        outer_circle = tk.Canvas(
            indicator_container,
            width=20,
            height=20,
            bg=self.colors['card_background'],
            highlightthickness=0
        )
        outer_circle.pack(expand=True)

        # Create text label
        text_label = tk.Label(
            radio_frame,
            text=text,
            bg=self.colors['card_background'],
            fg=self.colors['on_surface'],
            font=('Segoe UI', 11)
        )
        text_label.pack(side='left', padx=(0, 8))

        def update_state(*args):
            """Update the radio button appearance"""
            if variable.get() == value:
                # Selected state
                outer_circle.delete('all')
                # Draw outer circle
                outer_circle.create_oval(4, 4, 16, 16,
                                         outline=self.colors['primary'],
                                         width=2)
                # Draw inner circle (filled dot)
                outer_circle.create_oval(7, 7, 13, 13,
                                         fill=self.colors['primary'],
                                         outline=self.colors['primary'])
                text_label.config(fg=self.colors['primary'])
            else:
                # Unselected state
                outer_circle.delete('all')
                outer_circle.create_oval(4, 4, 16, 16,
                                         outline=self.colors['outline'],
                                         width=2)
                text_label.config(fg=self.colors['on_surface_variant'])

        def on_click(event=None):
            """Handle radio button click"""
            variable.set(value)
            if command:
                command()

        # Bind click events to all components
        for widget in [radio_frame, indicator_container, outer_circle, text_label]:
            widget.bind('<Button-1>', on_click)
            widget.bind('<Enter>', lambda e: text_label.config(cursor='hand2'))
            widget.bind('<Leave>', lambda e: text_label.config(cursor=''))

        # Track variable changes
        variable.trace('w', update_state)

        # Initial state
        update_state()

        return radio_frame

    def _update_all_radio_buttons(self):
        """Update all radio button appearances"""
        # This method should be called when any radio button changes
        # to update the visual state of all radio buttons in groups
        pass

    def authenticate(self):
        """Handle Midway authentication"""
        try:
            self.auth_button.config(text="⌛ Authenticating...")
            self.update()

            etl = ETLDataFetcher()
            if messagebox.askokcancel(
                    "Midway Authentication",
                    "Please complete the Midway authentication in your browser.\n\nClick OK once you've logged in."
            ):
                self.is_authenticated = True
                self.auth_button.config(text="✅ Authenticated")
                self.log("✅ Successfully authenticated with Midway")
            else:
                raise Exception("Authentication cancelled by user")

        except Exception as e:
            self.log(f"❌ Authentication failed: {str(e)}")
            self.auth_button.config(text="❌ Auth Failed")
            self.is_authenticated = False

            messagebox.showerror(
                "Authentication Error",
                "Failed to authenticate with Midway.\nPlease try again or use Excel files."
            )

    def _create_card(self, parent, title, pady=(0, 16)):
        """Create a Dark Material Design card-like frame"""
        # Main card container
        card_container = tk.Frame(parent, bg=self.colors['background'])
        card_container.pack(fill="x", pady=pady, padx=24)

        # Card with elevation effect
        card = tk.Frame(card_container, bg=self.colors['card_background'], relief='flat', bd=0)
        card.pack(fill="x")

        # Add subtle shadow/elevation effect
        shadow_frame = tk.Frame(card_container, bg='#0d0d0d', height=1)
        shadow_frame.pack(fill="x", pady=(0, 4))

        # Card header
        if title:
            header = tk.Frame(card, bg=self.colors['card_background'], height=56)
            header.pack(fill="x", pady=(0, 8))
            header.pack_propagate(False)

            title_label = tk.Label(
                header,
                text=title,
                bg=self.colors['card_background'],
                fg=self.colors['on_surface'],
                font=('Segoe UI', 14, 'bold'),
                anchor='w'
            )
            title_label.pack(side='left', padx=20, pady=16)

        # Card content frame
        content = tk.Frame(card, bg=self.colors['card_background'])
        content.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        return card, content

    def _create_input_field(self, parent, label_text, string_var, button_text="Browse", command=None):
        """Create a Dark Material Design input field"""
        field_frame = tk.Frame(parent, bg=self.colors['card_background'])
        field_frame.pack(fill="x", pady=(0, 20))

        # Label
        label = tk.Label(
            field_frame,
            text=label_text,
            bg=self.colors['card_background'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 11, 'bold')
        )
        label.pack(anchor='w', pady=(0, 8))

        # Input container
        input_container = tk.Frame(field_frame, bg=self.colors['card_background'])
        input_container.pack(fill="x")

        # Entry field with Dark Material Design styling
        entry_frame = tk.Frame(
            input_container,
            bg=self.colors['surface_variant'],
            relief='solid',
            bd=1,
            highlightthickness=0
        )
        entry_frame.pack(side='left', fill='x', expand=True, padx=(0, 12))

        entry = tk.Entry(
            entry_frame,
            textvariable=string_var,
            font=('Segoe UI', 10),
            bg=self.colors['surface_variant'],
            fg=self.colors['on_surface'],
            relief='flat',
            bd=0,
            insertbackground=self.colors['primary'],
            highlightthickness=0
        )
        entry.pack(fill='both', expand=True, padx=12, pady=8)

        # Browse button
        if command:
            browse_btn = ttk.Button(
                input_container,
                text=button_text,
                command=command,
                style='Secondary.TButton'
            )
            browse_btn.pack(side='right')

        return field_frame, entry

    def _build_ui(self):
        """Build the main UI with Dark Material Design principles"""
        # Configure grid weights for main window
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Main container with scrolling
        canvas = tk.Canvas(self, bg=self.colors['background'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['background'])

        # Configure scrollable frame
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        # Create window in canvas and store the window id
        window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Grid scrolling components with proper weights and sticky
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # App bar / Header
        self._create_app_bar(scrollable_frame)

        # Data Source Card
        self._create_data_source_card(scrollable_frame)

        # Report Configuration Card
        self._create_report_config_card(scrollable_frame)

        # File Selection Card
        self._create_file_selection_card(scrollable_frame)

        # Parameters Card
        self._create_parameters_card(scrollable_frame)

        # Action Buttons Card
        self._create_action_buttons_card(scrollable_frame)

        # Activity Log Card
        self._create_activity_log_card(scrollable_frame)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Update canvas width when window resizes
        def _on_canvas_configure(event):
            canvas.itemconfig(window_id, width=event.width)

        canvas.bind("<Configure>", _on_canvas_configure)

    def _create_app_bar(self, parent):
        """Create Dark Material Design app bar"""
        app_bar = tk.Frame(parent, bg=self.colors['surface'], height=64)
        app_bar.pack(fill="x")
        app_bar.pack_propagate(False)

        # Title
        title = tk.Label(
            app_bar,
            text="📊 Email Report Generator",
            bg=self.colors['surface'],
            fg=self.colors['on_surface'],
            font=('Segoe UI', 20, 'bold')
        )
        title.pack(side='left', padx=24, pady=16)

        # Subtitle
        subtitle = tk.Label(
            app_bar,
            text="Generate and send productivity reports with ease",
            bg=self.colors['surface'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 11)
        )
        subtitle.pack(side='left', padx=(0, 24), pady=16)

    def _create_data_source_card(self, parent):
        """Create data source selection card"""
        card, content = self._create_card(parent, "🔗 Data Source")

        # Radio button container
        radio_container = tk.Frame(content, bg=self.colors['card_background'])
        radio_container.pack(fill="x", pady=(0, 16))

        # Excel Files radio button
        excel_radio = self._create_custom_radio_button(
            radio_container,
            "📊 Excel Files",
            self.data_source,
            "excel",
            self._toggle_data_source
        )
        excel_radio.pack(fill="x", pady=(0, 8))

        # ETL Data radio button
        etl_radio = self._create_custom_radio_button(
            radio_container,
            "🔄 ETL Data",
            self.data_source,
            "etl",
            self._toggle_data_source
        )
        etl_radio.pack(fill="x")

        # Create authenticate button
        self.auth_button = ttk.Button(
            content,
            text="🔐 Authenticate Midway",
            command=self.authenticate,
            style='Primary.TButton',
            state="disabled"
        )
        self.auth_button.pack(side="left", pady=(8, 0))

        # Store reference to content frame
        self.source_frame = content

        return card

    def _create_report_config_card(self, parent):
        """Create report configuration card"""
        card, content = self._create_card(parent, "📋 Report Configuration")

        # Report type selection
        type_frame = tk.Frame(content, bg=self.colors['card_background'])
        type_frame.pack(fill="x", pady=(0, 16))

        tk.Label(
            type_frame,
            text="Report Type",
            bg=self.colors['card_background'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 11, 'bold')
        ).pack(anchor='w', pady=(0, 8))

        self.report_type = ttk.Combobox(
            type_frame,
            values=["RDR Report", "Precision Correction Report", "Timesheet Missing Report"],
            state="readonly",
            font=('Segoe UI', 11),
            style='Material.TCombobox',
            width=40
        )
        self.report_type.set("RDR Report")
        self.report_type.pack(anchor='w')
        self.report_type.bind('<<ComboboxSelected>>', lambda e: self._update_fields())

        # Info label
        self.info_label = tk.Label(
            content,
            text="",
            bg=self.colors['card_background'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 10),
            wraplength=600,
            justify='left'
        )
        self.info_label.pack(anchor='w', pady=(16, 0))

    def _create_file_selection_card(self, parent):
        """Create file selection card"""
        card, content = self._create_card(parent, "📁 File Selection")

        # Create file input frames
        self.data_frame, _ = self._create_input_field(
            content, "📊 Excel File (with Volume, Quality, TimeSheet, and Manager sheets)",
            self.data,
            command=lambda: self._browse(self.data)
        )

        # Initialize other frames but don't pack them
        self.qual_frame = tk.Frame(content)
        self.mgr_frame = tk.Frame(content)
        self.timesheet_frame = tk.Frame(content)

        return card

    def _create_parameters_card(self, parent):
        """Create parameters card"""
        card, content = self._create_card(parent, "⚙️ Parameters")

        # Parameters grid
        params_grid = tk.Frame(content, bg=self.colors['card_background'])
        params_grid.pack(fill="x")
        params_grid.columnconfigure(0, weight=1)
        params_grid.columnconfigure(1, weight=1)

        # Weeks input
        weeks_frame = tk.Frame(params_grid, bg=self.colors['card_background'])
        weeks_frame.grid(row=0, column=0, sticky="ew", padx=(0, 16))

        tk.Label(
            weeks_frame,
            text="📅 Weeks (comma-separated)",
            bg=self.colors['card_background'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 11, 'bold')
        ).pack(anchor='w', pady=(0, 8))

        weeks_entry_frame = tk.Frame(
            weeks_frame,
            bg=self.colors['surface_variant'],
            relief='solid',
            bd=1
        )
        weeks_entry_frame.pack(fill='x')

        tk.Entry(
            weeks_entry_frame,
            textvariable=self.wks,
            font=('Segoe UI', 10),
            bg=self.colors['surface_variant'],
            fg=self.colors['on_surface'],
            relief='flat',
            bd=0
        ).pack(fill='x', padx=12, pady=8)

        # User input
        user_frame = tk.Frame(params_grid, bg=self.colors['card_background'])
        user_frame.grid(row=0, column=1, sticky="ew")

        tk.Label(
            user_frame,
            text="👤 User Login (for single reports)",
            bg=self.colors['card_background'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 11, 'bold')
        ).pack(anchor='w', pady=(0, 8))

        user_entry_frame = tk.Frame(
            user_frame,
            bg=self.colors['surface_variant'],
            relief='solid',
            bd=1
        )
        user_entry_frame.pack(fill='x')

        tk.Entry(
            user_entry_frame,
            textvariable=self.user,
            font=('Segoe UI', 10),
            bg=self.colors['surface_variant'],
            fg=self.colors['on_surface'],
            relief='flat',
            bd=0
        ).pack(fill='x', padx=12, pady=8)

        # Email mode selection
        mode_frame = tk.Frame(content, bg=self.colors['card_background'])
        mode_frame.pack(fill="x", pady=(24, 0))

        tk.Label(
            mode_frame,
            text="📧 Email Mode",
            bg=self.colors['card_background'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 11, 'bold')
        ).pack(anchor='w', pady=(0, 12))

        # Email mode radio buttons container
        mode_radio_container = tk.Frame(mode_frame, bg=self.colors['card_background'])
        mode_radio_container.pack(fill="x")

        # Preview Only radio button
        preview_radio = self._create_custom_radio_button(
            mode_radio_container,
            "👀 Preview Only",
            self.mode,
            "preview"
        )
        preview_radio.pack(side='left', padx=(0, 24))

        # Send Email radio button
        send_radio = self._create_custom_radio_button(
            mode_radio_container,
            "📤 Send Email",
            self.mode,
            "send"
        )
        send_radio.pack(side='left')

        # Month selection
        month_frame = tk.Frame(content, bg=self.colors['card_background'])
        month_frame.pack(fill="x", pady=(24, 0))

        tk.Label(
            month_frame,
            text="📅 Month Selection",
            bg=self.colors['card_background'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 11, 'bold')
        ).pack(anchor='w', pady=(0, 12))

        # Month radio buttons container
        month_radio_container = tk.Frame(month_frame, bg=self.colors['card_background'])
        month_radio_container.pack(fill="x")

        # Current Month radio button
        current_radio = self._create_custom_radio_button(
            month_radio_container,
            "📅 Current Month",
            self.month_selection,
            "current"
        )
        current_radio.pack(side='left', padx=(0, 24))

        # Previous Month radio button
        prev_radio = self._create_custom_radio_button(
            month_radio_container,
            "📅 Previous Month",
            self.month_selection,
            "previous"
        )
        prev_radio.pack(side='left')

    def _create_action_buttons_card(self, parent):
        """Create action buttons card"""
        card, content = self._create_card(parent, "")

        # Button container
        button_container = tk.Frame(content, bg=self.colors['card_background'])
        button_container.pack(pady=8)

        # Single report button
        ttk.Button(
            button_container,
            text="📧 Send Single Report",
            command=self.send_single,
            style='Action.TButton'
        ).pack(side="left", padx=(0, 16))

        # Bulk report button
        ttk.Button(
            button_container,
            text="📨 Send Bulk Reports",
            command=self.send_bulk,
            style='Action.TButton'
        ).pack(side="left")

    def _create_activity_log_card(self, parent):
        """Create activity log card"""
        try:
            card, content = self._create_card(parent, "💬 Activity Log")

            # Text widget with scrollbar
            text_container = tk.Frame(content, bg=self.colors['card_background'])
            text_container.pack(fill="both", expand=True)

            # Scrollbar
            log_scrollbar = ttk.Scrollbar(text_container, orient="vertical")
            log_scrollbar.pack(side="right", fill="y")

            # Text widget with Dark Material Design colors
            self.status = tk.Text(
                text_container,
                height=12,
                font=('Consolas', 10),
                bg='#0d1117',  # Dark code background
                fg='#c9d1d9',  # Light code text
                insertbackground=self.colors['primary'],
                selectbackground=self.colors['primary_variant'],
                selectforeground='white',
                wrap=tk.WORD,
                relief='flat',
                bd=0,
                yscrollcommand=log_scrollbar.set
            )
            self.status.pack(side="left", fill="both", expand=True)

            # Connect scrollbar
            log_scrollbar.config(command=self.status.yview)

            # Welcome message
            self.log("🚀 Welcome to Email Report Generator!")
            self.log("📋 Select a report type and configure the required files to get started.")

        except Exception as e:
            print(f"Error creating activity log card: {str(e)}")

    def html_metric_value_table_with_latest(self, data, weeks, percentiles, section="Quality", is_quality=False,
                                            month_data=None):
        from datetime import datetime
        import calendar

        styles = get_common_styles()

        # Get the month name based on current/previous selection
        current_date = datetime.now()
        if self.month_selection.get() == "current":
            month_name = calendar.month_name[current_date.month]
        else:
            # For previous month
            previous_month = 12 if current_date.month == 1 else current_date.month - 1
            month_name = calendar.month_name[previous_month]

        # Calculate column widths
        num_weeks = len(weeks)
        pipeline_width = 25
        week_width = (60 - 15) / num_weeks
        latest_width = 15
        month_width = 15

        # Create header cells
        ths = "".join(
            f"<th style='{styles['header_style']} width: {week_width}%;'>{w}</th>"
            for w in weeks
        )
        ths += f"<th style='{styles['header_style']} width: {latest_width}%;'>Latest</th>"
        ths += f"<th style='{styles['header_style']} width: {month_width}%;'>Month({month_name})</th>"

        # Create rows with monthly data
        rows = ""
        for pipe in sorted(data):
            tds = ""
            for w in weeks:
                if is_quality:
                    val_dict = data[pipe].get(w, None)
                    val = val_dict['score'] if isinstance(val_dict, dict) and 'score' in val_dict else 0
                    err = val_dict['err'] if isinstance(val_dict, dict) and 'err' in val_dict else 0
                    total = val_dict['total'] if isinstance(val_dict, dict) and 'total' in val_dict else 0
                    disp = f"{val:.1%}<br>({err}/{total})" if total else "-"
                else:
                    val = data[pipe].get(w, 0)
                    disp = f"{val:.1f}" if val else "-"

                tds += f"<td style='{styles['cell_style']} width: {week_width}%;'>{disp}</td>"

            # Add latest benchmark
            val = data[pipe].get(weeks[-1], 0)
            val_scalar = val['score'] if is_quality and isinstance(val, dict) and 'score' in val else val
            pctls = percentiles[weeks[-1]].get(pipe, None) if percentiles and weeks[-1] in percentiles else None
            bench = percentile_label(val_scalar, pctls)
            bg_color, text_color = pct_bg_fg(bench)

            benchmark_style = f'''
                padding: 8px 6px;
                border: 1px solid white;
                text-align: center;
                font-size: 12px;
                background-color: {bg_color};
                color: {text_color};
                height: 32px;
                vertical-align: middle;
                word-wrap: break-word;
                overflow: hidden;
                width: {latest_width}%;
            '''

            tds += f"<td style='{benchmark_style}'>{bench if bench != '-' else '-'}</td>"

            # Add month data
            if month_data and pipe in month_data:
                month_val = month_data[pipe]
                month_disp = f"{month_val:.1f}" if not is_quality else f"{month_val:.1%}"
            else:
                month_disp = "-"

            tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{month_disp}</td>"

            rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'>{pipe}</td>{tds}</tr>"

        return f"""
        <div style="{styles['container_style']}">
            <table style="{styles['table_style']}">
                <tr>
                    <td colspan="{len(weeks) + 3}" style="{styles['title_style']}">
                        Weekly {section} Metrics
                    </td>
                </tr>
                <tr>
                    <th style="{styles['header_style']} width: {pipeline_width}%;">Pipeline</th>
                    {ths}
                </tr>
                {rows}
            </table>
        </div>
        """

    # Keep all the existing methods for functionality
    def _toggle_data_source(self):
        """Toggle between data sources"""
        current_source = self.data_source.get()

        if current_source == "etl":
            # Enable authentication button
            self.auth_button.configure(state="normal")

            # Hide file input fields
            self.data_frame.pack_forget()
            self.qual_frame.pack_forget()

            # Update info label
            self.info_label.config(
                text="Using ETL as data source. Authentication required to proceed."
            )
        else:
            # Disable authentication button
            self.auth_button.configure(state="disabled")

            # Show file input fields based on current report type
            self._update_fields()

            # Update info label
            self.info_label.config(
                text="Using Excel files as data source."
            )

        # Force update
        self.update_idletasks()

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
            widget.pack_forget()

        # Only show fields if using Excel as data source
        if self.data_source.get() == "excel":
            if report_type == "RDR Report":
                self.data_frame.pack(fill="x", pady=(0, 20))
                self.qual_frame.pack(fill="x", pady=(0, 20))
                self.mgr_frame.pack(fill="x", pady=(0, 20))
                self.info_label.config(
                    text="📋 RDR Report requires Data Excel, Quality Excel, and Manager Mapping files")

            elif report_type == "Precision Correction Report":
                self.data_frame.pack(fill="x", pady=(0, 20))
                self.mgr_frame.pack(fill="x", pady=(0, 20))
                self.info_label.config(
                    text="🎯 Precision Correction Report requires Data Excel and Manager Mapping files")

            elif report_type == "Timesheet Missing Report":
                self.timesheet_frame.pack(fill="x", pady=(0, 20))
                self.mgr_frame.pack(fill="x", pady=(0, 20))
                self.info_label.config(
                    text="⏰ Timesheet Missing Report requires Timesheet Excel and Manager Mapping files")

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
            if timesheet_data is None or timesheet_data.empty:
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
                    timesheet_data = analyze_timesheet_missing(self.timesheet.get(), weeks[-1], user)
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

    def send_single(self):
        """Send single user report with detailed debugging"""
        user = self.user.get().strip()
        if not user:
            return self.log("❌ User required.")

        weeks = self._weeks()
        if not weeks:
            return

        try:
            # Get Excel file path
            excel_file_path = self.data.get()
            if not excel_file_path or not os.path.exists(excel_file_path):
                self.log("❌ Excel file not found or not specified")
                return
            self.log(f"Using Excel file: {excel_file_path}")

            # Validate Excel file structure
            is_valid, message = validate_excel_file(excel_file_path)
            if not is_valid:
                self.log(f"❌ Invalid Excel file: {message}")
                return

            # Load productivity data
            self.log("\nProcessing productivity data...")
            df = self.load_data()
            if df is None:
                return

            self.log(f"Productivity data loaded, shape: {df.shape}")

            # Calculate productivity metrics
            prod = user_prod_dict(df, user, weeks)
            self.log(f"User productivity calculated for {len(prod)} pipelines")

            allprod = all_prod_dict(df, weeks)
            self.log(f"All users productivity calculated for {len(allprod)} users")

            prod_pct = productivity_percentiles(allprod, weeks)
            self.log("Productivity percentiles calculated")

            # Calculate monthly productivity metrics
            monthly_prod_metrics = calculate_monthly_productivity_metrics(df, self.month_selection.get() == "current")
            self.log("Monthly productivity metrics calculated")

            # In send_single method, add:
            monthly_subreason = calculate_monthly_subreason_metrics(qdf, self.month_selection.get() == "current")
            monthly_correction = calculate_monthly_correction_metrics(qdf, self.month_selection.get() == "current")

            qc2_left = self.html_qc2_reason_value_table(subreason, weeks, month_data=monthly_subreason)
            correction_quality_table = self.html_correction_type_quality_table(correction_data, weeks,
                                                                               month_data=monthly_correction)

            # Initialize quality tables
            qual_table = qual_pct_table = qc2_left = qc2_right = ""
            correction_quality_table = correction_count_table = ""

            # Load quality data from Quality sheet
            self.log("\nProcessing quality data...")
            try:
                qdf = load_quality(excel_file_path)
                self.log(f"Quality data loaded, shape: {qdf.shape}")
                self.log(f"Quality columns: {qdf.columns.tolist()}")

                if not qdf.empty:
                    # Calculate quality metrics
                    qual = user_quality_dict(qdf, user, weeks)
                    self.log(f"Quality metrics calculated for user: {bool(qual)}")

                    allqual = all_quality_dict(qdf, weeks)
                    self.log(f"Quality metrics calculated for all users: {len(allqual)}")

                    qual_pct = quality_percentiles(allqual, weeks)
                    self.log("Quality percentiles calculated")

                    # Calculate monthly quality metrics
                    monthly_qual_metrics = calculate_monthly_quality_metrics(qdf,
                                                                             self.month_selection.get() == "current")
                    self.log("Monthly quality metrics calculated")

                    # Generate quality tables
                    qual_table = self.html_metric_value_table_with_latest(
                        qual, weeks, qual_pct, section="Quality",
                        is_quality=True, month_data=monthly_qual_metrics)
                    qual_pct_table = html_metric_pct_table(
                        qual, weeks, qual_pct, section="Quality")

                    # Calculate QC2 analysis
                    self.log("\nProcessing QC2 analysis...")
                    subreason = qc2_subreason_analysis(qdf, user, weeks)
                    qc2_left = html_qc2_reason_value_table(subreason, weeks)
                    qc2_right = html_qc2_reason_pct_table(subreason, weeks)

                    # Calculate correction analysis
                    self.log("\nProcessing correction type analysis...")
                    correction_data = calculate_correction_type_data(qdf, user, weeks)
                    correction_quality_table = html_correction_type_quality_table(correction_data, weeks)
                    correction_count_table = html_correction_type_count_table(correction_data, weeks)
                else:
                    self.log("⚠️ No quality data found for processing")
            except Exception as qe:
                self.log(f"⚠️ Error processing quality data: {str(qe)}")
                import traceback
                self.log(traceback.format_exc())

            # Process timesheet data
            timesheet_table = ""
            self.log("\nProcessing timesheet data...")
            try:
                timesheet_data = analyze_timesheet_missing(excel_file_path, weeks[-1], user)
                if timesheet_data is not None and not timesheet_data.empty:
                    timesheet_table = html_timesheet_missing_table(timesheet_data)
                    self.log("Timesheet data processed successfully")
                else:
                    self.log("ℹ️ No timesheet missing data found")
            except Exception as te:
                self.log(f"⚠️ Error processing timesheet: {str(te)}")

            # Generate productivity tables
            prod_table = self.html_metric_value_table_with_latest(
                prod, weeks, prod_pct, section="Productivity",
                is_quality=False, month_data=monthly_prod_metrics)
            prod_pct_table = html_metric_pct_table(
                prod, weeks, prod_pct, section="Productivity")

            # Compose final HTML
            self.log("\nGenerating final HTML report...")
            html = compose_html(
                user=user,
                prod_table=prod_table,
                prod_pct_table=prod_pct_table,
                qual_table=qual_table,
                qual_pct_table=qual_pct_table,
                qc2_left=qc2_left,
                qc2_right=qc2_right,
                correction_quality_table=correction_quality_table,
                correction_count_table=correction_count_table,
                timesheet_table=timesheet_table
            )

            # Send or preview email
            self.log("\nPreparing email...")
            send_mail_html(
                to=f"{user}@amazon.com",
                cc="",
                subject="RDR Productivity & Quality Metrics Report",
                html=html,
                preview=self.mode.get() == "preview"
            )

            mode_text = "previewed" if self.mode.get() == "preview" else "sent"
            self.log(f"✅ Report for {user} {mode_text} successfully!")

        except Exception as e:
            self.log(f"❌ Error generating report: {str(e)}")
            import traceback
            self.log(traceback.format_exc())

    def send_bulk(self):
        """Send bulk reports to multiple users"""
        if not self.mgr.get():
            return self.log("❌ Manager mapping required.")

        weeks = self._weeks()
        if not weeks:
            return

        try:
            # Load data based on selected source
            df = self.load_data()
            if df is None:
                return

            # Calculate monthly productivity metrics
            monthly_prod_metrics = calculate_monthly_productivity_metrics(df, self.month_selection.get() == "current")
            self.log("Monthly productivity metrics calculated")

            mgr = pd.read_excel(self.mgr.get())
            allprod = all_prod_dict(df, weeks)
            prod_pct = productivity_percentiles(allprod, weeks)
            allqual = qual_pct = {}
            qdf = None

            # Process quality data if available
            if self.qual.get() and os.path.exists(self.qual.get()):
                qdf = load_quality(self.qual.get())
                allqual = all_quality_dict(qdf, weeks)
                qual_pct = quality_percentiles(allqual, weeks)
                monthly_qual_metrics = calculate_monthly_quality_metrics(qdf, self.month_selection.get() == "current")

            for user in mgr['loginname'].dropna().unique():
                try:
                    # Generate productivity tables
                    prod = allprod.get(user, {})
                    prod_table = self.html_metric_value_table_with_latest(
                        prod, weeks, prod_pct, section="Productivity",
                        is_quality=False, month_data=monthly_prod_metrics)
                    prod_pct_table = html_metric_pct_table(
                        prod, weeks, prod_pct, section="Productivity")

                    # Initialize quality tables
                    qual_table = qual_pct_table = qc2_left = qc2_right = ""
                    correction_quality_table = correction_count_table = ""

                    # Generate quality tables if quality data exists
                    if allqual:
                        qual = allqual.get(user, {})
                        qual_table = self.html_metric_value_table_with_latest(
                            qual, weeks, qual_pct, section="Quality",
                            is_quality=True, month_data=monthly_qual_metrics)
                        qual_pct_table = html_metric_pct_table(
                            qual, weeks, qual_pct, section="Quality")

                        # Generate QC2 analysis
                        subreason = qc2_subreason_analysis(qdf, user, weeks)
                        qc2_left = html_qc2_reason_value_table(subreason, weeks)
                        qc2_right = html_qc2_reason_pct_table(subreason, weeks)

                        # Generate correction analysis
                        correction_data = calculate_correction_type_data(qdf, user, weeks)
                        correction_quality_table = html_correction_type_quality_table(correction_data, weeks)
                        correction_count_table = html_correction_type_count_table(correction_data, weeks)

                    # Process timesheet data if available
                    timesheet_table = ""
                    if self.timesheet.get() and os.path.exists(self.timesheet.get()):
                        timesheet_data = analyze_timesheet_missing(self.timesheet.get(), weeks[-1], user)
                        if timesheet_data is not None and not timesheet_data.empty:
                            timesheet_table = html_timesheet_missing_table(timesheet_data)

                    # Compose and send email
                    html = compose_html(
                        user, prod_table, prod_pct_table, qual_table, qual_pct_table,
                        qc2_left, qc2_right, correction_quality_table, correction_count_table,
                        timesheet_table
                    )

                    # Get CC email if supervisor exists
                    cc = ""
                    if 'supervisorloginname' in mgr.columns:
                        cc_val = mgr[mgr['loginname'] == user]['supervisorloginname'].iloc[0]
                        if pd.notna(cc_val):
                            cc = f"{cc_val}@amazon.com"

                    send_mail_html(
                        f"{user}@amazon.com",
                        cc,
                        "RDR Productivity & Quality Metrics Report",
                        html,
                        preview=self.mode.get() == "preview"
                    )

                    mode_text = "previewed" if self.mode.get() == "preview" else "sent"
                    self.log(f"✅ Report for {user} {mode_text}")

                except Exception as e:
                    self.log(f"⚠️ Error processing {user}: {str(e)}")
                    continue

            self.log("🎉 Bulk processing completed!")

        except Exception as e:
            self.log(f"❌ Error generating reports: {str(e)}")

    def send_precision_single(self):
        user = self.user.get().strip()
        if not user or not self.data.get():
            self.log("❌ User and data required.")
            return

        weeks = self._weeks()
        if not weeks:
            return

        try:
            # Load data based on selected source
            df = self.load_data()
            if df is None:
                return

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
                self.log(f"✅ Precision mail for {user} {'previewed' if self.mode.get() == 'preview' else 'sent'}")
            else:
                self.log(f"ℹ️ No high precision corrections found for {user}")

        except Exception as e:
            self.log(f"❌ Error generating precision report: {str(e)}")

    def send_precision_bulk(self):
        """Send precision correction reports in bulk"""
        try:
            # Validate input data
            if not self.data.get():
                self.log("❌ Data required.")
                return False

            # Get weeks
            weeks = self._weeks()
            if not weeks:
                self.log("❌ No valid weeks specified")
                return False

            # Load and process data
            try:
                df = load_productivity(self.data.get())
                precision_data = calculate_precision_corrections(df, weeks)

                if not precision_data:
                    self.log("ℹ️ No high precision corrections found for any user")
                    return False

                self.log(f"   Processing precision reports for {len(precision_data)} users...")

                # Process each user
                for user in precision_data:
                    try:
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

                        # Send email
                        send_mail_html(
                            f"{user}@amazon.com",
                            "",
                            "High Precision Corrections Alert",
                            html,
                            preview=self.mode.get() == "preview"
                        )

                        # Log success
                        mode_text = "previewed" if self.mode.get() == "preview" else "sent"
                        self.log(f"✅ Precision mail for {user} {mode_text}")

                    except Exception as user_exc:
                        self.log(f"⚠️ Error processing user {user}: {str(user_exc)}")
                        continue

                # Log completion
                self.log("   Bulk precision correction processing completed!")
                return True

            except Exception as data_exc:
                self.log(f"❌ Error processing data: {str(data_exc)}")
                return False

        except Exception as exc:
            self.log(f"❌ Error in bulk precision processing: {str(exc)}")
            return False

if __name__ == "__main__":
    ReporterApp().mainloop()
