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
from datetime import datetime
import calendar

# ----------------- Data Processing -------------------
def load_productivity(data_path):
    try:
        cols = ['useralias', 'pipeline', 'p_week', 'processed_volume', 'processed_volumetr',
                'processed_time', 'processed_time_tr', 'precision_correction', 'acl', 'level','p_month']

        df = pd.read_excel(data_path, sheet_name='Volume', usecols=cols)

        for c in ['processed_volume', 'processed_volumetr', 'processed_time', 'processed_time_tr',
                  'precision_correction']:
            df[c] = df[c].fillna(0)

        # Extract numeric part from level values (e.g., 'L3' becomes '3')
        df['level'] = df['level'].fillna(0).astype(str).str.extract(r'(\d+)').fillna(0).astype(int)
        auditor_levels = df.groupby('useralias')['level'].first().to_dict()

        return df, auditor_levels
    except Exception as e:
        print(f"Error in load_productivity: {str(e)}")
        return None, None


def get_common_styles():
    """Common styles for all tables to ensure consistency, including fixing top-left border issue."""
    return {
        'container_style': '''
            margin: 10px 0;
            font-family: Arial, sans-serif;
        ''',
        'table_style': '''
            border-collapse: separate;
            border-spacing: 0;
            width: 100%;
            font-family: Arial, sans-serif;
        ''',
        'title_style': '''
            background-color: #000000;
            color: white;
            text-align: center;
            font-size: 14px;
            font-weight: bold;
            padding: 12px;
            border-top: 3px solid white;
            border-bottom: 1px solid white;
            border-left: 3px solid white;   /* Ensures alignment with table's left border */
            border-right: 3px solid white;
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
        print("Rows in loaded sheet:", df.shape[0])
        print("Program values:", df['program'].unique())
        print("\nDebugging load_quality:")
        print(f"Raw quality data shape: {df.shape}")
        print(f"All available columns: {df.columns.tolist()}")

        # Check if the column name might be different
        possible_correction_columns = [col for col in df.columns if 'correction' in col.lower()]
        print(f"Possible correction-related columns: {possible_correction_columns}")

        # Define required columns with correction type
        required_cols = ['volume', 'auditor_login', 'program', 'usecase',
                         'qc2_judgement', 'qc2_subreason', 'week',
                         'auditor_correction_type','auditor_reappeal_final_judgement','month']  # Make sure this matches exact column name

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
    try:
        dfu = df[df['useralias'] == user]
        if weeks:
            dfu = dfu[dfu['p_week'].isin(weeks)]

        dfu['total_volume'] = dfu['processed_volume'] + dfu['processed_volumetr']
        dfu['total_time'] = dfu['processed_time'] + dfu['processed_time_tr']

        g = dfu.groupby(['pipeline', 'p_week'], as_index=False).agg({
            'total_volume': 'sum',
            'total_time': 'sum',
            'level': 'first'
        })

        g['productivity'] = np.where(g['total_time'] > 0, g['total_volume'] / g['total_time'], 0)

        result = {}
        for _, row in g.iterrows():
            result.setdefault(row['pipeline'], {})[row['p_week']] = row['productivity']

        return result
    except Exception as e:
        print(f"Error in user_prod_dict: {str(e)}")
        return {}

def all_prod_dict(df, weeks=None):
    if weeks: df = df[df['p_week'].isin(weeks)]
    df.loc[:, 'total_volume'] = df['processed_volume'] + df['processed_volumetr']
    df.loc[:, 'total_time'] = df['processed_time'] + df['processed_time_tr']
    g = df.groupby(['useralias','pipeline','p_week'], as_index=False)[['total_volume','total_time']].sum()
    g['productivity'] = np.where(g['total_time'] > 0, g['total_volume']/g['total_time'], 0)
    d = {}
    for row in g.itertuples(index=False):
        d.setdefault(row.useralias, {}).setdefault(row.pipeline, {})[row.p_week] = row.productivity
    return d


def productivity_percentiles(allprod, weeks, auditor_levels):
    out = {}
    try:
        for week in weeks:
            pipeline_vals = {}
            for auditor, pipelines in allprod.items():
                level = auditor_levels.get(auditor)
                if level:
                    for pipe, values in pipelines.items():
                        if week in values:
                            pipeline_vals.setdefault(pipe, {}).setdefault(level, []).append(values[week])

            out[week] = {}
            for pipe in pipeline_vals:
                out[week][pipe] = {}
                for level in pipeline_vals[pipe]:
                    vals = pipeline_vals[pipe][level]
                    if vals:
                        out[week][pipe][level] = {
                            'p30': np.percentile(vals, 30),
                            'p50': np.percentile(vals, 50),
                            'p75': np.percentile(vals, 75),
                            'p90': np.percentile(vals, 90)
                        }
        return out
    except Exception as e:
        print(f"Error in productivity_percentiles: {str(e)}")
        return {}


def all_quality_dict(df, weeks=None):
    try:
        if weeks:
            df = df[df['week'].isin(weeks)]
        d = {}
        for auditor, grp in df.groupby('auditor_login'):
            d[auditor] = {}
            for (pipe, week), subgrp in grp.groupby(['usecase', 'week']):
                if isinstance(subgrp, int):
                    continue
                total_volume = subgrp['volume'].sum() if hasattr(subgrp, '__len__') else 0

                # Modified error calculation
                error_mask = (
                        df['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                        ~df['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
                )
                errors = subgrp[error_mask]['volume'].sum() if hasattr(subgrp, '__len__') else 0

                score = (total_volume - errors) / total_volume if total_volume else 0
                d[auditor].setdefault(pipe, {})[week] = {
                    'score': score,
                    'err': errors,
                    'total': total_volume
                }
        return d
    except Exception as e:
        print(f"Error in all_quality_dict: {str(e)}")
        return {}


def user_quality_dict(df, user, weeks=None):
    try:
        dfu = df[df['auditor_login'] == user].copy()
        if weeks:
            dfu = dfu[dfu['week'].isin(weeks)]
        d = {}
        for (pipe, week), grp in dfu.groupby(['usecase', 'week']):
            if isinstance(grp, int):
                continue

            total_volume = grp['volume'].sum() if hasattr(grp, '__len__') else 0

            # Modified error calculation to properly filter reappeal judgments
            error_mask = (
                    df['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                    ~df['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
            )
            errors = grp[error_mask]['volume'].sum() if hasattr(grp, '__len__') else 0

            score = (total_volume - errors) / total_volume if total_volume else 0
            d.setdefault(pipe, {})[week] = {
                'score': score,
                'err': errors,
                'total': total_volume
            }
        return d
    except Exception as e:
        print(f"Error in user_quality_dict: {str(e)}")
        return {}

def analyze_timesheet_missing(timesheet_path, week, user):
    """Analyze timesheet missing data with improved efficiency"""
    try:
        # Read only required columns
        df = pd.read_excel(timesheet_path, sheet_name='TimeSheet',
                           usecols=['work_date', 'week', 'timesheet_missing', 'loginid'])

        # Filter data using loc instead of creating a copy
        mask = (df['week'] == week) & (df['loginid'] == user)
        weekly_data = df.loc[mask].copy()  # Create explicit copy

        if weekly_data.empty:
            return None

        # Convert dates using loc
        weekly_data.loc[:, 'work_date'] = pd.to_datetime(weekly_data['work_date'])

        # Filter for missing timesheet entries
        daily_missing = weekly_data.loc[weekly_data['timesheet_missing'] > 35,
        ['work_date', 'timesheet_missing']]

        return daily_missing.sort_values('work_date') if not daily_missing.empty else None

    except Exception as exc:
        print(f"Error processing timesheet: {str(exc)}")
        return None


def quality_percentiles(allqual, weeks, auditor_levels):
    out = {}
    for week in weeks:
        pipeline_vals = {}
        for auditor, pipelines in allqual.items():
            level = auditor_levels.get(auditor)
            if level:  # Only process if level is known
                for pipe, d in pipelines.items():
                    if week in d:
                        pipeline_vals.setdefault(pipe, {}).setdefault(level, []).append(d[week]['score'])

        out[week] = {}
        for pipe in pipeline_vals:
            out[week][pipe] = {}
            for level in pipeline_vals[pipe]:
                vals = pipeline_vals[pipe][level]
                out[week][pipe][level] = {
                    'p30': np.percentile(vals, 30) if vals else 0,
                    'p50': np.percentile(vals, 50) if vals else 0,
                    'p75': np.percentile(vals, 75) if vals else 0,
                    'p90': np.percentile(vals, 90) if vals else 0
                }
    return out


def qc2_subreason_analysis(df, user, weeks=None):
    try:
        dfu = df[df['auditor_login'] == user]
        if weeks:
            dfu = dfu[dfu['week'].isin(weeks)]
        out = {}
        for (pipe, week), grp in dfu.groupby(['usecase', 'week']):
            if isinstance(grp, int):
                continue

            # Modified error calculation
            error_mask = (
                    df['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                    ~df['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
            )
            incorrect = grp[error_mask]

            total_incorrect = len(incorrect) if hasattr(incorrect, '__len__') else 0
            if not total_incorrect:
                continue

            counts = incorrect.groupby('qc2_subreason')['volume'].sum().to_dict()
            total_volume = incorrect['volume'].sum()
            percentages = (incorrect.groupby('qc2_subreason')['volume'].sum() / total_volume * 100).round(1).to_dict()

            out.setdefault(pipe, {})[week] = {
                'counts': counts,
                'percentages': percentages,
                'total': total_incorrect
            }
        return out
    except Exception as e:
        print(f"Error in qc2_subreason_analysis: {str(e)}")
        return {}


def percentile_label(val, percentiles, level):
    if percentiles is None or val == 0:
        return "-"

    value = val['score'] if isinstance(val, dict) and 'score' in val else val

    if not percentiles or level not in percentiles:
        return "-"

    level_pcts = percentiles[level]
    if value >= level_pcts.get('p90', 0): return "P90 +"
    if value >= level_pcts.get('p75', 0): return "P75-P90"
    if value >= level_pcts.get('p50', 0): return "P50-P75"
    if value >= level_pcts.get('p30', 0): return "P30-P50"
    return "<P30"

def pct_bg_fg(bench):
    if bench == "P90 +":
        return "#00FF00", "black"  # Bright green
    elif bench == "P75-P90":
        return "#90EE90", "black"  # Light green
    elif bench == "P50-P75":
        return "#FFFF00", "black"  # Yellow
    elif bench == "P30-P50":
        return "#FFA500", "black"  # Orange
    elif bench == "<P30":
        return "#FF0000", "white"  # Red
    else:
        return "#ffffff", "black"  # Default white background


def html_metric_value_table_with_latest(self, data, weeks, percentiles, section="Quality", is_quality=False, month_data=None, user_level=None):
    styles = get_common_styles()
    if not weeks:
        weeks = []

    # Get month name
    month_num = self.month_selection.get()
    month_name = calendar.month_name[month_num]

    # Get all unique pipelines from both weekly and monthly data
    all_pipelines = set()
    for pipe in data:
        if pipe != "Overall":
            all_pipelines.add(pipe)
    if month_data:
        for pipe in month_data:
            if pipe != "Overall":
                all_pipelines.add(pipe)
    all_pipelines = sorted(all_pipelines)

    # Calculate column widths
    pipeline_width = 25
    data_column_width = 60 / len(weeks)
    month_width = 15

    # Create header cells
    ths = "".join(
        f"<th style='{styles['header_style']} width: {data_column_width}%;'>{w}</th>"
        for w in weeks
    )
    ths += f"<th style='{styles['header_style']} width: {month_width}%;'>Month({month_name})</th>"

    # Create rows
    rows = ""
    pipeline_volumes = {}
    pipeline_times = {}

    for pipe in all_pipelines:
        tds = ""
        # Process weekly data
        for w in weeks:
            if is_quality:
                val_dict = data.get(pipe, {}).get(w, {})
                val = val_dict.get('score', 0) if isinstance(val_dict, dict) else 0
                err = val_dict.get('err', 0) if isinstance(val_dict, dict) else 0
                total = val_dict.get('total', 0) if isinstance(val_dict, dict) else 0
                disp = f"{val:.1%}<br>({err}/{total})" if total else "-"
            else:
                if w not in pipeline_volumes:
                    pipeline_volumes[w] = 0
                    pipeline_times[w] = 0

                if pipe in data and w in data[pipe]:
                    vol = data[pipe][w] * 1
                    pipeline_volumes[w] += vol
                    pipeline_times[w] += 1

                val = data.get(pipe, {}).get(w, 0)
                disp = f"{val:.1f}" if val else "-"

            tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{disp}</td>"

        # Add monthly data
        if month_data and pipe in month_data:
            month_val = month_data[pipe]
            if isinstance(month_val, dict):
                score = month_val.get('score', 0)
                err = month_val.get('errors', 0)
                total = month_val.get('total', 0)
                month_disp = f"{score:.1%}<br>({err}/{total})" if total else "-"
            else:
                month_disp = f"{month_val:.1f}" if month_val else "-"
        else:
            month_disp = "-"
        tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{month_disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'>{pipe}</td>{tds}</tr>"

    # Add Overall row with calculations
    overall_tds = ""
    for w in weeks:
        if is_quality:
            total_volume = sum(data.get(pipe, {}).get(w, {}).get('total', 0) for pipe in all_pipelines)
            total_errors = sum(data.get(pipe, {}).get(w, {}).get('err', 0) for pipe in all_pipelines)
            overall_score = (total_volume - total_errors) / total_volume if total_volume else 0
            overall_disp = f"{overall_score:.1%}<br>({total_errors}/{total_volume})" if total_volume else "-"
        else:
            overall_prod = pipeline_volumes[w] / pipeline_times[w] if pipeline_times[w] > 0 else 0
            overall_disp = f"{overall_prod:.1f}" if overall_prod else "-"

        overall_tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{overall_disp}</td>"

    # Add monthly Overall
    if month_data and 'Overall' in month_data:
        monthly_overall = month_data['Overall']
        if isinstance(monthly_overall, dict):
            score = monthly_overall.get('score', 0)
            err = monthly_overall.get('errors', 0)
            total = monthly_overall.get('total', 0)
            monthly_disp = f"{score:.1%}<br>({err}/{total})" if total else "-"
        else:
            monthly_disp = f"{monthly_overall:.1f}" if monthly_overall else "-"
    else:
        monthly_disp = "-"
    overall_tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{monthly_disp}</td>"

    rows += f"""<tr style='border-top: 2px solid white;'>
        <td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'><strong>Overall</strong></td>
        {overall_tds}
    </tr>"""

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 2}" style="{styles['title_style']}">
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


def html_metric_pct_table(data, weeks, percentiles, section="Quality", user_level=None):
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
            bench = percentile_label(val_scalar, pctls, user_level) or "-"
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


from datetime import datetime, timedelta

def get_week_start_date(week, year):
    return datetime.strptime(f'{year}-W{int(week)}-1', "%Y-W%W-%w")


def calculate_monthly_productivity_metrics(df, month_num, user=None):
    try:
        df_monthly = df.copy()
        # Add debug statements here
        print("Available columns:", df_monthly.columns.tolist())
        print("Month values:", df_monthly['p_month'].unique() if 'p_month' in df_monthly.columns else "p_month not found")

        # Filter by user if specified
        if user:
            df_monthly = df_monthly[df_monthly['useralias'] == user]

        # Filter by month
        df_monthly = df_monthly[df_monthly['p_month'] == month_num]

        # Calculate total volume and time
        df_monthly['total_volume'] = df_monthly['processed_volume'] + df_monthly['processed_volumetr']
        df_monthly['total_time'] = df_monthly['processed_time'] + df_monthly['processed_time_tr']

        # Get all unique pipelines for the month
        all_pipelines = df_monthly['pipeline'].unique()

        # Calculate metrics for each pipeline
        monthly_metrics = {}
        for pipeline in all_pipelines:
            pipeline_data = df_monthly[df_monthly['pipeline'] == pipeline]
            total_volume = pipeline_data['total_volume'].sum()
            total_time = pipeline_data['total_time'].sum()
            productivity = total_volume / total_time if total_time > 0 else 0
            monthly_metrics[pipeline] = productivity

        # Calculate overall metrics
        total_volume = df_monthly['total_volume'].sum()
        total_time = df_monthly['total_time'].sum()
        monthly_metrics['Overall'] = total_volume / total_time if total_time > 0 else 0

        return monthly_metrics
    except Exception as e:
        print(f"Error in monthly productivity metrics: {str(e)}")
        return {}


from datetime import datetime

def calculate_monthly_quality_metrics(df, month_num, user=None):
    try:
        df_monthly = df.copy()

        # Filter by user
        if user:
            df_monthly = df_monthly[df_monthly['auditor_login'] == user]

        # Filter by month only since month column exists
        df_monthly = df_monthly[df_monthly['month'] == month_num]

        if df_monthly.empty:
            return {}

        result = {}
        for usecase, grp in df_monthly.groupby('usecase'):
            total_volume = grp['volume'].sum()

            # Apply error mask within group to avoid index mismatch
            error_mask = (
                grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                ~grp['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
            )
            errors = grp.loc[error_mask, 'volume'].sum()

            result[usecase] = {
                'score': (total_volume - errors) / total_volume if total_volume > 0 else 0,
                'errors': int(errors),
                'total': int(total_volume)
            }

        # Calculate overall metrics
        total_volume = df_monthly['volume'].sum()
        error_mask = (
            df_monthly['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
            ~df_monthly['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
        )
        total_errors = df_monthly.loc[error_mask, 'volume'].sum()

        result['Overall'] = {
            'score': (total_volume - total_errors) / total_volume if total_volume > 0 else 0,
            'errors': int(total_errors),
            'total': int(total_volume)
        }

        return result

    except Exception as e:
        print(f"[Monthly Metric Error] {e}")
        return {}


def calculate_monthly_subreason_metrics(df, month_num, user=None):
    try:
        df = df.copy()

        # Add debug prints
        print("\nDebugging monthly_subreason_metrics:")
        print(f"Initial data shape: {df.shape}")
        print(f"Month number: {month_num}")
        print(f"User: {user}")

        # Filter by user if specified
        if user:
            df = df[df['auditor_login'] == user]
            print(f"After user filter shape: {df.shape}")

        # Filter by month only
        df = df[df['month'] == month_num]
        print(f"After month filter shape: {df.shape}")
        print(f"Unique months in data: {df['month'].unique()}")

        incorrect_data = df[
            (df['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT'])) &
            ~(df['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct']))
            ]
        print(f"Incorrect data shape: {incorrect_data.shape}")

        monthly_metrics = {}
        if not incorrect_data.empty:
            subreason_counts = incorrect_data.groupby('qc2_subreason')['volume'].sum()
            print(f"Subreason counts:\n{subreason_counts}")

            for subreason, volume in subreason_counts.items():
                if pd.notna(subreason):
                    monthly_metrics[subreason] = {
                        'count': int(volume),
                        'percentage': (volume / incorrect_data['volume'].sum() * 100) if incorrect_data[
                                                                                             'volume'].sum() > 0 else 0
                    }

        print(f"Final monthly metrics: {monthly_metrics}")
        return monthly_metrics

    except Exception as e:
        print(f"Error in monthly subreason metrics: {str(e)}")
        return {}


def calculate_monthly_correction_metrics(df, month_num, user=None):
    try:
        df = df.copy()

        # Filter by user if specified
        if user:
            df = df[df['auditor_login'] == user]

        # Filter by month only
        df = df[df['month'] == month_num]

        monthly_metrics = {}

        for correction_type, grp in df.groupby('auditor_correction_type'):
            total_volume = grp['volume'].sum()

            error_mask = (
                df['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                ~df['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
            )
            errors = grp.loc[error_mask, 'volume'].sum() if hasattr(grp, '__len__') else 0

            monthly_metrics[correction_type] = {
                'score': (total_volume - errors) / total_volume if total_volume > 0 else 0,
                'errors': errors,
                'total': total_volume
            }

        # Calculate overall metrics
        total_volume = df['volume'].sum()
        error_mask = (
            df['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
            ~df['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
        )
        total_errors = df[error_mask]['volume'].sum()

        monthly_metrics['Overall'] = {
            'score': (total_volume - total_errors) / total_volume if total_volume > 0 else 0,
            'errors': total_errors,
            'total': total_volume
        }

        return monthly_metrics
    except Exception as e:
        print(f"Error in monthly correction metrics: {str(e)}")
        return {}


def calculate_correction_type_data(df, user, weeks=None):
    """Calculate correction type quality data"""
    try:
        dfu = df[df['auditor_login'] == user]
        if weeks:
            dfu = dfu[dfu['week'].isin(weeks)]

        correction_data = {}
        for (correction_type, week), grp in dfu.groupby(['auditor_correction_type', 'week']):
            if isinstance(grp, int):
                continue

            total_volume = grp['volume'].sum()
            errors = grp[
                (grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT'])) &
                ~(grp['auditor_reappeal_final_judgement'].isin(['CORRECT', 'BOTH CORRECT', 'GL', 'Auditor Correct']))
                ]['volume'].sum()

            score = (total_volume - errors) / total_volume if total_volume > 0 else 0

            correction_data.setdefault(correction_type, {})[week] = {
                'score': score,
                'err': errors,
                'total': total_volume
            }

        return correction_data

    except Exception as e:
        print(f"Error in calculate_correction_type_data: {str(e)}")
        return {}


def html_qc2_reason_value_table(self, subreason_data, weeks, month_data=None):
    # Add debug prints
    print("\nDebugging html_qc2_reason_value_table:")
    print(f"Subreason data: {subreason_data}")
    print(f"Weeks: {weeks}")
    print(f"Month data: {month_data}")

    styles = get_common_styles()
    if not weeks:
        weeks = []
    if not subreason_data:
        subreason_data = {}

    # Get month name
    month_num = self.month_selection.get()
    month_name = calendar.month_name[month_num]
    print(f"Month number: {month_num}, Month name: {month_name}")

    # Calculate column widths
    subreason_width = 30
    data_column_width = (70) / (len(weeks) + 1)

    # Create header cells
    ths = "".join(
        f"<th style='{styles['header_style']} width: {data_column_width}%;'>{w}</th>"
        for w in weeks
    )
    ths += f"<th style='{styles['header_style']} width: {data_column_width}%;'>Month({month_name})</th>"

    # Get all unique subreasons
    all_subs = set()
    for pipe in subreason_data:
        for w in subreason_data[pipe]:
            all_subs.update(subreason_data[pipe][w]['counts'].keys())
    if month_data:
        all_subs.update(month_data.keys())
    all_subs = sorted(all_subs)
    print(f"All subreasons: {all_subs}")

    # Debug monthly data processing
    if month_data:
        print("\nProcessing monthly data:")
        for sub in month_data:
            print(f"Subreason: {sub}")
            print(f"Data: {month_data[sub]}")

    # Create rows
    rows = ""
    for sub in all_subs:
        print(f"\nProcessing subreason: {sub}")
        tds = ""
        # Weekly data
        for w in weeks:
            count = sum(subreason_data[pipe][w]['counts'].get(sub, 0)
                        for pipe in subreason_data if w in subreason_data[pipe])
            total = sum(subreason_data[pipe][w]['total']
                        for pipe in subreason_data if w in subreason_data[pipe])
            percentage = (count/total * 100) if total else 0
            disp = f"{count}/{total}<br>({percentage:.1f}%)" if count and total else "-"
            print(f"Week {w}: count={count}, total={total}, percentage={percentage:.1f}%")
            tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{disp}</td>"

        # Monthly data
        if month_data and sub in month_data:
            month_count = month_data[sub]['count']
            month_percentage = month_data[sub]['percentage']
            total_count = sum(data['count'] for data in month_data.values())
            month_disp = f"{month_count}/{total_count}<br>({month_percentage:.1f}%)"
            print(f"Month: count={month_count}, total={total_count}, percentage={month_percentage:.1f}%")
        else:
            month_disp = "-"
            print("No monthly data for this subreason")
        tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{month_disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {subreason_width}%;'>{sub}</td>{tds}</tr>"

    print("\nFinished generating table")
    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 2}" style="{styles['title_style']}">
                    QC2 Subreason Counts
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
        """Initialize the Reporter App"""
        super().__init__()

        # Configure main window
        self.title("📊 Email Report Generator")
        self.geometry("1920x1200")
        self.state('zoomed')
        self.configure(bg='#121212')

        # Initialize UI variables
        self.data_source = tk.StringVar(value="excel")  # For data source selection
        self.data = tk.StringVar()  # For Excel file path
        self.qual = tk.StringVar()  # For Quality data
        self.mgr = tk.StringVar()  # For Manager mapping
        self.timesheet = tk.StringVar()  # Add this line for timesheet
        self.wks = tk.StringVar()  # For weeks input
        self.user = tk.StringVar()  # For user login
        self.mode = tk.StringVar(value="preview")  # For email mode (preview/send)
        self.month_selection = tk.StringVar(value="current")  # For month selection

        # Initialize data caches
        self.productivity_cache = {}
        self.quality_cache = {}
        self.allprod_cache = None
        self.allqual_cache = None
        self.prod_pct_cache = {}
        self.qual_pct_cache = {}
        self.monthly_prod_cache = None
        self.monthly_qual_cache = None

        # Initialize UI components
        self.status = None  # For status text widget
        self.report_type = None  # For report type combobox

        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('alt')

        # Configure Dark Material Design colors
        self.colors = {
            'primary': '#bb86fc',  # Purple 200
            'primary_dark': '#985eff',  # Purple 300
            'primary_light': '#d0bcff',  # Purple 100
            'secondary': '#03dac6',  # Teal 200
            'surface': '#1e1e1e',  # Dark surface
            'background': '#121212',  # Dark background
            'card_background': '#1e1e1e',  # Card background
            'on_surface': '#e1e1e1',  # Light text
            'on_surface_variant': '#c1c1c1',  # Medium light text
            'error': '#cf6679',  # Red 300
            'success': '#4caf50',  # Green 500
            'warning': '#ffb74d',  # Orange 300
            'outline': '#404040'  # Border color
        }

        # Configure percentile colors with your specified color scheme
        self.percentile_colors = {
            "P90 +": {"bg": "#00FF00", "fg": "black"},  # Bright green
            "P75-P90": {"bg": "#90EE90", "fg": "black"},  # Light green
            "P50-P75": {"bg": "#FFFF00", "fg": "black"},  # Yellow
            "P30-P50": {"bg": "#FFA500", "fg": "black"},  # Orange
            "<P30": {"bg": "#FF0000", "fg": "white"}  # Red
        }

        # Configure styles
        self._configure_styles()

        # Build UI
        self._build_ui()

        # Force refresh after initialization
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
        """Load data from Excel file"""
        try:
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

        # Bind resize event
        self.bind("<Configure>", self._on_resize)

        # Update UI
        self.update_idletasks()

        # Initial UI refresh
        self.after(100, lambda: self._force_refresh())

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
            content, "📊 Excel File (with Volume, Quality, TimeSheet sheets)",
            self.data,
            command=lambda: self._browse(self.data)
        )

        # Initialize other frames but don't pack them
        self.qual_frame = tk.Frame(content)
        self.mgr_frame = tk.Frame(content)

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
        # Replace the month selection radio buttons with an input field
        month_frame = tk.Frame(content, bg=self.colors['card_background'])
        month_frame.pack(fill="x", pady=(24, 0))

        tk.Label(
            month_frame,
            text="📅 Month Number (1-12)",
            bg=self.colors['card_background'],
            fg=self.colors['on_surface_variant'],
            font=('Segoe UI', 11, 'bold')
        ).pack(anchor='w', pady=(0, 12))

        month_entry_frame = tk.Frame(
            month_frame,
            bg=self.colors['surface_variant'],
            relief='solid',
            bd=1
        )
        month_entry_frame.pack(fill='x')

        # Change month_selection from StringVar to IntVar
        self.month_selection = tk.IntVar(value=datetime.now().month)

        tk.Entry(
            month_entry_frame,
            textvariable=self.month_selection,
            font=('Segoe UI', 10),
            bg=self.colors['surface_variant'],
            fg=self.colors['on_surface'],
            relief='flat',
            bd=0
        ).pack(fill='x', padx=12, pady=8)

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
                                            month_data=None, user_level=None):
        styles = get_common_styles()
        if not weeks:
            weeks = []

        # Get month name
        current_date = datetime.now()
        if self.month_selection.get() == "current":
            month_name = calendar.month_name[current_date.month]
        else:
            previous_month = 12 if current_date.month == 1 else current_date.month - 1
            month_name = calendar.month_name[previous_month]

        # Calculate column widths - removed Latest column width
        pipeline_width = 25
        data_column_width = 60 / len(weeks)  # Distribute 60% among week columns
        month_width = 15

        # Create header cells - removed Latest column
        ths = "".join(
            f"<th style='{styles['header_style']} width: {data_column_width}%;'>{w}</th>"
            for w in weeks
        )
        ths += f"<th style='{styles['header_style']} width: {month_width}%;'>Month({month_name})</th>"

        # Create rows
        rows = ""
        pipeline_volumes = {}  # Store volumes for productivity calculation
        pipeline_times = {}  # Store times for productivity calculation

        for pipe in sorted(data):
            if pipe == "Overall":  # Skip Overall here, we'll add it later
                continue

            tds = ""
            for w in weeks:
                if is_quality:
                    val_dict = data[pipe].get(w, {})
                    val = val_dict.get('score', 0) if isinstance(val_dict, dict) else 0
                    err = val_dict.get('err', 0) if isinstance(val_dict, dict) else 0
                    total = val_dict.get('total', 0) if isinstance(val_dict, dict) else 0
                    disp = f"{val:.1%}<br>({err}/{total})" if total else "-"
                else:
                    # For productivity, store volumes and times for later calculation
                    if w not in pipeline_volumes:
                        pipeline_volumes[w] = 0
                        pipeline_times[w] = 0

                    if w in data[pipe]:
                        vol = data[pipe][w] * 1  # Assuming time unit is 1
                        pipeline_volumes[w] += vol
                        pipeline_times[w] += 1

                    val = data[pipe].get(w, 0)
                    disp = f"{val:.1f}" if val else "-"

                tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{disp}</td>"

            # Add monthly data
            if month_data and pipe in month_data:
                month_val = month_data[pipe]
                if isinstance(month_val, dict):
                    score = month_val.get('score', 0)
                    err = month_val.get('errors', 0)
                    total = month_val.get('total', 0)
                    month_disp = f"{score:.1%}<br>({err}/{total})" if total else "-"
                else:
                    month_disp = f"{month_val:.1%}" if is_quality else f"{month_val:.1f}"
            else:
                month_disp = "-"
            tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{month_disp}</td>"

            rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'>{pipe}</td>{tds}</tr>"

        # Add Overall row with corrected calculation
        overall_tds = ""
        for w in weeks:
            if is_quality:
                total_volume = sum(data[pipe].get(w, {}).get('total', 0) for pipe in data if pipe != "Overall")
                total_errors = sum(data[pipe].get(w, {}).get('err', 0) for pipe in data if pipe != "Overall")
                overall_score = (total_volume - total_errors) / total_volume if total_volume else 0
                overall_disp = f"{overall_score:.1%}<br>({total_errors}/{total_volume})" if total_volume else "-"
            else:
                # Calculate overall productivity using stored volumes and times
                overall_prod = pipeline_volumes[w] / pipeline_times[w] if pipeline_times[w] > 0 else 0
                overall_disp = f"{overall_prod:.1f}" if overall_prod else "-"

            overall_tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{overall_disp}</td>"

        # Add monthly Overall
        if month_data and 'Overall' in month_data:
            monthly_overall = month_data['Overall']
            if isinstance(monthly_overall, dict):
                monthly_overall = monthly_overall.get('score', 0)
            monthly_disp = f"{monthly_overall:.1%}" if is_quality else f"{monthly_overall:.1f}"
        else:
            monthly_disp = "-"
        overall_tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{monthly_disp}</td>"

        # Add Overall row to table
        rows += f"""<tr style='border-top: 2px solid white;'>
            <td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'><strong>Overall</strong></td>
            {overall_tds}
        </tr>"""

        return f"""
        <div style="{styles['container_style']}">
            <table style="{styles['table_style']}">
                <tr>
                    <td colspan="{len(weeks) + 2}" style="{styles['title_style']}">
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

    def html_qc2_reason_value_table(self, subreason_data, weeks, month_data=None):
        """Generate HTML table for QC2 subreason counts with adjusted column widths"""
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
        subreason_width = 30  # Reduced width for subreason column
        data_column_width = (70) / (
                    len(weeks) + 1)  # Evenly distribute remaining width among week columns and month column

        # Modify styles for adjusted widths
        subreason_cell_style = f'''
            {styles['pipeline_cell_style']}
            width: {subreason_width}%;
            max-width: {subreason_width}%;
            white-space: normal;
            word-wrap: break-word;
        '''

        data_cell_style = f'''
            {styles['cell_style']}
            width: {data_column_width}%;
            max-width: {data_column_width}%;
        '''

        # Create header cells
        ths = "".join(
            f"<th style='{styles['header_style']} width: {data_column_width}%;'>{w}</th>"
            for w in weeks
        )
        ths += f"<th style='{styles['header_style']} width: {data_column_width}%;'>Month({month_name})</th>"

        # Create rows
        rows = ""
        all_subs = sorted({sub for pipe in subreason_data for w in subreason_data[pipe]
                           for sub in subreason_data[pipe][w]['counts']})

        for sub in all_subs:
            tds = ""
            for w in weeks:
                count = sum(subreason_data[pipe][w]['counts'].get(sub, 0)
                            for pipe in subreason_data if w in subreason_data[pipe])
                tds += f"<td style='{data_cell_style}'>{count if count else '-'}</td>"

            # Add monthly data
            if month_data:
                month_count = sum(pipe_data.get(sub, {}).get('count', 0)
                                  for pipe_data in month_data.values())
                tds += f"<td style='{data_cell_style}'>{month_count if month_count else '-'}</td>"
            else:
                tds += f"<td style='{data_cell_style}'>-</td>"

            rows += f"<tr><td style='{subreason_cell_style}'>{sub}</td>{tds}</tr>"

        return f"""
        <div style="{styles['container_style']}">
            <table style="{styles['table_style']}">
                <tr>
                    <td colspan="{len(weeks) + 2}" style="{styles['title_style']}">
                        QC2 Subreason Counts
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
        from datetime import datetime
        import calendar

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
                self.data_frame.pack(fill="x", pady=(0, 20))  # Changed from timesheet_frame to data_frame
                self.mgr_frame.pack(fill="x", pady=(0, 20))
                self.info_label.config(
                    text="⏰ Timesheet Missing Report requires Data Excel and Manager Mapping files")

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
        if not user or not self.data.get():  # Changed from self.timesheet.get()
            self.log("❌ User and data file required.")
            return

        weeks = self._weeks()
        if not weeks:
            return

        try:
            # Get timesheet data for user
            timesheet_data = analyze_timesheet_missing(self.data.get(), weeks[-1], user)

            # If no timesheet missing data, log and return without sending
            if timesheet_data is None or timesheet_data.empty:
                self.log(f"ℹ️ No timesheet missing data found for {user}")
                return

            # Continue with sending report only if there is data
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

            send_mail_html(f"{user}@amazon.com", "", "Timesheet Missing Report", html,
                           preview=self.mode.get() == "preview")

            mode_text = "previewed" if self.mode.get() == "preview" else "sent"
            self.log(f"✅ Timesheet mail for {user} {mode_text} successfully!")

        except Exception as exc:
            self.log(f"❌ Error generating timesheet report: {str(exc)}")

    def send_timesheet_bulk(self):
        """Send timesheet reports in bulk to multiple users"""
        if not self.data.get() or not self.mgr.get():
            self.log("❌ Data file and manager map required.")
            return

        weeks = self._weeks()
        if not weeks:
            return

        try:
            mgr = pd.read_excel(self.mgr.get())
            self.log(f"⏰ Processing timesheet reports...")

            # Track users with data
            users_with_data = 0

            for user in mgr['loginname'].dropna().unique():
                try:
                    # Get timesheet data for user
                    timesheet_data = analyze_timesheet_missing(self.data.get(), weeks[-1], user)

                    # Skip users with no timesheet missing data
                    if timesheet_data is None or timesheet_data.empty:
                        self.log(f"ℹ️ No timesheet data found for {user}")
                        continue

                    users_with_data += 1
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

                    send_mail_html(
                        to=f"{user}@amazon.com",
                        cc=cc,
                        subject="Timesheet Missing Report",
                        html=html,
                        preview=self.mode.get() == "preview"
                    )

                    mode_text = "previewed" if self.mode.get() == "preview" else "sent"
                    self.log(f"✅ Timesheet mail for {user} {mode_text}")

                except Exception as user_exc:
                    self.log(f"⚠️ Error processing user {user}: {str(user_exc)}")
                    continue

            self.log(
                f"🎉 Bulk timesheet processing completed! Reports sent to {users_with_data} users with missing timesheet data.")

        except Exception as exc:
            self.log(f"❌ Error processing bulk timesheet: {str(exc)}")

    def send_single(self):
        """Send single user report based on selected report type"""
        report_type = self.report_type.get()
        user = self.user.get().strip()

        if not user:
            return self.log("❌ User required.")

        if not self.data.get():
            return self.log("❌ Data file required.")

        if report_type == "RDR Report":
            self.send_rdr_single()
        elif report_type == "Precision Correction Report":
            self.send_precision_single()
        elif report_type == "Timesheet Missing Report":
            self.send_timesheet_single()

    def send_bulk(self):
        """Send bulk reports based on selected report type"""
        report_type = self.report_type.get()

        if not self.mgr.get():
            return self.log("❌ Manager mapping required.")

        if not self.data.get():
            return self.log("❌ Data file required.")

        if report_type == "RDR Report":
            self.send_rdr_bulk()
        elif report_type == "Precision Correction Report":
            self.send_precision_bulk()
        elif report_type == "Timesheet Missing Report":
            self.send_timesheet_bulk()

    def send_rdr_single(self):
        """Send single user report with detailed debugging"""
        try:
            user = self.user.get().strip()
            if not user:
                return self.log("❌ User required.")

            weeks = self._weeks()
            if not weeks:
                return

            # Load data and get auditor levels
            self.log(f"\nProcessing data for {user}...")
            df, auditor_levels = self.load_data()
            if df is None:
                return

            # Get user's level from first entry
            user_data = df[df['useralias'] == user]
            if user_data.empty:
                self.log(f"❌ No data found for user {user}")
                return

            user_level = user_data['level'].iloc[0]
            self.log(f"User level: {user_level}")

            # Calculate productivity metrics
            prod = user_prod_dict(df, user, weeks)
            allprod = all_prod_dict(df, weeks)
            month_num = self.month_selection.get()

            # Get levels for all auditors
            auditor_levels = df.groupby('useralias')['level'].first().to_dict()
            prod_pct = productivity_percentiles(allprod, weeks, auditor_levels)

            # Calculate monthly productivity metrics
            monthly_prod_metrics = calculate_monthly_productivity_metrics(df, self.month_selection.get())

            # Initialize quality tables
            qual_table = qual_pct_table = qc2_left = qc2_right = ""
            correction_quality_table = correction_count_table = ""

            # Load quality data
            self.log("\nProcessing quality data...")
            try:
                qdf = load_quality(self.data.get())
                if not qdf.empty:
                    # Calculate quality metrics
                    qual = user_quality_dict(qdf, user, weeks)
                    allqual = all_quality_dict(qdf, weeks)
                    qual_pct = quality_percentiles(allqual, weeks, auditor_levels)

                    # Calculate monthly metrics
                    monthly_qual_metrics = calculate_monthly_quality_metrics(qdf, self.month_selection.get(), user=user)
                    monthly_subreason = calculate_monthly_subreason_metrics(qdf, month_num, user=user)
                    monthly_correction = calculate_monthly_correction_metrics(qdf, self.month_selection.get(), user=user)

                    # Generate quality tables
                    qual_table = self.html_metric_value_table_with_latest(
                        qual, weeks, qual_pct, section="Quality",
                        is_quality=True, month_data=monthly_qual_metrics,
                        user_level=user_level)  # Pass user_level here

                    qual_pct_table = html_metric_pct_table(
                        qual, weeks, qual_pct, section="Quality",
                        user_level=user_level)  # Pass user_level here

                    # QC2 analysis
                    subreason = qc2_subreason_analysis(qdf, user, weeks)
                    qc2_left = self.html_qc2_reason_value_table(subreason, weeks, month_data=monthly_subreason)
                    qc2_right = html_qc2_reason_pct_table(subreason, weeks)

                    # Correction analysis
                    correction_data = calculate_correction_type_data(qdf, user, weeks)
                    correction_quality_table = self.html_correction_type_quality_table(
                        correction_data, weeks, month_data=monthly_correction)
                    correction_count_table = html_correction_type_count_table(
                        correction_data, weeks)

            except Exception as qe:
                self.log(f"⚠️ Error processing quality data: {str(qe)}")

            # Process timesheet data
            timesheet_table = ""
            try:
                timesheet_data = analyze_timesheet_missing(self.data.get(), weeks[-1], user)
                if timesheet_data is not None and not timesheet_data.empty:
                    timesheet_table = html_timesheet_missing_table(timesheet_data)
            except Exception as te:
                self.log(f"⚠️ Error processing timesheet: {str(te)}")

            # Generate productivity tables
            prod_table = self.html_metric_value_table_with_latest(
                prod, weeks, prod_pct, section="Productivity",
                is_quality=False, month_data=monthly_prod_metrics, user_level=user_level)

            prod_pct_table = html_metric_pct_table(
                prod, weeks, prod_pct, section="Productivity", user_level=user_level)

            # Compose final HTML
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

    def send_rdr_bulk(self):
        """Send bulk reports to multiple users"""
        if not self.mgr.get():
            return self.log("❌ Manager mapping required.")

        weeks = self._weeks()
        if not weeks:
            return

        try:
            # Load data
            df, auditor_levels = self.load_data()
            if df is None:
                return

            # Get auditor levels
            auditor_levels = df.groupby('useralias')['level'].first().to_dict()

            # Calculate monthly productivity metrics
            monthly_prod_metrics = calculate_monthly_productivity_metrics(df, self.month_selection.get())

            # Read manager mapping
            mgr = pd.read_excel(self.mgr.get())

            # Calculate all productivity metrics
            allprod = all_prod_dict(df, weeks)
            prod_pct = productivity_percentiles(allprod, weeks)

            # Load quality data if available
            qdf = None
            allqual = qual_pct = {}
            monthly_qual_metrics = monthly_subreason = monthly_correction = None

            try:
                qdf = load_quality(self.data.get())
                if not qdf.empty:
                    allqual = all_quality_dict(qdf, weeks)
                    qual_pct = quality_percentiles(allqual, weeks)
                    monthly_qual_metrics = calculate_monthly_quality_metrics(qdf, self.month_selection.get(), user=user)
                    monthly_subreason = calculate_monthly_subreason_metrics(qdf, self.month_selection.get(), user=user)
                    monthly_correction = calculate_monthly_correction_metrics(qdf, self.month_selection.get(), user=user)
            except Exception as qe:
                self.log(f"⚠️ Error loading quality data: {str(qe)}")

            # Process each user
            for user in mgr['loginname'].dropna().unique():
                try:
                    # Get user's level
                    user_level = auditor_levels.get(user)
                    if user_level is None:
                        self.log(f"⚠️ Skipping user {user} - no level found")
                        continue

                    self.log(f"Processing {user} (Level {user_level})...")

                    # Generate productivity tables
                    prod = user_prod_dict(df, user, weeks)
                    prod_table = self.html_metric_value_table_with_latest(
                        prod, weeks, prod_pct, section="Productivity",
                        is_quality=False, month_data=monthly_prod_metrics, user_level=user_level)

                    prod_pct_table = html_metric_pct_table(
                        prod, weeks, prod_pct, section="Productivity", user_level=user_level)

                    # Initialize quality tables
                    qual_table = qual_pct_table = qc2_left = qc2_right = ""
                    correction_quality_table = correction_count_table = ""

                    # Generate quality tables if data exists
                    if qdf is not None and not qdf.empty:
                        qual = allqual.get(user, {})
                        qual_table = self.html_metric_value_table_with_latest(
                            qual, weeks, qual_pct, section="Quality",
                            is_quality=True, month_data=monthly_qual_metrics, user_level=user_level)

                        qual_pct_table = html_metric_pct_table(
                            qual, weeks, qual_pct, section="Quality")

                        # QC2 analysis
                        subreason = qc2_subreason_analysis(qdf, user, weeks)
                        qc2_left = self.html_qc2_reason_value_table(
                            subreason, weeks, month_data=monthly_subreason)
                        qc2_right = html_qc2_reason_pct_table(subreason, weeks)

                        # Correction analysis
                        correction_data = calculate_correction_type_data(qdf, user, weeks)
                        correction_quality_table = self.html_correction_type_quality_table(
                            correction_data, weeks, month_data=monthly_correction)
                        correction_count_table = html_correction_type_count_table(
                            correction_data, weeks)

                    # Process timesheet data
                    timesheet_table = ""
                    timesheet_data = analyze_timesheet_missing(self.data.get(), weeks[-1], user)
                    if timesheet_data is not None and not timesheet_data.empty:
                        timesheet_table = html_timesheet_missing_table(timesheet_data)

                    # Compose HTML
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

                    # Get CC email if supervisor exists
                    cc = ""
                    if 'supervisorloginname' in mgr.columns:
                        cc_val = mgr[mgr['loginname'] == user]['supervisorloginname'].iloc[0]
                        if pd.notna(cc_val):
                            cc = f"{cc_val}@amazon.com"

                    # Send email
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
    app = ReporterApp()
    app.mainloop()