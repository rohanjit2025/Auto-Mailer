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
    try:
        if weeks:
            df = df[df['p_week'].isin(weeks)]

        # First get total precision corrections by user and week, handling NaN values
        df['precision_correction'] = pd.to_numeric(df['precision_correction'], errors='coerce').fillna(0)

        total_precision = df.groupby(['useralias', 'p_week'])['precision_correction'].sum().reset_index()
        high_precision_users = total_precision[total_precision['precision_correction'] > 20]

        # Then get ACL breakdown for users with high total precision
        result = {}
        for _, row in high_precision_users.iterrows():
            user = row['useralias']
            week = row['p_week']

            # Get ACL breakdown for this user and week
            acl_breakdown = df[
                (df['useralias'] == user) &
                (df['p_week'] == week)
                ].groupby('acl')['precision_correction'].sum().reset_index()

            # Store all ACL data for this user and week
            result.setdefault(user, {}).setdefault(week, [])
            for _, acl_row in acl_breakdown.iterrows():
                if acl_row['precision_correction'] > 0:  # Only include ACLs with corrections
                    result[user][week].append({
                        'value': float(acl_row['precision_correction']),  # Ensure float conversion
                        'acl': acl_row['acl']
                    })

        return result
    except Exception as e:
        print(f"Error in calculate_precision_corrections: {str(e)}")
        return {}


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
        # Filter for the specific user
        dfu = df[df['useralias'] == user]
        if weeks:
            dfu = dfu[dfu['p_week'].isin(weeks)]

        # Calculate totals
        dfu['total_volume'] = dfu['processed_volume'] + dfu['processed_volumetr']
        dfu['total_time'] = dfu['processed_time'] + dfu['processed_time_tr']

        # Group by pipeline and week
        g = dfu.groupby(['pipeline', 'p_week'], as_index=False).agg({
            'total_volume': 'sum',
            'total_time': 'sum',
            'level': 'first'
        })

        # Create result dictionary with volume and time
        result = {}
        for _, row in g.iterrows():
            result.setdefault(row['pipeline'], {})[row['p_week']] = {
                'volume': row['total_volume'],
                'time': row['total_time']
            }

        return result
    except Exception as e:
        print(f"Error in user_prod_dict: {str(e)}")
        return {}


def all_prod_dict(df, weeks=None):
    try:
        # Create a copy of the dataframe to avoid SettingWithCopyWarning
        df = df.copy()

        if weeks:
            df = df[df['p_week'].isin(weeks)]

        # Calculate totals using loc
        df.loc[:, 'total_volume'] = df['processed_volume'] + df['processed_volumetr']
        df.loc[:, 'total_time'] = df['processed_time'] + df['processed_time_tr']

        # Group data
        g = df.groupby(['useralias', 'pipeline', 'p_week'], as_index=False)[['total_volume', 'total_time']].sum()
        # Create dictionary
        d = {}
        for row in g.itertuples(index=False):
            d.setdefault(row.useralias, {}).setdefault(row.pipeline, {})[row.p_week] = {
                'volume': row.total_volume,
                'time': row.total_time
            }
        return d
    except Exception as e:
        print(f"Error in all_prod_dict: {str(e)}")
        return {}


def productivity_percentiles(allprod, weeks, auditor_levels):
    out = {}
    try:
        # First collect all unique pipelines across all users
        all_pipelines = set()
        for auditor, pipelines in allprod.items():
            all_pipelines.update(pipelines.keys())

        for week in weeks:
            pipeline_vals = {}
            for auditor, pipelines in allprod.items():
                level = auditor_levels.get(auditor)
                if level:  # Only include users with known levels
                    # Initialize all pipelines for this level
                    for pipe in all_pipelines:
                        pipeline_vals.setdefault(pipe, {}).setdefault(level, [])

                    # Add values where they exist
                    for pipe, values in pipelines.items():
                        if week in values:
                            vol = values[week].get('volume', 0)
                            time = values[week].get('time', 0)
                            if time > 0:  # Avoid division by zero
                                prod = vol / time
                                pipeline_vals[pipe][level].append(prod)

            out[week] = {}
            for pipe in all_pipelines:  # Use all_pipelines instead of pipeline_vals
                out[week][pipe] = {}
                for level in set(auditor_levels.values()):  # Use all levels
                    vals = pipeline_vals.get(pipe, {}).get(level, [])
                    if vals:  # Only calculate if there are values
                        out[week][pipe][level] = {
                            'p30': np.percentile(vals, 30),
                            'p50': np.percentile(vals, 50),
                            'p75': np.percentile(vals, 75),
                            'p90': np.percentile(vals, 90)
                        }
    except Exception as e:
        print(f"Error in productivity_percentiles: {str(e)}")
    return out


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

                # ✅ Fixed: use subgrp for error filtering
                error_mask = (
                    subgrp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                    ~subgrp['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
                )
                errors = subgrp.loc[error_mask, 'volume'].sum() if hasattr(subgrp, '__len__') else 0

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
        # Create a fresh copy of the dataframe for this user
        dfu = df[df['auditor_login'] == user].copy()
        if weeks:
            dfu = dfu[dfu['week'].isin(weeks)]

        d = {}
        print(f"\nProcessing quality data for user: {user}")

        # Process each pipeline-week combination separately
        for (pipe, week), pipeline_group in dfu.groupby(['usecase', 'week']):
            print(f"\nPipeline: {pipe}")
            print(f"Week: {week}")

            # Calculate metrics for this specific pipeline
            total_volume = pipeline_group['volume'].sum()

            # Create error mask specific to this pipeline group
            pipeline_error_mask = (
                    pipeline_group['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                    ~pipeline_group['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
            )

            # Calculate errors only for this pipeline
            pipeline_errors = pipeline_group.loc[pipeline_error_mask, 'volume'].sum()

            print(f"Total Volume: {total_volume}")
            print(f"Error Count: {pipeline_errors}")

            if total_volume > 0:
                score = (total_volume - pipeline_errors) / total_volume
                print(f"Calculated Score: {score:.4f}")

                d.setdefault(pipe, {})[week] = {
                    'score': score,
                    'err': pipeline_errors,
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
        print(f"\n=== Processing Week {week} ===")
        out[week] = {}
        week_data = {}

        for auditor, pipelines in allqual.items():
            level = auditor_levels.get(auditor)
            if not level:
                continue

            for pipe, pipe_data in pipelines.items():
                if week in pipe_data:
                    week_val = pipe_data[week]
                    total = week_val.get('total', 0)
                    errors = week_val.get('err', 0)

                    # Calculate score the same way as html_metric_value_table_with_latest
                    if total > 0:
                        score = (total - errors) / total
                        print(f"\nProcessing {pipe} for auditor {auditor}:")
                        print(f"Total: {total}, Errors: {errors}")
                        print(f"Calculated Score: {score:.4f}")
                        week_data.setdefault(pipe, {}).setdefault(level, []).append(score)

        for pipe in week_data:
            print(f"\nPipeline: {pipe}")
            out[week][pipe] = {}

            for level in week_data[pipe]:
                scores = week_data[pipe][level]
                if scores:
                    try:
                        percentiles = {
                            'p30': np.percentile(scores, 30),
                            'p50': np.percentile(scores, 50),
                            'p75': np.percentile(scores, 75),
                            'p90': np.percentile(scores, 90)
                        }
                        out[week][pipe][level] = percentiles
                        print(f"Level {level} percentiles:")
                        print(f"P30: {percentiles['p30']:.4f}")
                        print(f"P50: {percentiles['p50']:.4f}")
                        print(f"P75: {percentiles['p75']:.4f}")
                        print(f"P90: {percentiles['p90']:.4f}")
                    except Exception as e:
                        print(f"Error calculating percentiles: {e}")
                        continue

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
                    grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                    ~grp['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
            )
            incorrect = grp[error_mask]

            total_errors = incorrect['volume'].sum()  # Total error volume
            if not total_errors:
                continue

            counts = incorrect.groupby('qc2_subreason')['volume'].sum().to_dict()
            # Calculate percentages based on total errors instead of total volume
            percentages = {k: (v/total_errors * 100) for k, v in counts.items()}

            out.setdefault(pipe, {})[week] = {
                'counts': counts,
                'percentages': percentages,
                'total': total_errors
            }
        return out
    except Exception as e:
        print(f"Error in qc2_subreason_analysis: {str(e)}")
        return {}



def html_metric_value_table_with_latest(self, data, weeks, percentiles, section="Quality", is_quality=False, month_data=None, user_level=None):
    styles = get_common_styles()
    weeks = weeks or []

    # Initialize sets
    user_pipelines = set()
    all_pipelines = set()

    # Add weekly pipelines where user has worked
    if data:
        for pipe, weeks_data in data.items():
            if pipe != "Overall":
                has_valid_data = False
                for w in weeks:
                    week_data = weeks_data.get(w, {})
                    if isinstance(week_data, dict):
                        vol = week_data.get('volume', 0)
                        time = week_data.get('time', 0)
                        if vol > 0 and time > 0:  # Check for actual work done
                            has_valid_data = True
                            break
                if has_valid_data:
                    all_pipelines.add(pipe)

    # Add monthly pipelines where user has worked
    if month_data:
        for pipe, val in month_data.items():
            if pipe != "Overall":
                if isinstance(val, (float, int)) and val > 0:
                    all_pipelines.add(pipe)
                elif isinstance(val, dict):
                    total = val.get('total', 0)
                    if total > 0:
                        all_pipelines.add(pipe)

    # Sort pipelines
    all_pipelines = sorted(all_pipelines)

    # Monthly info
    month_num = self.month_selection.get()
    month_name = calendar.month_name[int(month_num)]

    # Calculate column widths
    pipeline_width = 25
    data_column_width = 60 / len(weeks) if weeks else 60
    month_width = 15

    # Create table headers
    ths = "".join(
        f"<th style='{styles['header_style']} width: {data_column_width}%;'>{w}</th>"
        for w in weeks
    )
    ths += f"<th style='{styles['header_style']} width: {month_width}%;'>Month({month_name})</th>"

    # Initialize weekly totals
    weekly_totals = {w: {'volume': 0, 'time': 0} for w in weeks}

    # Build rows
    rows = ""
    for pipe in all_pipelines:
        tds = ""
        for w in weeks:
            if is_quality:
                val_dict = data.get(pipe, {}).get(w, {})
                val = val_dict.get('score', 0)
                err = val_dict.get('err', 0)
                total = val_dict.get('total', 0)
                disp = f"{val:.1%}<br>({err}/{total})" if total else "-"
            else:
                entry = data.get(pipe, {}).get(w, {})
                if isinstance(entry, dict):
                    vol = entry.get('volume', 0)
                    time = entry.get('time', 0)
                    weekly_totals[w]['volume'] += vol
                    weekly_totals[w]['time'] += time
                    prod = vol / time if time else 0
                    disp = f"{prod:.1f}" if prod else "-"
                else:
                    disp = "-"
            tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{disp}</td>"

        # Monthly data
        if month_data and pipe in month_data:
            month_val = month_data[pipe]
            if is_quality and isinstance(month_val, dict):
                score = month_val.get('score', 0)
                err = month_val.get('errors', 0)
                total = month_val.get('total', 0)
                month_disp = f"{score:.1%}<br>({err}/{total})" if total else "-"
            elif not is_quality and isinstance(month_val, (float, int)):
                month_disp = f"{month_val:.1f}" if month_val else "-"
            else:
                month_disp = "-"
        else:
            month_disp = "-"
        tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{month_disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'>{pipe}</td>{tds}</tr>"

    # Overall Row
    overall_tds = ""
    for w in weeks:
        if is_quality:
            total_volume = sum(data.get(pipe, {}).get(w, {}).get('total', 0) for pipe in all_pipelines)
            total_errors = sum(data.get(pipe, {}).get(w, {}).get('err', 0) for pipe in all_pipelines)
            overall_score = (total_volume - total_errors) / total_volume if total_volume else 0
            overall_disp = f"{overall_score:.1%}<br>({total_errors}/{total_volume})" if total_volume else "-"
        else:
            total_vol = weekly_totals[w]['volume']
            total_time = weekly_totals[w]['time']
            prod = total_vol / total_time if total_time > 0 else 0
            overall_disp = f"{prod:.1f}" if prod else "-"
        overall_tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{overall_disp}</td>"

    # Overall month
    if month_data and 'Overall' in month_data:
        monthly_overall = month_data['Overall']
        if is_quality and isinstance(monthly_overall, dict):
            score = monthly_overall.get('score', 0)
            err = monthly_overall.get('errors', 0)
            total = monthly_overall.get('total', 0)
            monthly_disp = f"{score:.1%}<br>({err}/{total})" if total else "-"
        elif not is_quality and isinstance(monthly_overall, (float, int)):
            monthly_disp = f"{monthly_overall:.1f}" if monthly_overall else "-"
        else:
            monthly_disp = "-"
    else:
        monthly_disp = "-"
    overall_tds += f"<td style='{styles['cell_style']} width: {month_width}%;'>{monthly_disp}</td>"

    # Add Overall row
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
                # Handle list of ACL data
                for acl_data in precision_data[user][week]:
                    acl = acl_data['acl']
                    value = acl_data['value']
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

    # Initialize pipelines
    all_pipelines = sorted({
        pipe for pipe in data
        if pipe != "Overall" and any(data[pipe].get(w) for w in weeks)
    })

    # Column widths
    num_weeks = len(weeks)
    pipeline_width = 30
    week_width = 70 / num_weeks if num_weeks > 0 else 70

    # Headers
    ths = "".join(
        f"<th style='{styles['header_style']} width: {week_width}%;'>{w}</th>"
        for w in weeks
    )

    # Table rows
    rows = ""
    for pipe in all_pipelines:
        tds = ""
        for w in weeks:
            # Default
            p50_value = "-"
            if percentiles and w in percentiles:
                pctls = percentiles[w].get(pipe, {})
                if pctls and user_level in pctls:
                    p50 = pctls[user_level].get("p50")
                    if p50 is not None:
                        # Format based on section type
                        if section.lower() == "quality":
                            p50_value = f"{p50:.1%}"  # percent format
                        else:
                            p50_value = f"{p50:.2f}"  # numeric

            tds += f"<td style='{styles['cell_style']} width: {week_width}%;'>{p50_value}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {pipeline_width}%;'>{pipe}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 1}" style="{styles['title_style']}">
                    Weekly {section} P50 Benchmark Values
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

        # Filter by user
        if user:
            df_monthly = df_monthly[df_monthly['useralias'] == user]

        # Filter by month
        df_monthly = df_monthly[df_monthly['p_month'] == month_num]

        # Calculate total volume and time
        df_monthly['total_volume'] = df_monthly['processed_volume'] + df_monthly['processed_volumetr']
        df_monthly['total_time'] = df_monthly['processed_time'] + df_monthly['processed_time_tr']

        # Calculate metrics only for pipelines with actual work done
        monthly_metrics = {}
        pipeline_groups = df_monthly.groupby('pipeline')

        for pipeline, group in pipeline_groups:
            total_volume = group['total_volume'].sum()
            total_time = group['total_time'].sum()

            # Only include pipelines where the user actually did work
            if total_volume > 0 and total_time > 0:
                productivity = total_volume / total_time
                monthly_metrics[pipeline] = productivity

        # Calculate overall only from pipelines where work was done
        if monthly_metrics:  # Only if there are valid pipelines
            total_volume = df_monthly[df_monthly['total_volume'] > 0]['total_volume'].sum()
            total_time = df_monthly[df_monthly['total_time'] > 0]['total_time'].sum()
            if total_time > 0:
                monthly_metrics['Overall'] = total_volume / total_time

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

        # Filter by month
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
            count = len(grp)  # Add this line to get count

            error_mask = (
                    grp['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
                    ~grp['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
            )
            errors = grp[error_mask]['volume'].sum()

            monthly_metrics[correction_type] = {
                'score': (total_volume - errors) / total_volume if total_volume > 0 else 0,
                'errors': errors,
                'total': total_volume,
                'count': count  # Add this line to include count in output
            }

        # Calculate overall metrics
        total_volume = df['volume'].sum()
        total_count = len(df)  # Add this line for overall count
        error_mask = (
            df['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT']) &
            ~df['auditor_reappeal_final_judgement'].isin(['Both Correct', 'Auditor Correct'])
        )
        total_errors = df[error_mask]['volume'].sum()

        monthly_metrics['Overall'] = {
            'score': (total_volume - total_errors) / total_volume if total_volume > 0 else 0,
            'errors': total_errors,
            'total': total_volume,
            'count': total_count  # Add this line to include overall count
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


def html_qc2_reason_value_table(self, subreason_data, weeks):
    styles = get_common_styles()
    if not weeks:
        weeks = []
    if not subreason_data:
        subreason_data = {}

    # Calculate column widths without month column
    subreason_width = 30
    data_column_width = 70 / len(weeks)  # Distribute remaining width among week columns only

    # Create header cells without month column
    ths = "".join(
        f"<th style='{styles['header_style']} width: {data_column_width}%;'>{w}</th>"
        for w in weeks
    )

    # Create rows without month column
    rows = ""
    all_subs = sorted({sub for pipe in subreason_data for w in subreason_data[pipe]
                       for sub in subreason_data[pipe][w]['counts']})

    for sub in all_subs:
        tds = ""
        for w in weeks:
            count = sum(subreason_data[pipe][w]['counts'].get(sub, 0)
                        for pipe in subreason_data if w in subreason_data[pipe])
            tds += f"<td style='{styles['cell_style']} width: {data_column_width}%;'>{count if count else '-'}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {subreason_width}%;'>{sub}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 1}" style="{styles['title_style']}">
                    Error Reason Count
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

    # Get all unique subreasons from counts
    all_subs = sorted({sub for pipe in subreason_data for w in subreason_data[pipe]
                       for sub in subreason_data[pipe][w].get('counts', {})})

    # Column widths
    num_weeks = len(weeks)
    subreason_width = 30
    week_width = 70 / num_weeks

    # Table header
    ths = "".join(
        f"<th style='{styles['header_style']} width: {week_width}%;'>{w}</th>"
        for w in weeks
    )

    # Table rows
    rows = ""
    for sub in all_subs:
        tds = ""
        for w in weeks:
            numerator = 0
            denominator = 0
            for pipe in subreason_data:
                if w in subreason_data[pipe]:
                    data = subreason_data[pipe][w]
                    numerator += data.get('counts', {}).get(sub, 0)
                    denominator += data.get('total', 0)

            perc = (numerator / denominator * 100) if denominator else 0
            disp = f"{perc:.1f}%" if denominator else "-"
            tds += f"<td style='{styles['cell_style']} width: {week_width}%;'>{disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']} width: {subreason_width}%;'>{sub}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 1}" style="{styles['title_style']}">
                    Error Reason %
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
    # Sort correction types to ensure consistent order
    correction_types = sorted(correction_data.keys())

    for correction_type in correction_types:
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
            month_disp = f"{month_val['score']:.1%}<br>({month_val['errors']}/{month_val['total']})<br>Count: {month_val['count']}"
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


def html_correction_type_count_table(correction_data, weeks, month_data=None, month_num=None):
    styles = get_common_styles()

    # Get month name
    if month_num:
        month_name = calendar.month_name[int(month_num)]
    else:
        month_name = calendar.month_name[datetime.now().month]

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
            val = correction_data.get(correction_type, {}).get(w, {})
            total = val.get('total', 0)
            disp = str(total) if total else "-"
            tds += f"<td style='{styles['cell_style']}'>{disp}</td>"

        # Add monthly data
        if month_data and correction_type in month_data:
            month_val = month_data[correction_type]
            month_total = month_val.get('total', None)
            month_disp = str(month_total) if month_total is not None else "-"
        else:
            month_disp = "-"
        tds += f"<td style='{styles['cell_style']}'>{month_disp}</td>"

        rows += f"<tr><td style='{styles['pipeline_cell_style']}'>{correction_type}</td>{tds}</tr>"

    return f"""
    <div style="{styles['container_style']}">
        <table style="{styles['table_style']}">
            <tr>
                <td colspan="{len(weeks) + 2}" style="{styles['title_style']}">
                    Correction Type Total Case Count
                </td>
            </tr>
            <tr>
                <th style="{styles['header_style']}">Type</th>
                {ths}
            </tr>
            {rows}
        </table>
    </div>
    """

def compose_html(user, prod_table, prod_pct_table, qual_table, qual_pct_table, qc2_left, qc2_right,
                 correction_quality_table, correction_count_table, timesheet_table=""):
    table_style = """
    <style>
        /* Remove fixed layout from outer tables */
        table.outer-table {
            border: 3px solid white;
            border-collapse: separate !important;
            border-spacing: 0;
            background-color: #1f2937;
            width: 100%;
        }
        /* Correction tables: fixed layout for even columns */
        .correction-table {
            table-layout: fixed;
            width: 100%;
        }
        .correction-table th, .correction-table td {
            text-align: center;
            /* Remove or adjust min-width as needed */
            min-width: 60px;
        }
        /* General table styles */
        td, th {
            /* Remove global min-width */
            padding: 8px;
        }
        * {
            box-sizing: border-box;
        }
        @media (max-width: 768px) {
            .table-container {
                display: block !important;
                width: 100% !important;
            }
            .table-container td {
                display: block !important;
                width: 100% !important;
                padding: 0 !important;
                margin-bottom: 20px !important;
            }
        }
    </style>
    """
    return f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        {table_style}
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
                    <h2 style="font-size:18px;margin:0;font-weight:bold;color:white;">📈 Productivity Metrics</h2>
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
                <td style="padding:12px;text-align:center;font-size:13px;color:#d1d5db;">
                    Note: Week 26 is pre-appeal metrics. July Month is not the final Quality.
                </td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table" style="margin-bottom:30px;">
            <tr>
                <td style="padding:12px;text-align:center;">
                    <h2 style="font-size:18px;margin:0;font-weight:bold;color:white;">🎯 QC2 Subreason Analysis</h2>
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
                    <p style="font-size:14px;color:#d1d5db;margin:10px 0 0 0;line-height:1.5;">
                        <strong style="color:#ffffff;font-size:16px;">📊 QC2 Reason Analysis:</strong><br>
                        <span style="font-size:13px;">If only the table header is shown with no data, it means there were no errors found in the selected weeks.<br>
                        Percentages are calculated as (Error Count / Total Errors) × 100</span>
                    </p>
                    <p style="font-size:14px;color:#d1d5db;margin:10px 0 0 0;line-height:1.5;">
                        <strong style="color:#ffffff;font-size:16px;">⚡ Support:</strong><br>
                        <span style="font-size:13px;">Please reach out to Rohan/Nadeem/Pranav/Sridhar for any queries regarding format and new changes needed.</span>
                    </p>
                </td>
            </tr>
        </table>

        <table cellpadding="0" cellspacing="0" class="outer-table" style="margin-top:20px;">
            <tr>
                <td style="padding:15px;text-align:center;">
                    <p style="font-size:14px;color:#d1d5db;margin:0;font-style:italic;">
                        "Excellence is not a destination; it's a continuous journey that never ends."
                    </p>
                    <p style="font-size:12px;color:#d1d5db;margin:15px 0 0 0;">
                        Regards,<br>
                        RDR Operations
                    </p>
                </td>
            </tr>
        </table>

    </body>
    </html>
    """

def send_mail_html(to, cc, subject, html, preview=True):
    outlook = win32.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace("MAPI")

    # Find the shared mailbox
    shared_mailbox = None
    for account in outlook.Session.Accounts:
        if "rdr-team-operational-metrics" in account.DisplayName.lower():
            shared_mailbox = account
            break

    if shared_mailbox:
        mail = outlook.CreateItem(0)
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, shared_mailbox))  # Force the From field
        mail.To = to
        mail.CC = cc
        mail.Subject = subject
        mail.HTMLBody = html

        if preview:
            mail.Display()
        else:
            mail.Send()
    else:
        print("Shared mailbox not found")

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
        card, content = self._create_card(parent, "   File Selection")

        # Create file input frames - only for main Excel file
        self.data_frame, _ = self._create_input_field(
            content, "   Excel File (with Volume, Quality, TimeSheet, and Manager sheets)",
            self.data,
            command=lambda: self._browse(self.data)
        )

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

        # Get month name from selection
        month_num = self.month_selection.get()
        month_name = calendar.month_name[month_num]

        # Get all unique pipelines from both weekly and monthly data
        all_pipelines = set()

        # Add pipelines from weekly data
        if data:
            for pipe in data:
                if pipe != "Overall":
                    all_pipelines.add(pipe)

        # Add pipelines from monthly data
        if month_data:
            for pipe, val in month_data.items():
                if pipe != "Overall":
                    if isinstance(val, dict):
                        if any(v > 0 for v in val.values()):
                            all_pipelines.add(pipe)
                    elif isinstance(val, (float, int)) and val > 0:
                        all_pipelines.add(pipe)
        # Sort pipelines
        all_pipelines = sorted(all_pipelines)

        # Initialize weekly totals
        weekly_totals = {w: {'volume': 0, 'time': 0} for w in weeks}

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
        for pipe in all_pipelines:
            tds = ""
            for w in weeks:
                if is_quality:
                    val_dict = data.get(pipe, {}).get(w, {})
                    val = val_dict.get('score', 0)
                    err = val_dict.get('err', 0)
                    total = val_dict.get('total', 0)
                    disp = f"{val:.1%}<br>({err}/{total})" if total else "-"
                else:
                    entry = data.get(pipe, {}).get(w, {})
                    if isinstance(entry, dict):
                        vol = entry.get('volume', 0)
                        time = entry.get('time', 0)
                        weekly_totals[w]['volume'] += vol
                        weekly_totals[w]['time'] += time
                        prod = vol / time if time else 0
                        disp = f"{prod:.1f}" if prod else "-"
                    else:
                        disp = "-"

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

        # Calculate Overall row
        overall_tds = ""
        for w in weeks:
            if is_quality:
                total_volume = sum(data.get(pipe, {}).get(w, {}).get('total', 0) for pipe in all_pipelines)
                total_errors = sum(data.get(pipe, {}).get(w, {}).get('err', 0) for pipe in all_pipelines)
                overall_score = (total_volume - total_errors) / total_volume if total_volume else 0
                overall_disp = f"{overall_score:.1%}<br>({total_errors}/{total_volume})" if total_volume else "-"
            else:
                total_vol = weekly_totals[w]['volume']
                total_time = weekly_totals[w]['time']
                overall_prod = total_vol / total_time if total_time > 0 else 0
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

    def html_qc2_reason_value_table(self, subreason_data, weeks):
        styles = get_common_styles()
        if not weeks:
            weeks = []
        if not subreason_data:
            subreason_data = {}

        # Calculate column widths
        subreason_width = 30  # Reduced width for subreason column
        data_column_width = 70 / len(weeks)  # Evenly distribute remaining width among week columns

        # Define cell styles
        data_cell_style = f"{styles['cell_style']} width: {data_column_width}%;"
        subreason_cell_style = f"{styles['pipeline_cell_style']} width: {subreason_width}%;"

        # Create header cells
        ths = "".join(
            f"<th style='{styles['header_style']} width: {data_column_width}%;'>{w}</th>"
            for w in weeks
        )

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

            rows += f"<tr><td style='{subreason_cell_style}'>{sub}</td>{tds}</tr>"

        return f"""
        <div style="{styles['container_style']}">
            <table style="{styles['table_style']}">
                <tr>
                    <td colspan="{len(weeks) + 1}" style="{styles['title_style']}">
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
        styles = get_common_styles()

        # Get month name
        current_date = datetime.now()
        month_name = calendar.month_name[current_date.month]

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

            # Add monthly data - removed count display
            if month_data and correction_type in month_data:
                month_val = month_data[correction_type]
                score = month_val.get('score', 0)
                errors = month_val.get('errors', 0)
                total = month_val.get('total', 0)
                month_disp = f"{score:.1%}<br>({errors}/{total})"
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

        # Remove the mgr check and only check for data file
        if not self.data.get():
            return self.log("❌ Excel file required (must include Manager sheet)")

        # Call appropriate bulk sending function based on report type
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

            # Get levels for all auditors
            auditor_levels = df.groupby('useralias')['level'].first().to_dict()
            prod_pct = productivity_percentiles(allprod, weeks, auditor_levels)

            # Calculate monthly productivity metrics
            monthly_prod_metrics = calculate_monthly_productivity_metrics(df, self.month_selection.get(), user=user)

            # Check for productivity data
            has_weekly_prod = any(bool(week_data) for pipe_data in prod.values() for week_data in pipe_data.values())
            has_monthly_prod = bool(monthly_prod_metrics)

            if not has_weekly_prod and not has_monthly_prod:
                self.log(
                    f"⚠️ No productivity data found for {user} in specified weeks/month - skipping productivity section")
                prod_table = ""
                prod_pct_table = ""
            else:
                # Generate productivity tables only if data exists
                prod_table = self.html_metric_value_table_with_latest(
                    prod, weeks, prod_pct, section="Productivity",
                    is_quality=False, month_data=monthly_prod_metrics, user_level=user_level)

                prod_pct_table = html_metric_pct_table(
                    prod, weeks, prod_pct, section="Productivity", user_level=user_level)

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
                    monthly_subreason = calculate_monthly_subreason_metrics(qdf, self.month_selection.get(), user=user)
                    monthly_correction = calculate_monthly_correction_metrics(qdf, self.month_selection.get(),
                                                                              user=user)

                    # Generate quality tables
                    qual_table = self.html_metric_value_table_with_latest(
                        qual, weeks, qual_pct, section="Quality",
                        is_quality=True, month_data=monthly_qual_metrics,
                        user_level=user_level)

                    qual_pct_table = html_metric_pct_table(
                        qual, weeks, qual_pct, section="Quality",
                        user_level=user_level)

                    # QC2 analysis
                    subreason = qc2_subreason_analysis(qdf, user, weeks)
                    qc2_left = self.html_qc2_reason_value_table(
                        subreason, weeks)
                    qc2_right = html_qc2_reason_pct_table(subreason, weeks)

                    # Correction analysis
                    correction_data = calculate_correction_type_data(qdf, user, weeks)
                    correction_quality_table = self.html_correction_type_quality_table(
                        correction_data, weeks, month_data=monthly_correction)
                    correction_count_table = html_correction_type_count_table(
                        correction_data,
                        weeks,
                        month_data=monthly_correction,
                        month_num=self.month_selection.get()
                    )

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
        if not self.data.get():
            return self.log("❌ Excel file required (must include Manager sheet)")

        weeks = self._weeks()
        if not weeks:
            return

        try:
            # Load manager mapping from Manager sheet
            try:
                mgr = pd.read_excel(self.data.get(), sheet_name='Manager')
                if 'loginname' not in mgr.columns:
                    self.log("❌ Manager sheet must have 'loginname' column")
                    return
                if 'supervisorloginname' not in mgr.columns:
                    self.log("❌ Manager sheet must have 'supervisorloginname' column")
                    return

                valid_users = mgr['loginname'].dropna().unique()
                if len(valid_users) == 0:
                    self.log("❌ No valid users found in Manager sheet")
                    return

                self.log(f"📋 Found {len(valid_users)} users in Manager sheet")

            except Exception as e:
                self.log(f"❌ Error reading Manager sheet: {str(e)}")
                return

            # Load productivity data and get auditor levels
            self.log("\nProcessing productivity data...")
            df, auditor_levels = self.load_data()
            if df is None:
                return

            # Load quality data if available
            self.log("\nProcessing quality data...")
            qdf = None
            try:
                qdf = load_quality(self.data.get())
            except Exception as qe:
                self.log(f"⚠️ Error loading quality data: {str(qe)}")
            # Process each valid user
            processed_count = 0
            for user in valid_users:
                try:
                    # Get supervisor email from Manager sheet
                    supervisor_info = mgr[mgr['loginname'] == user]['supervisorloginname'].iloc[0]
                    cc_email = f"{supervisor_info}@amazon.com" if pd.notna(supervisor_info) else ""

                    # Get user's level
                    user_level = auditor_levels.get(user)
                    if user_level is None:
                        self.log(f"⚠️ Skipping user {user} - no level found")
                        continue

                    self.log(f"Processing {user} (Level {user_level})...")

                    # Calculate user-specific productivity metrics
                    user_df = df[df['useralias'] == user]
                    prod = user_prod_dict(df, user, weeks)

                    # Calculate all users' metrics for percentiles
                    allprod = all_prod_dict(df, weeks)
                    prod_pct = productivity_percentiles(allprod, weeks, auditor_levels)

                    # Calculate monthly metrics specific to this user
                    monthly_prod_metrics = calculate_monthly_productivity_metrics(df, self.month_selection.get(),
                                                                                  user=user)

                    # Check for productivity data
                    has_weekly_prod = any(
                        bool(week_data) for pipe_data in prod.values() for week_data in pipe_data.values())
                    has_monthly_prod = bool(monthly_prod_metrics)

                    if not has_weekly_prod and not has_monthly_prod:
                        self.log(f"⚠️ No productivity data found for {user} - skipping productivity section")
                        prod_table = ""
                        prod_pct_table = ""
                    else:
                        # Generate productivity tables
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
                        # Calculate quality metrics for this user
                        qual = user_quality_dict(qdf, user, weeks)
                        allqual = all_quality_dict(qdf, weeks)
                        qual_pct = quality_percentiles(allqual, weeks, auditor_levels)

                        # Calculate monthly quality metrics for this user
                        monthly_qual_metrics = calculate_monthly_quality_metrics(qdf, self.month_selection.get(),
                                                                                 user=user)
                        monthly_subreason = calculate_monthly_subreason_metrics(qdf, self.month_selection.get(),
                                                                                user=user)
                        monthly_correction = calculate_monthly_correction_metrics(qdf, self.month_selection.get(),
                                                                                  user=user)

                        qual_table = self.html_metric_value_table_with_latest(
                            qual, weeks, qual_pct, section="Quality",
                            is_quality=True, month_data=monthly_qual_metrics, user_level=user_level)

                        qual_pct_table = html_metric_pct_table(
                            qual, weeks, qual_pct, section="Quality", user_level=user_level)

                        # QC2 analysis
                        subreason = qc2_subreason_analysis(qdf, user, weeks)
                        qc2_left = self.html_qc2_reason_value_table(subreason, weeks)
                        qc2_right = html_qc2_reason_pct_table(subreason, weeks)

                        # Correction analysis
                        correction_data = calculate_correction_type_data(qdf, user, weeks)
                        correction_quality_table = self.html_correction_type_quality_table(
                            correction_data, weeks, month_data=monthly_correction)
                        correction_count_table = html_correction_type_count_table(
                            correction_data,
                            weeks,
                            month_data=monthly_correction,
                            month_num=self.month_selection.get()
                        )

                    # Process timesheet data
                    timesheet_table = ""
                    try:
                        timesheet_data = analyze_timesheet_missing(self.data.get(), weeks[-1], user)
                        if timesheet_data is not None and not timesheet_data.empty:
                            timesheet_table = html_timesheet_missing_table(timesheet_data)
                    except Exception as te:
                        self.log(f"⚠️ Error processing timesheet for {user}: {str(te)}")

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

                    # Send email
                    send_mail_html(
                        to=f"{user}@amazon.com",
                        cc=cc_email,
                        subject="RDR Productivity & Quality Metrics Report",
                        html=html,
                        preview=self.mode.get() == "preview"
                    )

                    mode_text = "previewed" if self.mode.get() == "preview" else "sent"
                    supervisor_text = f" (CC: {supervisor_info})" if cc_email else ""
                    self.log(f"✅ Report for {user}{supervisor_text} {mode_text}")
                    processed_count += 1

                except Exception as e:
                    self.log(f"⚠️ Error processing {user}: {str(e)}")
                    continue

            self.log(
                f"🎉 Bulk processing completed! Successfully processed {processed_count} out of {len(valid_users)} users.")

        except Exception as e:
            self.log(f"❌ Error in bulk processing: {str(e)}")

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
            df, auditor_levels = self.load_data()  # Unpack both values
            if df is None:
                return

            # Load manager mapping from Manager sheet
            try:
                mgr = pd.read_excel(self.data.get(), sheet_name='Manager')
                supervisor_info = mgr[mgr['loginname'] == user]['supervisorloginname'].iloc[0]
                cc_email = f"{supervisor_info}@amazon.com" if pd.notna(supervisor_info) else ""
            except Exception as e:
                self.log(f"⚠️ Warning: Could not get supervisor email: {str(e)}")
                cc_email = ""

            precision_data = calculate_precision_corrections(df, weeks)

            if user in precision_data:
                precision_table = html_precision_correction_table({user: precision_data[user]}, weeks)
                html = f"""
                <html>
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <style>
                        body {{
                            margin: 0;
                            padding: 20px;
                            font-family: Segoe UI, Arial, sans-serif;
                            color: #333;
                            line-height: 1.4;
                            background: transparent;  /* Changed from #f9f9fb to transparent */
                        }}
                        .container {{
                            margin: 0 auto;
                            background: transparent;  /* Changed from white to transparent */
                            padding: 30px;
                        }}
                        h2 {{
                            font-size: 18px;
                            margin-bottom: 10px;
                            font-weight: normal;
                            color: #1a365d;
                        }}
                        .warning {{
                            margin-top: 20px;
                            padding: 15px;
                            border-left: 4px solid #dc2626;
                            color: #dc2626;
                            font-style: italic;
                            font-size: 13px;
                            background: transparent;  /* Added transparent background */
                        }}
                        .signature {{
                            margin-top: 30px;
                            padding-top: 20px;
                            border-top: 1px solid #e5e7eb;
                            font-size: 12px;
                            color: #666;
                            background: transparent;  /* Added transparent background */
                        }}
                    </style>
                </head>
                <body>
                    <div class="container">
                        <h2>High Precision Correction Alert</h2>
                        <p style="margin-top:0;font-size:13px;">Hi {user}, you have made high precision corrections:</p>

                        {precision_table}

                        <div class="warning">
                            ⚠️ Important: Please ensure corrections are made only when absolutely necessary. Unnecessary corrections may lead to escalations and impact quality metrics. If unsure, please consult with your supervisor before making corrections.
                        </div>

                        <div class="signature">
                            <p style="margin:0;">Regards,<br>RDR Operations</p>
                        </div>
                    </div>
                </body>
                </html>
                """

                weeks_str = ", ".join(str(w) for w in weeks)
                send_mail_html(
                    to=f"{user}@amazon.com",
                    cc=cc_email,
                    subject=f"High Precision Corrections Alert - Week {weeks_str}",
                    html=html,
                    preview=self.mode.get() == "preview"
                )

                mode_text = "previewed" if self.mode.get() == "preview" else "sent"
                supervisor_text = f" (CC: {supervisor_info})" if cc_email else ""
                self.log(f"✅ Precision mail for {user}{supervisor_text} {mode_text}")
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
                df, auditor_levels = load_productivity(self.data.get())
                precision_data = calculate_precision_corrections(df, weeks)

                if not precision_data:
                    self.log("ℹ️ No high precision corrections found for any user")
                    return False

                # Load manager mapping from Manager sheet
                try:
                    mgr = pd.read_excel(self.data.get(), sheet_name='Manager')
                except Exception as e:
                    self.log(f"⚠️ Warning: Could not load Manager sheet: {str(e)}")
                    mgr = pd.DataFrame(columns=['loginname', 'supervisorloginname'])

                self.log(f"   Processing precision reports for {len(precision_data)} users...")

                # Process each user
                for user in precision_data:
                    try:
                        # Get supervisor email
                        supervisor_info = mgr[mgr['loginname'] == user]['supervisorloginname'].iloc[
                            0] if not mgr.empty else None
                        cc_email = f"{supervisor_info}@amazon.com" if pd.notna(supervisor_info) else ""

                        precision_table = html_precision_correction_table({user: precision_data[user]}, weeks)
                        html = f"""
                        <html>
                        <head>
                            <meta charset="UTF-8">
                            <meta name="viewport" content="width=device-width, initial-scale=1.0">
                            <style>
                                body {{
                                    margin: 0;
                                    padding: 20px;
                                    font-family: Segoe UI, Arial, sans-serif;
                                    color: #333;
                                    line-height: 1.4;
                                    background: transparent;  /* Changed from #f9f9fb to transparent */
                                }}
                                .container {{
                                    margin: 0 auto;
                                    background: transparent;  /* Changed from white to transparent */
                                    padding: 30px;
                                }}
                                h2 {{
                                    font-size: 18px;
                                    margin-bottom: 10px;
                                    font-weight: normal;
                                    color: #1a365d;
                                }}
                                .warning {{
                                    margin-top: 20px;
                                    padding: 15px;
                                    border-left: 4px solid #dc2626;
                                    color: #dc2626;
                                    font-style: italic;
                                    font-size: 13px;
                                    background: transparent;  /* Added transparent background */
                                }}
                                .signature {{
                                    margin-top: 30px;
                                    padding-top: 20px;
                                    border-top: 1px solid #e5e7eb;
                                    font-size: 12px;
                                    color: #666;
                                    background: transparent;  /* Added transparent background */
                                }}
                            </style>
                        </head>
                        <body>
                            <div class="container">
                                <h2>High Precision Correction Alert</h2>
                                <p style="margin-top:0;font-size:13px;">Hi {user}, you have made high precision corrections:</p>

                                {precision_table}

                                <div class="warning">
                                    ⚠️ Important: Please ensure corrections are made only when absolutely necessary. Unnecessary corrections may lead to escalations and impact quality metrics. If unsure, please consult with your supervisor before making corrections.
                                </div>

                                <div class="signature">
                                    <p style="margin:0;">Regards,<br>RDR Operations</p>
                                </div>
                            </div>
                        </body>
                        </html>
                        """

                        # Create weeks string for email subject
                        weeks_str = ", ".join(str(w) for w in weeks)

                        # Send email
                        send_mail_html(
                            f"{user}@amazon.com",
                            cc_email,
                            f"High Precision Corrections Alert - Week {weeks_str}",
                            html,
                            preview=self.mode.get() == "preview"
                        )

                        # Log success
                        mode_text = "previewed" if self.mode.get() == "preview" else "sent"
                        supervisor_text = f" (CC: {supervisor_info})" if cc_email else ""
                        self.log(f"✅ Precision mail for {user}{supervisor_text} {mode_text}")

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