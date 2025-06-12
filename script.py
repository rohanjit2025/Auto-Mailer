import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import win32com.client as win32
import pandas as pd
import numpy as np

import os
from collections import defaultdict


def calculate_productivity_data_optimized(data_file, user_login):
    """Optimized version - only reads data once and uses vectorized operations"""
    try:
        # Read only required columns to reduce memory usage
        required_cols = ['useralias', 'pipeline', 'p_week',
                         'processed_volume', 'processed_volume_tr',
                         'processed_time', 'processed_time_tr']

        # Read with specific dtypes to improve performance
        dtype_dict = {
            'useralias': 'string',
            'pipeline': 'string',
            'p_week': 'int32',
            'processed_volume': 'float32',
            'processed_volume_tr': 'float32',
            'processed_time': 'float32',
            'processed_time_tr': 'float32'
        }

        df = pd.read_excel(data_file, usecols=required_cols, dtype=dtype_dict)

        # Filter for the user immediately
        user_data = df[df['useralias'] == user_login].copy()
        if user_data.empty:
            return None, f"No data found for user: {user_login}"

        # Vectorized operations - fill NaN values
        user_data[['processed_volume', 'processed_volume_tr',
                   'processed_time', 'processed_time_tr']] = user_data[
            ['processed_volume', 'processed_volume_tr',
             'processed_time', 'processed_time_tr']].fillna(0)

        # Vectorized calculations
        user_data['total_volume'] = user_data['processed_volume'] + user_data['processed_volume_tr']
        user_data['total_time'] = user_data['processed_time'] + user_data['processed_time_tr']

        # Group and aggregate in one operation
        grouped = user_data.groupby(['pipeline', 'p_week'], as_index=False).agg({
            'total_volume': 'sum',
            'total_time': 'sum'
        })

        # Vectorized productivity calculation
        grouped['productivity'] = np.where(
            grouped['total_time'] > 0,
            grouped['total_volume'] / grouped['total_time'],
            0
        )

        # Convert to nested dictionary more efficiently
        grouped_data = defaultdict(dict)
        for row in grouped.itertuples():
            grouped_data[row.pipeline][row.p_week] = row.productivity

        return grouped_data, None

    except Exception as e:
        return None, f"Error processing data: {str(e)}"


def calculate_all_auditors_productivity_optimized(data_file, weeks_to_show=None):
    """Optimized version - processes all auditors in one pass"""
    try:
        # Read only required columns
        required_cols = ['useralias', 'pipeline', 'p_week',
                         'processed_volume', 'processed_volume_tr',
                         'processed_time', 'processed_time_tr']

        dtype_dict = {
            'useralias': 'string',
            'pipeline': 'string',
            'p_week': 'int32',
            'processed_volume': 'float32',
            'processed_volume_tr': 'float32',
            'processed_time': 'float32',
            'processed_time_tr': 'float32'
        }

        df = pd.read_excel(data_file, usecols=required_cols, dtype=dtype_dict)

        # Vectorized operations for all data at once
        df[['processed_volume', 'processed_volume_tr',
            'processed_time', 'processed_time_tr']] = df[
            ['processed_volume', 'processed_volume_tr',
             'processed_time', 'processed_time_tr']].fillna(0)

        df['total_volume'] = df['processed_volume'] + df['processed_volume_tr']
        df['total_time'] = df['processed_time'] + df['processed_time_tr']

        # Group by all three columns at once
        grouped = df.groupby(['useralias', 'pipeline', 'p_week'], as_index=False).agg({
            'total_volume': 'sum',
            'total_time': 'sum'
        })

        # Vectorized productivity calculation
        grouped['productivity'] = np.where(
            grouped['total_time'] > 0,
            grouped['total_volume'] / grouped['total_time'],
            0
        )

        # Convert to nested dictionary structure
        all_productivity_data = defaultdict(lambda: defaultdict(dict))
        for row in grouped.itertuples():
            all_productivity_data[row.useralias][row.pipeline][row.p_week] = row.productivity

        return all_productivity_data, None

    except Exception as e:
        return None, f"Error calculating all auditors productivity: {str(e)}"


def calculate_percentiles_for_all_weeks_optimized(all_productivity_data, weeks_to_show):
    """Calculate percentiles for ALL weeks, not just the latest one"""
    if not weeks_to_show:
        return {}, {}

    # Dictionary to store percentiles for each week
    weekly_pipeline_percentiles = {}
    weekly_overall_percentiles = {}

    for week in weeks_to_show:
        # Use lists for faster appending, then convert to numpy arrays
        pipeline_productivities = defaultdict(list)
        overall_productivities = []

        # Single pass through all data for this specific week
        for auditor, auditor_data in all_productivity_data.items():
            auditor_total_productivity = 0
            auditor_pipeline_count = 0

            for pipeline, weekly_data in auditor_data.items():
                if week in weekly_data:  # Changed condition to include zero values
                    productivity = weekly_data[week]
                    pipeline_productivities[pipeline].append(productivity)
                    auditor_total_productivity += productivity
                    auditor_pipeline_count += 1

            if auditor_pipeline_count > 0:
                avg_productivity = auditor_total_productivity / auditor_pipeline_count
                overall_productivities.append({
                    'auditor': auditor,
                    'avg_productivity': avg_productivity,
                    'total_productivity': auditor_total_productivity,
                    'pipeline_count': auditor_pipeline_count
                })

        # Vectorized percentile calculations for this week
        pipeline_percentiles = {}
        for pipeline, productivities in pipeline_productivities.items():
            # Remove the >= 2 condition to calculate percentiles for all data
            prod_array = np.array(productivities)
            pipeline_percentiles[pipeline] = {
                'p30': np.percentile(prod_array, 30) if len(productivities) > 0 else 0,
                'p50': np.percentile(prod_array, 50) if len(productivities) > 0 else 0,
                'p75': np.percentile(prod_array, 75) if len(productivities) > 0 else 0,
                'p90': np.percentile(prod_array, 90) if len(productivities) > 0 else 0,
                'count': len(productivities),
                'values': productivities
            }

        # Overall percentiles for this week
        overall_percentiles = {}
        if overall_productivities:
            avg_values = np.array([item['avg_productivity'] for item in overall_productivities])
            overall_percentiles = {
                'p30': np.percentile(avg_values, 30),
                'p50': np.percentile(avg_values, 50),
                'p75': np.percentile(avg_values, 75),
                'p90': np.percentile(avg_values, 90),
                'auditor_data': overall_productivities,
                'count': len(overall_productivities)
            }

        # Store results for this week
        weekly_pipeline_percentiles[week] = pipeline_percentiles
        weekly_overall_percentiles[week] = overall_percentiles

    return weekly_pipeline_percentiles, weekly_overall_percentiles


def calculate_all_auditors_quality_optimized(quality_file, weeks_to_show=None):
    """Calculate quality metrics for all auditors"""
    try:
        required_cols = ['week', 'program', 'auditor_login', 'usecase',
                         'qc2_judgement', 'auditor_appeal_judgement',
                         'auditor_reappeal_final_judgement']

        df = pd.read_excel(quality_file, usecols=required_cols)
        df = df[df['program'] == 'RDR']

        all_quality_data = defaultdict(lambda: defaultdict(dict))

        for week in weeks_to_show:
            week_data = df[df['week'] == week]

            for pipeline in week_data['usecase'].unique():
                pipeline_data = week_data[week_data['usecase'] == pipeline]

                for auditor in pipeline_data['auditor_login'].unique():
                    auditor_data = pipeline_data[pipeline_data['auditor_login'] == auditor]

                    if not auditor_data.empty:
                        total_cases = len(auditor_data)
                        incorrect_cases = len(auditor_data[
                                                  auditor_data['qc2_judgement'].isin(
                                                      ['AUDITOR_INCORRECT', 'BOTH_INCORRECT'])
                                              ])

                        quality = (total_cases - incorrect_cases) / total_cases if total_cases > 0 else 0

                        # Store both quality score and case counts
                        all_quality_data[auditor][pipeline][week] = {
                            'score': quality,
                            'error_count': incorrect_cases,
                            'total_count': total_cases
                        }

        return all_quality_data, None

    except Exception as e:
        return None, f"Error calculating quality data: {str(e)}"


def calculate_quality_percentiles_for_all_weeks(all_quality_data, weeks_to_show):
    """Calculate quality percentiles for all weeks"""
    if not weeks_to_show:
        return {}, {}

    weekly_pipeline_percentiles = {}
    weekly_overall_percentiles = {}

    for week in weeks_to_show:
        pipeline_qualities = defaultdict(list)
        overall_qualities = []

        for auditor, auditor_data in all_quality_data.items():
            auditor_total_quality = 0
            auditor_pipeline_count = 0

            for pipeline, weekly_data in auditor_data.items():
                if week in weekly_data:
                    # Extract the quality score from the dictionary
                    quality = weekly_data[week]['score'] if isinstance(weekly_data[week], dict) else weekly_data[week]
                    pipeline_qualities[pipeline].append(quality)
                    auditor_total_quality += quality
                    auditor_pipeline_count += 1

            if auditor_pipeline_count > 0:
                avg_quality = auditor_total_quality / auditor_pipeline_count
                overall_qualities.append({
                    'auditor': auditor,
                    'avg_quality': avg_quality,
                    'total_quality': auditor_total_quality,
                    'pipeline_count': auditor_pipeline_count
                })

        # Calculate percentiles for this week
        pipeline_percentiles = {}
        for pipeline, qualities in pipeline_qualities.items():
            quality_array = np.array(qualities)
            pipeline_percentiles[pipeline] = {
                'p30': np.percentile(quality_array, 30) if len(qualities) > 0 else 0,
                'p50': np.percentile(quality_array, 50) if len(qualities) > 0 else 0,
                'p75': np.percentile(quality_array, 75) if len(qualities) > 0 else 0,
                'p90': np.percentile(quality_array, 90) if len(qualities) > 0 else 0,
                'count': len(qualities),
                'values': qualities
            }

        overall_percentiles = {}
        if overall_qualities:
            avg_values = np.array([item['avg_quality'] for item in overall_qualities])
            overall_percentiles = {
                'p30': np.percentile(avg_values, 30),
                'p50': np.percentile(avg_values, 50),
                'p75': np.percentile(avg_values, 75),
                'p90': np.percentile(avg_values, 90),
                'auditor_data': overall_qualities,
                'count': len(overall_qualities)
            }

        weekly_pipeline_percentiles[week] = pipeline_percentiles
        weekly_overall_percentiles[week] = overall_percentiles

    return weekly_pipeline_percentiles, weekly_overall_percentiles

def get_percentile_background_color(category):
    colors = {
        "P90 +": "#28a745",  # Green
        "P75-P90": "#17a2b8",  # Teal
        "P50-P75": "#ffc107",  # Yellow
        "P30-P50": "#fd7e14",  # Orange
        "< P30": "#dc3545",    # Red
        "": "#ffffff"          # White for empty
    }
    return colors.get(category, "#ffffff")

def get_percentile_text_color(category):
    # Use white text for darker backgrounds, black for lighter ones
    white_text_categories = {"P90 +", "P75-P90", "P30-P50", "< P30"}
    return "white" if category in white_text_categories else "black"


def generate_percentile_table(grouped_data, weekly_pipeline_percentiles, weeks_to_show):
    """Generate a compact, low-height Outlook-friendly percentile table"""
    html = f'''
    <table style="border-collapse: collapse; width: 300px; font-size: 11px; font-family: Arial, sans-serif; color: #333; background: #fff; border: 1px solid #ccc; line-height: 0.8;">
        <tr>
            <th colspan="{len(weeks_to_show) + 1}" style="background-color: #34495e; color: white; padding: 1px 2px; height: 15px; text-align: center; font-weight: bold;">
                Weekly Benchmarks
            </th>
        </tr>
        <tr>
            <th style="background-color: #f2f2f2; padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc;">Pipeline</th>'''

    for week in weeks_to_show:
        html += f'''
            <th style="background-color: #f2f2f2; padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc;">{week}</th>'''

    html += "</tr>"

    pipelines = sorted(grouped_data.keys())
    for pipeline in pipelines:
        html += f'<tr>'
        html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">{pipeline}</td>'

        for week in weeks_to_show:
            if week in weekly_pipeline_percentiles and pipeline in weekly_pipeline_percentiles[week]:
                val = grouped_data[pipeline].get(week, 0)
                percentile_category = (
                    get_percentile_category_for_week(val, weekly_pipeline_percentiles, pipeline, week)
                    if val > 0 else ""
                )

                # Color logic
                bg_color = "#ffffff"
                text_color = "#333"
                if percentile_category == "P90 +":
                    bg_color = "#2ecc71"
                    text_color = "white"
                elif percentile_category == "P75-P90":
                    bg_color = "#3498db"
                    text_color = "white"
                elif percentile_category == "P50-P75":
                    bg_color = "#f1c40f"
                    text_color = "#333"
                elif percentile_category == "P30-P50":
                    bg_color = "#e67e22"
                    text_color = "white"
                elif percentile_category == "< P30":
                    bg_color = "#e74c3c"
                    text_color = "white"

                html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; background-color: {bg_color}; color: {text_color}; border: 1px solid #ccc;">{percentile_category}</td>'
            else:
                html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td>'
        html += '</tr>'

    # Overall row
    html += '<tr>'
    html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; font-weight: bold; border: 1px solid #ccc;">Overall</td>'
    for _ in weeks_to_show:
        html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td>'
    html += '</tr></table>'

    return html


def generate_quality_percentile_table(quality_data, weekly_quality_percentiles, weeks_to_show):
    """Generate quality percentile table - same format as productivity percentile table"""
    html = f'''
    <table style="border-collapse: collapse; width: 300px; font-size: 11px; font-family: Arial, sans-serif; color: #333; background: #fff; border: 1px solid #ccc; line-height: 0.8;">
        <tr>
            <th colspan="{len(weeks_to_show) + 1}" style="background-color: #ff7f00; color: white; padding: 1px 2px; height: 15px; text-align: center; font-weight: bold;">
                Weekly Quality Benchmarks
            </th>
        </tr>
        <tr>
            <th style="background-color: #000000; color: white; padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc;">Pipeline</th>'''

    for week in weeks_to_show:
        html += f'''
            <th style="background-color: #000000; color: white; padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc;">{week}</th>'''

    html += "</tr>"

    pipelines = sorted(quality_data.keys())
    for pipeline in pipelines:
        html += f'<tr>'
        html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">{pipeline}</td>'

        for week in weeks_to_show:
            if week in quality_data[pipeline]:
                val = quality_data[pipeline][week]
                score = val['score'] if isinstance(val, dict) else val

                if score > 0:
                    percentile_category = get_percentile_category(
                        score, weekly_quality_percentiles, pipeline, week)

                    bg_color = "#ffffff"
                    text_color = "#333"
                    if percentile_category == "P90 +":
                        bg_color = "#28a745"
                        text_color = "white"
                    elif percentile_category == "P75-P90":
                        bg_color = "#17a2b8"
                        text_color = "white"
                    elif percentile_category == "P50-P75":
                        bg_color = "#ffc107"
                        text_color = "#333"
                    elif percentile_category == "P30-P50":
                        bg_color = "#fd7e14"
                        text_color = "white"
                    elif percentile_category == "< P30":
                        bg_color = "#dc3545"
                        text_color = "white"

                    html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; background-color: {bg_color}; color: {text_color}; border: 1px solid #ccc;">{percentile_category}</td>'
                else:
                    html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td>'
            else:
                html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td>'
        html += '</tr>'

    # Overall row
    html += '<tr>'
    html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; font-weight: bold; border: 1px solid #ccc;">Overall</td>'
    for _ in weeks_to_show:
        html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td>'
    html += '</tr></table>'

    return html

def get_percentile_category(value, weekly_quality_percentiles, pipeline, week):
    """Helper function to determine percentile category"""
    if (week not in weekly_quality_percentiles or
            pipeline not in weekly_quality_percentiles[week]):
        return "-"

    percentiles = weekly_quality_percentiles[week][pipeline]

    if value >= percentiles['p90']:
        return "P90 +"
    elif value >= percentiles['p75']:
        return "P75-P90"
    elif value >= percentiles['p50']:
        return "P50-P75"
    elif value >= percentiles['p30']:
        return "P30-P50"
    else:
        return "< P30"

def get_percentile_category_for_week(user_productivity, weekly_pipeline_percentiles, pipeline, week):
    """Determine which percentile category the user falls into for a specific week"""
    if (week not in weekly_pipeline_percentiles or
            pipeline not in weekly_pipeline_percentiles[week] or
            user_productivity <= 0 or
            user_productivity is None or
            pd.isna(user_productivity)):
        return ""

    percentiles = weekly_pipeline_percentiles[week][pipeline]

    if user_productivity >= percentiles['p90']:
        return "P90 +"
    elif user_productivity >= percentiles['p75']:
        return "P75-P90"
    elif user_productivity >= percentiles['p50']:
        return "P50-P75"
    elif user_productivity >= percentiles['p30']:
        return "P30-P50"
    elif user_productivity >= 0:
        return "< P30"
    else:
        return "-"


def generate_productivity_table(grouped_data, weekly_pipeline_percentiles, weeks_to_show):
    html = f'''
        <table style="border-collapse: collapse; width: 320px; font-size: 11px; font-family: Arial, sans-serif; color: #333; background: #fff; border: 1px solid #ccc; line-height: 0.8;">
            <tr>
                <th style="padding: 1px 2px; height: 15px; background-color: #34495e; color: white; text-align: center; border: 1px solid #ccc;">Pipeline</th>'''

    for week in weeks_to_show:
        html += f'<th style="padding: 1px 2px; height: 15px; background-color: #34495e; color: white; text-align: center; border: 1px solid #ccc;">{week}</th>'

    html += '<th style="padding: 1px 2px; height: 15px; background-color: #34495e; color: white; text-align: center; border: 1px solid #ccc;">Latest Week Percentile</th></tr>'

    pipelines = sorted(grouped_data.keys())
    weekly_sums = {week: [] for week in weeks_to_show}
    latest_week = max(weeks_to_show) if weeks_to_show else None

    for pipeline in pipelines:
        html += '<tr>'
        html += f'<td style="padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc;">{pipeline}</td>'

        for week in weeks_to_show:
            val = grouped_data[pipeline].get(week, 0)
            if val > 0:
                weekly_sums[week].append(val)
                html += f'<td style="padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc; background-color: #f5f5f5;">{val:.2f}</td>'
            else:
                html += '<td style="padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc;">-</td>'

        latest_val = grouped_data[pipeline].get(latest_week, 0)
        percentile_value, bg_color, text_color = '-', '#ffffff', '#333'

        if (latest_week in weekly_pipeline_percentiles and pipeline in weekly_pipeline_percentiles[latest_week]):
            percentiles = weekly_pipeline_percentiles[latest_week][pipeline]
            if latest_val > 0:
                if latest_val >= percentiles['p90']:
                    percentile_value, bg_color, text_color = 'P90 +', '#28a745', 'white'
                elif latest_val >= percentiles['p75']:
                    percentile_value, bg_color, text_color = 'P75-P90', '#3498db', 'white'
                elif latest_val >= percentiles['p50']:
                    percentile_value, bg_color, text_color = 'P50-P75', '#f1c40f', '#333'
                elif latest_val >= percentiles['p30']:
                    percentile_value, bg_color, text_color = 'P30-P50', '#fd7e14', 'white'
                else:
                    percentile_value, bg_color, text_color = '< P30', '#dc3545', 'white'

        html += f'<td style="padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc; background-color: {bg_color}; color: {text_color};">{percentile_value}</td>'
        html += '</tr>'

    html += '<tr><td style="padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc; font-weight: bold;">Overall</td>'
    for week in weeks_to_show:
        if weekly_sums[week]:
            avg = sum(weekly_sums[week]) / len(weekly_sums[week])
            html += f'<td style="padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc; background-color: #f5f5f5; font-weight: bold;">{avg:.2f}</td>'
        else:
            html += '<td style="padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc;">-</td>'
    html += '<td style="padding: 1px 2px; height: 15px; text-align: center; border: 1px solid #ccc;">-</td></tr>'
    html += '</table>'
    return html


def generate_quality_table(quality_data, weekly_quality_percentiles, weeks_to_show):
    """Generate quality table showing quality values with error counts in brackets"""
    html = f'''
    <table style="border-collapse: collapse; width: 320px; font-size: 11px; font-family: Arial, sans-serif; color: #333; background: #fff; border: 1px solid #ccc; line-height: 0.8;">
        <tr>
            <th colspan="{len(weeks_to_show) + 2}" style="background-color: #34495e; color: white; padding: 1px 2px; height: 15px; text-align: center; font-weight: bold; border: 1px solid #ccc;">
                Weekly Quality Benchmarks
            </th>
        </tr>
        <tr>
            <th style="padding: 1px 2px; height: 15px; background-color: #34495e; color: white; text-align: center; border: 1px solid #ccc; width: 100px;">Pipeline/Week</th>'''

    # Fixed width for week columns
    for week in weeks_to_show:
        html += f'<th style="padding: 1px 2px; height: 15px; background-color: #34495e; color: white; text-align: center; border: 1px solid #ccc; width: 60px;">{week}</th>'

    # Latest week with (B) notation
    latest_week = max(weeks_to_show) if weeks_to_show else None
    html += f'<th style="padding: 1px 2px; height: 15px; background-color: #34495e; color: white; text-align: center; border: 1px solid #ccc; width: 100px;">{latest_week} (B)</th></tr>'

    # Fixed width for percentile column
    html += '<th style="padding: 1px 2px; height: 15px; background-color: #34495e; color: white; text-align: center; border: 1px solid #ccc; width: 100px;">Latest Week Percentile</th></tr>'

    pipelines = sorted(quality_data.keys())

    for pipeline in pipelines:
        html += '<tr>'
        html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">{pipeline}</td>'

        for week in weeks_to_show:
            if week in quality_data[pipeline]:
                data = quality_data[pipeline][week]
                if isinstance(data, dict):
                    score = data['score']
                    error_count = data['error_count']
                    total_count = data['total_count']
                    html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">{score:.2%} ({error_count}/{total_count})</td>'
                else:
                    html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">{data:.2%}</td>'
            else:
                html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td>'
        html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td></tr>'

    # Overall row
    html += '<tr>'
    html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc; font-weight: bold;">Overall</td>'

    for week in weeks_to_show:
        total_score = 0
        total_pipelines = 0
        total_errors = 0
        total_cases = 0

        for pipeline in pipelines:
            if week in quality_data[pipeline]:
                data = quality_data[pipeline][week]
                if isinstance(data, dict):
                    total_score += data['score']
                    total_errors += data['error_count']
                    total_cases += data['total_count']
                    total_pipelines += 1
                else:
                    total_score += data
                    total_pipelines += 1

        if total_pipelines > 0:
            avg_score = total_score / total_pipelines
            if total_cases > 0:
                html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc; font-weight: bold;">{avg_score:.2%} ({total_errors}/{total_cases})</td>'
            else:
                html += f'<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc; font-weight: bold;">{avg_score:.2%}</td>'
        else:
            html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td>'

    html += '<td style="padding: 1px 2px; height: 15px; font-size: 11px; line-height: 0.8; text-align: center; border: 1px solid #ccc;">-</td></tr>'
    html += "</table>"
    return html


def get_percentile_styling(value, weekly_percentiles, pipeline, week):
    """Helper function to get percentile styling - extracted for reuse"""
    if (week not in weekly_percentiles or
            pipeline not in weekly_percentiles[week] or
            value <= 0 or value is None or pd.isna(value)):
        return "", "#ffffff", "black"

    percentiles = weekly_percentiles[week][pipeline]

    if value >= percentiles['p90']:
        return "P90 +", "#28a745", "white"
    elif value >= percentiles['p75']:
        return "P75-P90", "#17a2b8", "white"
    elif value >= percentiles['p50']:
        return "P50-P75", "#ffc107", "black"
    elif value >= percentiles['p30']:
        return "P30-P50", "#fd7e14", "white"
    else:
        return "< P30", "#dc3545", "white"


def calculate_qc2_subreason_analysis(data_file, user_login, weeks_to_show=None):
    """Calculate QC2 subreason counts and percentages for incorrect cases"""
    try:
        required_cols = ['week', 'program', 'auditor_login', 'usecase',
                         'qc2_judgement', 'qc2_subreason']

        df = pd.read_excel(data_file, usecols=required_cols)
        df = df[df['program'] == 'RDR']

        subreason_data = defaultdict(lambda: defaultdict(dict))

        for week in weeks_to_show:
            week_data = df[df['week'] == week]

            for pipeline in week_data['usecase'].unique():
                pipeline_data = week_data[week_data['usecase'] == pipeline]
                auditor_data = pipeline_data[pipeline_data['auditor_login'] == user_login]

                if not auditor_data.empty:
                    # Filter incorrect cases
                    incorrect_cases = auditor_data[
                        auditor_data['qc2_judgement'].isin(['AUDITOR_INCORRECT', 'BOTH_INCORRECT'])
                    ]

                    if not incorrect_cases.empty:
                        # Count by subreason
                        subreason_counts = incorrect_cases['qc2_subreason'].value_counts()
                        total_incorrect = len(incorrect_cases)

                        # Calculate percentages
                        subreason_percentages = round((subreason_counts / total_incorrect * 100), 2)

                        # Store both counts and percentages
                        subreason_data[pipeline][week] = {
                            'counts': subreason_counts.to_dict(),
                            'percentages': subreason_percentages.to_dict(),
                            'total_incorrect': total_incorrect
                        }

        return subreason_data, None

    except Exception as e:
        return None, f"Error calculating QC2 subreason analysis: {str(e)}"


def generate_subreason_tables(subreason_data, weeks_to_show):
    """Generate HTML tables for QC2 subreason counts and percentages"""
    # Count table
    count_html = f'''
    <table class="data-table">
        <thead>
            <tr>
                <th>Subreason</th>'''

    # Add week headers
    for week in weeks_to_show:
        count_html += f'<th>{week}</th>'
    count_html += "</tr></thead><tbody>"

    # Percentage table
    percentage_html = f'''
    <table class="data-table">
        <thead>
            <tr>
                <th>Subreason</th>'''

    # Add week headers for percentage table
    for week in weeks_to_show:
        percentage_html += f'<th>{week}</th>'
    percentage_html += "</tr></thead><tbody>"

    # Get all unique subreasons
    all_subreasons = set()
    for pipeline_data in subreason_data.values():
        for week_data in pipeline_data.values():
            all_subreasons.update(week_data['counts'].keys())

    # Add rows for each subreason
    for subreason in sorted(all_subreasons):
        # Count table row
        count_html += f'<tr><td>{subreason}</td>'

        # Percentage table row
        percentage_html += f'<tr><td>{subreason}</td>'

        for week in weeks_to_show:
            # Counts
            count = 0
            total_incorrect = 0
            for pipeline_data in subreason_data.values():
                if week in pipeline_data:
                    if subreason in pipeline_data[week]['counts']:
                        count += pipeline_data[week]['counts'][subreason]
                    total_incorrect += pipeline_data[week]['total_incorrect']

            if count > 0:
                count_html += f'<td class="highlight-negative">{count}</td>'
            else:
                count_html += '<td>0</td>'

            # Percentages
            percentage = (count / total_incorrect * 100) if total_incorrect > 0 else 0
            if percentage > 0:
                if percentage >= 20:
                    perc_class = "highlight-negative"
                elif percentage >= 10:
                    perc_class = "highlight-neutral"
                else:
                    perc_class = "highlight-positive"
                percentage_html += f'<td class="{perc_class}">{percentage:.1f}%</td>'
            else:
                percentage_html += '<td>0.0%</td>'

        count_html += "</tr>"
        percentage_html += "</tr>"

    count_html += "</tbody></table>"
    percentage_html += "</tbody></table>"

    return count_html, percentage_html


def send_html_email_to_auditor_with_all_week_benchmarks(data_file, quality_file, auditor_login, manager_email,
                                                        weeks_to_show=None, test_mode=True):
    """Send HTML-based email to a specific auditor with benchmarks for ALL weeks"""
    try:
        # Calculate productivity data for this auditor
        grouped_data, error = calculate_productivity_data_optimized(data_file, auditor_login)

        if error:
            print(f"Error for {auditor_login}: {error}")
            return False

        if not grouped_data:
            print(f"No productivity data found for {auditor_login}")
            return False

        # Calculate all auditors' productivity for percentile calculation
        all_productivity_data, all_error = calculate_all_auditors_productivity_optimized(data_file, weeks_to_show)

        weekly_pipeline_percentiles = {}
        weekly_overall_percentiles = {}
        if all_productivity_data and not all_error and weeks_to_show:
            weekly_pipeline_percentiles, weekly_overall_percentiles = calculate_percentiles_for_all_weeks_optimized(
                all_productivity_data, weeks_to_show)

        # Quality: prepare data if file is provided
        quality_data = {}
        weekly_quality_percentiles = {}
        if quality_file and os.path.exists(quality_file) and weeks_to_show:
            # QC2 Subreason Analysis
            subreason_data = {}
            subreason_count_table = ""
            subreason_percent_table = ""
            subreason_data, subreason_error = calculate_qc2_subreason_analysis(quality_file, auditor_login,
                                                                               weeks_to_show)
            if not subreason_error and subreason_data:
                subreason_count_table, subreason_percent_table = generate_subreason_tables(subreason_data,
                                                                                           weeks_to_show)
            all_quality_data, quality_error = calculate_all_auditors_quality_optimized(quality_file, weeks_to_show)
            if not quality_error and auditor_login in all_quality_data:
                quality_data = all_quality_data[auditor_login]
                weekly_quality_percentiles, _ = calculate_quality_percentiles_for_all_weeks(
                    all_quality_data, weeks_to_show)
        quality_table = generate_quality_table(quality_data, weekly_quality_percentiles, weeks_to_show)
        # Generate tables
        percentile_table = generate_percentile_table(grouped_data, weekly_pipeline_percentiles, weeks_to_show)
        productivity_table = generate_productivity_table(grouped_data, weekly_pipeline_percentiles, weeks_to_show)
        quality_table = generate_quality_table(quality_data, weekly_quality_percentiles, weeks_to_show)
        quality_percentile_table = generate_quality_percentile_table(quality_data, weekly_quality_percentiles,
                                                                     weeks_to_show)

        # Create email
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = f"{auditor_login}@amazon.com"
        if manager_email:
            mail.CC = manager_email
        mail.Subject = "RDR Productivity & Quality Metrics Report"

        # Create complete HTML body
        html_body = f"""
        <html>
        <head>
          <style>
            body {{
              font-family: Arial, sans-serif;
              color: #333;
              margin: 20px;
            }}
            .footer {{
              margin-top: 20px;
              padding: 10px;
              background-color: #f8f9fa;
            }}
            .percentile-legend {{
              margin-top: 10px;
              font-style: italic;
            }}
            table {{
              border-collapse: collapse;
              margin: 10px 0;
              width: 100%;
            }}
            th {{
              background-color: #2c3e50;
              color: white;
              padding: 12px 15px;
              text-align: left;
              font-weight: bold;
              border: 1px solid #34495e;
            }}
            td {{
              padding: 12px 15px;
              border: 1px solid #dee2e6;
            }}
            tr:nth-child(even) {{
              background-color: #f8f9fa;
            }}
            tr:nth-child(odd) {{
              background-color: white;
            }}
            tr:first-child {{
              background-color: #2c3e50;
            }}
          </style>
        </head>
        <body>
          <h2>RDR Productivity & Quality Metrics Report</h2>
          <p>Hi {auditor_login},</p>
          <p>Please find your RDR Productivity & Quality Metrics report below.</p>

          <table style="width: 100%; border-spacing: 10px 10px;">
            <tr>
              <td style="vertical-align: top; width: 50%;">
                <h3>Productivity</h3>
                {productivity_table}
              </td>
              <td style="vertical-align: top; width: 50%;">
                <h3>Quality</h3>
                {percentile_table}
              </td>
            </tr>
            <tr>
              <td style="vertical-align: top; width: 50%;">
                {quality_table}
              </td>
              <td style="vertical-align: top; width: 50%;">
                {quality_percentile_table}
              </td>
            </tr>
            <tr>
              <td style="vertical-align: top; width: 50%;">
                {subreason_count_table}
              </td>
              <td style="vertical-align: top; width: 50%;">
                {subreason_percent_table}
              </td>
            </tr>
          </table>

          <div class="percentile-legend">
            <p>Percentiles are calculated based on the latest week's data across all auditors for each pipeline.</p>
          </div>

          <div class="footer">
            <p><strong>Note:</strong> If you have any questions, please reach out to your manager.</p>
            <p>Best regards,<br>System Generated Report</p>
          </div>
        </body>
        </html>
        """
        mail.HTMLBody = html_body

        # Send or display
        if test_mode:
            mail.Display()
        else:
            mail.Send()

        return True

    except Exception as e:
        print(f"Error sending email to {auditor_login}: {str(e)}")
        return False


class EnhancedEmailSenderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Enhanced Excel Data-to-Email Sender")
        self.configure(bg="#f0f0f0")
        self.create_widgets()
        self.bind_mouse_wheel()
        self.update_idletasks()
        width = 500  # Match the width used in create_widgets
        height = min(800, self.winfo_reqheight())  # Limit maximum height
        self.geometry(f"{width}x{height}")

    def bind_mouse_wheel(self):  # Add this method
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def create_widgets(self):
        # Create a canvas with scrollbar
        self.canvas = tk.Canvas(self, width=500)  # Reduced width to match screenshot
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, width=500)  # Same width as canvas

        # Configure the canvas
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=500)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        # Pack the scrollbar and canvas
        self.grid_columnconfigure(0, weight=1)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Style configuration
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TLabel", padding=10, background="#f0f0f0", font=("Helvetica", 10))
        style.configure("TButton", padding=10, relief="flat")
        style.configure("Primary.TButton", background="#007bff", foreground="white")
        style.configure("Secondary.TButton", background="#6c757d", foreground="white")
        style.configure("Success.TButton", background="#28a745", foreground="white")

        # Main title
        title_label = ttk.Label(self.scrollable_frame, text="Data-to-Email Productivity Report Sender",
                                font=("Helvetica", 14, "bold"))
        title_label.pack(pady=10)

        # Data file selection
        data_frame = ttk.LabelFrame(self.scrollable_frame, text="Data File Selection", padding=15)
        data_frame.pack(pady=10, padx=20, fill="x")

        ttk.Label(data_frame, text="Select Data Dump Excel File:").pack(anchor="w")
        self.data_path = tk.StringVar()
        self.data_entry = ttk.Entry(data_frame, textvariable=self.data_path, width=60)
        self.data_entry.pack(pady=5, fill="x")
        ttk.Button(data_frame, text="Browse", command=self.browse_data_file).pack(pady=5)

        # Quality file selection (add after data_frame)
        quality_frame = ttk.LabelFrame(self.scrollable_frame, text="Quality File Selection (Optional)", padding=15)
        quality_frame.pack(pady=10, padx=20, fill="x")

        ttk.Label(quality_frame, text="Select Quality Dump Excel File:").pack(anchor="w")
        self.quality_path = tk.StringVar()
        self.quality_entry = ttk.Entry(quality_frame, textvariable=self.quality_path, width=60)
        self.quality_entry.pack(pady=5, fill="x")
        ttk.Button(quality_frame, text="Browse", command=self.browse_quality_file).pack(pady=5)

        # Manager mapping file selection
        manager_frame = ttk.LabelFrame(self.scrollable_frame, text="Manager Mapping (Optional)", padding=15)
        manager_frame.pack(pady=10, padx=20, fill="x")

        ttk.Label(manager_frame, text="Select Manager Mapping File (for bulk sending):").pack(anchor="w")
        self.manager_path = tk.StringVar()
        self.manager_entry = ttk.Entry(manager_frame, textvariable=self.manager_path, width=60)
        self.manager_entry.pack(pady=5, fill="x")
        ttk.Button(manager_frame, text="Browse", command=self.browse_manager_file).pack(pady=5)

        # Week selection
        week_frame = ttk.LabelFrame(self.scrollable_frame, text="Week Selection (Optional)", padding=15)
        week_frame.pack(pady=10, padx=20, fill="x")

        ttk.Label(week_frame, text="Specify weeks to include (comma-separated, e.g., 20,21,22,23):").pack(anchor="w")
        self.weeks_entry = ttk.Entry(week_frame, width=30)
        self.weeks_entry.pack(pady=5)
        ttk.Label(week_frame,
                  text="Leave empty to include all available weeks. Percentiles calculated from latest week.",
                  font=("Helvetica", 8)).pack(anchor="w")

        # Single user input
        single_frame = ttk.LabelFrame(self.scrollable_frame, text="Single User Mode", padding=15)
        single_frame.pack(pady=10, padx=20, fill="x")

        ttk.Label(single_frame, text="User Login (for single email):").pack(anchor="w")
        self.single_user = tk.StringVar()
        self.single_user_entry = ttk.Entry(single_frame, textvariable=self.single_user, width=30)
        self.single_user_entry.pack(pady=5)

        # Send mode selection
        mode_frame = ttk.LabelFrame(self.scrollable_frame, text="Send Mode", padding=15)
        mode_frame.pack(pady=10, padx=20, fill="x")

        self.send_mode = tk.StringVar(value="preview")
        ttk.Radiobutton(mode_frame, text="Preview Mode (Display emails)",
                        variable=self.send_mode, value="preview").pack(anchor="w")
        ttk.Radiobutton(mode_frame, text="Send Mode (Actually send emails)",
                        variable=self.send_mode, value="send").pack(anchor="w")

        # Action buttons
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.pack(pady=20)

        self.single_button = ttk.Button(button_frame, text="Send Single Email",
                                        command=self.send_single_email, style="Primary.TButton")
        self.single_button.pack(side="left", padx=10)

        self.bulk_button = ttk.Button(button_frame, text="Send to All Auditors",
                                      command=self.send_bulk_emails_optimized, style="Success.TButton")
        self.bulk_button.pack(side="left", padx=10)

        # Status text
        self.status_text = tk.Text(self.scrollable_frame, height=8, width=60)
        self.status_text.pack(pady=10, padx=20, fill="both", expand=True)

        scrollbar_status = ttk.Scrollbar(self.status_text)
        scrollbar_status.pack(side="right", fill="y")
        self.status_text.config(yscrollcommand=scrollbar_status.set)
        scrollbar_status.config(command=self.status_text.yview)

    def browse_data_file(self):
        filetypes = [("Excel files", "*.xlsx *.xlsm")]
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.data_path.set(filename)

    def browse_quality_file(self):
        filetypes = [("Excel files", "*.xlsx *.xlsm")]
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.quality_path.set(filename)

    def browse_manager_file(self):
        filetypes = [("Excel files", "*.xlsx *.xlsm")]
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.manager_path.set(filename)

    def get_weeks_list(self):
        """Parse weeks from entry field"""
        weeks_text = self.weeks_entry.get().strip()
        if not weeks_text:
            return None

        try:
            weeks = [int(w.strip()) for w in weeks_text.split(',')]
            return weeks
        except ValueError:
            messagebox.showerror("Error", "Invalid week format. Use comma-separated numbers (e.g., 20,21,22)")
            return None

    def log_status(self, message):
        """Add message to status text widget"""
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.update()

    def send_single_email(self):
        """Send single email with full HTML content, QC2 subreason, and percentile data"""
        user_login = self.single_user.get().strip()
        if not user_login:
            messagebox.showerror("Error", "Please enter a user login for single email mode.")
            return

        if not self.data_path.get():
            messagebox.showerror("Error", "Please select a data file.")
            return

        weeks = self.get_weeks_list()
        if not weeks:
            return

        quality_file = self.quality_path.get().strip()
        has_quality = bool(quality_file)

        # Calculate productivity data
        grouped_data, error = calculate_productivity_data_optimized(self.data_path.get(), user_login)
        if error:
            self.log_status(f"Error for {user_login}: {error}")
            return
        if not grouped_data:
            self.log_status(f"No productivity data found for {user_login}")
            return

        # Calculate all auditors' productivity for percentile calculation
        all_productivity_data, all_error = calculate_all_auditors_productivity_optimized(self.data_path.get(), weeks)
        weekly_pipeline_percentiles = {}
        weekly_overall_percentiles = {}
        if all_productivity_data and not all_error and weeks:
            weekly_pipeline_percentiles, weekly_overall_percentiles = calculate_percentiles_for_all_weeks_optimized(
                all_productivity_data, weeks
            )

        # Quality and Subreason Data
        quality_data = {}
        weekly_quality_percentiles = {}
        subreason_count_table = ""
        subreason_percent_table = ""

        if has_quality and os.path.exists(quality_file):
            subreason_data, subreason_error = calculate_qc2_subreason_analysis(quality_file, user_login, weeks)
            if not subreason_error and subreason_data:
                subreason_count_table, subreason_percent_table = generate_subreason_tables(subreason_data, weeks)

            all_quality_data, quality_error = calculate_all_auditors_quality_optimized(quality_file, weeks)
            if not quality_error and user_login in all_quality_data:
                quality_data = all_quality_data[user_login]
                weekly_quality_percentiles, _ = calculate_quality_percentiles_for_all_weeks(all_quality_data, weeks)

        # Generate tables
        percentile_table = generate_percentile_table(grouped_data, weekly_pipeline_percentiles, weeks)
        productivity_table = generate_productivity_table(grouped_data, weekly_pipeline_percentiles, weeks)
        quality_table = generate_quality_table(quality_data, weekly_quality_percentiles,
                                               weeks) if has_quality else "<p>No quality data available</p>"
        quality_percentile_table = generate_quality_percentile_table(quality_data, weekly_quality_percentiles,
                                                                     weeks) if has_quality else "<p>No quality data available</p>"

        test_mode = self.send_mode.get() == "preview"

        # Create email
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = f"{user_login}@amazon.com"
        mail.Subject = "RDR Productivity & Quality Metrics Report"

        # Email body with updated layout
        html_body = f"""
        <html>
        <head>
            <style>
                body {{
                    font-family: 'Segoe UI', Arial, sans-serif;
                    color: #333;
                    margin: 20px;
                    line-height: 1.2;
                }}

                .email-container {{
                    max-width: 1200px;
                    margin: 0 auto;
                    padding: 20px;
                }}

                .header {{
                    margin-bottom: 30px;
                }}

                .content {{
                    margin-bottom: 40px;
                }}

                .footer {{
                    margin-top: 30px;
                    padding: 20px;
                    background-color: #f8f9fa;
                    border-radius: 8px;
                }}

                .section-title {{
                    color: #2c3e50;
                    margin: 25px 0 15px;
                    font-size: 1.2em;
                }}

                table {{
                    border-collapse: separate;
                    border-spacing: 0;
                    width: 50%;
                    margin: 15px auto;
                    background: white;
                    border: 1px solid #ddd;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    font-size: 0.85em;
                }}

                th {{
                    background-color: #2c3e50;
                    color: white;
                    padding: 10px 12px;
                    text-align: center;
                    font-weight: 600;
                    text-transform: uppercase;
                    font-size: 0.9em;
                    letter-spacing: 0.5px;
                    border-right: 1px solid #34495e;
                    border-bottom: 1px solid #34495e;
                }}

                td {{
                    padding: 8px 12px;
                    text-align: center;
                    border-right: 1px solid #ddd;
                    border-bottom: 1px solid #ddd;
                    color: #333;
                }}

                th:last-child,
                td:last-child {{
                    border-right: none;
                }}

                tr:last-child td {{
                    border-bottom: none;
                }}

                tr:nth-child(even) {{
                    background-color: #f8f9fa;
                }}

                .highlight-positive {{
                    color: #28a745;
                    font-weight: 600;
                }}

                .highlight-negative {{
                    color: #dc3545;
                    font-weight: 600;
                }}

                .highlight-neutral {{
                    color: #17a2b8;
                    font-weight: 600;
                }}

                .grid-container {{
                    display: flex;
                    flex-direction: row;
                    gap: 20px;
                    margin: 20px 0;
                    justify-content: space-between;
                }}

                .grid-container > div {{
                    flex: 1;
                }}

                .legend {{
                    background-color: #f8f9fa;
                    padding: 15px;
                    border-radius: 8px;
                    margin: 20px 0;
                    font-size: 0.9em;
                }}
            </style>
        </head>
        <body>
            <div class="email-container">
                <div class="header">
                    <h1>RDR Productivity & Quality Metrics Report</h1>
                </div>

                <div class="content">
                    <div class="greeting">
                        <p>Hi {user_login},</p>
                        <p>Your latest performance metrics are ready for review. Here's your comprehensive report:</p>
                    </div>

                    <div class="grid-container">
                        <div>
                            <h3 class="section-title">Productivity Metrics</h3>
                            {productivity_table}
                        </div>
                        <div>
                            <h3 class="section-title">Quality Metrics</h3>
                            {quality_table}
                        </div>
                    </div>

                    <div class="grid-container">
                        <div>
                            <h3 class="section-title">Productivity Benchmarks</h3>
                            {percentile_table}
                        </div>
                        <div>
                            <h3 class="section-title">Quality Benchmarks</h3>
                            {quality_percentile_table}
                        </div>
                    </div>

                    <div class="grid-container">
                        <div>
                            <h3 class="section-title">QC2 Subreason Analysis - Counts</h3>
                            {subreason_count_table}
                        </div>
                        <div>
                            <h3 class="section-title">QC2 Subreason Analysis - Percentages</h3>
                            {subreason_percent_table}
                        </div>
                    </div>

                    <div class="legend">
                        <p><strong>Note:</strong> Percentiles are calculated based on the latest week's data across all auditors for each pipeline.</p>
                    </div>
                </div>

                <div class="footer">
                    <p><strong>Need Help?</strong> If you have any questions about your metrics, please reach out to your manager.</p>
                    <p>Best regards,<br><strong>Automated Reporting System</strong></p>
                </div>
            </div>
        </body>
        </html>
        """

        # IMPORTANT: You also need to update your table generation code to use the new CSS classes
        # Here's how to modify your existing table generation functions:

        def generate_styled_table(data, headers, table_class="data-table"):
            """
            Generate HTML table with the new styling classes
            """
            html = f'<table class="{table_class}">'

            # Add headers
            html += '<thead><tr>'
            for header in headers:
                html += f'<th>{header}</th>'
            html += '</tr></thead>'

            # Add data rows
            html += '<tbody>'
            for row in data:
                html += '<tr>'
                for cell in row:
                    # You can add conditional classes here based on your data
                    cell_class = ""
                    if isinstance(cell, (int, float)):
                        if cell > 0:  # Example: positive numbers in green
                            cell_class = ' class="highlight-positive"'
                        elif cell < 0:  # Example: negative numbers in red
                            cell_class = ' class="highlight-negative"'
                        else:
                            cell_class = ' class="highlight-neutral"'

                    html += f'<td{cell_class}>{cell}</td>'
                html += '</tr>'
            html += '</tbody>'

            html += '</table>'
            return html
        mail.HTMLBody = html_body

        # Send or preview
        if test_mode:
            mail.Display()
            self.log_status("✓ Preview email opened in Outlook!")
            messagebox.showinfo("Success", "Preview email opened in Outlook!")
        else:
            mail.Send()
            self.log_status("✓ Email sent successfully!")
            messagebox.showinfo("Success", "Email sent successfully!")

    def send_bulk_emails_optimized(self):
        """Optimized bulk email sending - calculate all data once"""
        if not self.data_path.get() or not self.manager_path.get():
            messagebox.showerror("Error", "Please select both data file and manager mapping file.")
            return

        weeks = self.get_weeks_list()
        if weeks == []:
            return

        quality_file = self.quality_path.get().strip()
        has_quality = bool(quality_file)

        try:
            # Read manager mapping once
            manager_df = pd.read_excel(self.manager_path.get())
            auditors = manager_df['loginname'].dropna().unique()

            # Calculate ALL productivity data once at the start
            self.log_status("Pre-calculating productivity data for all auditors...")
            all_productivity_data, all_error = calculate_all_auditors_productivity_optimized(
                self.data_path.get(), weeks)

            if all_error:
                self.log_status(f"✗ Error calculating productivity data: {all_error}")
                return

            # Calculate productivity percentiles once
            pipeline_percentiles, overall_percentiles = calculate_percentiles_for_all_weeks_optimized(
                all_productivity_data, weeks)

            # Calculate quality data if file is provided
            all_quality_data = {}
            weekly_quality_percentiles = {}
            if has_quality:
                self.log_status("Pre-calculating quality data for all auditors...")
                all_quality_data, quality_error = calculate_all_auditors_quality_optimized(quality_file, weeks)
                if not quality_error:
                    weekly_quality_percentiles, _ = calculate_quality_percentiles_for_all_weeks(all_quality_data, weeks)

            self.log_status("Pre-calculation complete. Starting email generation...")

            success_count = 0
            failed_count = 0
            test_mode = self.send_mode.get() == "preview"

            for i, auditor_login in enumerate(auditors, 1):
                self.log_status(f"Processing {i}/{len(auditors)}: {auditor_login}")

                # Get individual auditor data from pre-calculated results
                if auditor_login in all_productivity_data:
                    grouped_data = all_productivity_data[auditor_login]

                    # Get quality data for this auditor
                    quality_data = all_quality_data.get(auditor_login, {}) if has_quality else {}

                    # Get supervisor email
                    supervisor_info = manager_df[manager_df['loginname'] == auditor_login]['supervisorloginname'].iloc[
                        0]
                    manager_email = f"{supervisor_info}@amazon.com" if pd.notna(supervisor_info) else None

                    # Generate tables
                    percentile_table = generate_percentile_table(grouped_data, pipeline_percentiles, weeks)
                    productivity_table = generate_productivity_table(grouped_data, pipeline_percentiles, weeks)
                    quality_table = generate_quality_table(quality_data, weekly_quality_percentiles,
                                                           weeks) if has_quality else "<p>No quality data available</p>"
                    quality_percentile_table = generate_quality_percentile_table(quality_data,
                                                                                 weekly_quality_percentiles,
                                                                                 weeks) if has_quality else "<p>No quality data available</p>"

                    # Generate and send email
                    if self.generate_and_send_email_with_quality(
                            auditor_login, manager_email, weeks, test_mode,
                            productivity_table, quality_table, percentile_table, quality_percentile_table):
                        success_count += 1
                        self.log_status(f"✓ Email {'previewed' if test_mode else 'sent'} to {auditor_login}")
                    else:
                        failed_count += 1
                        self.log_status(f"✗ Failed to send email to {auditor_login}")
                else:
                    failed_count += 1
                    self.log_status(f"✗ No data found for {auditor_login}")

            # Summary
            summary_msg = f"Bulk operation completed!\nSuccess: {success_count}\nFailed: {failed_count}"
            self.log_status("\n" + "=" * 50)
            self.log_status(summary_msg)

        except Exception as e:
            error_msg = f"Error in bulk email process: {str(e)}"
            self.log_status(f"✗ {error_msg}")
            messagebox.showerror("Error", error_msg)


def generate_and_send_email_with_quality(self, auditor_login, manager_email, weeks, test_mode,
                                         productivity_table, quality_table, percentile_table, quality_percentile_table):
    """Generate and send email with pre-generated tables"""
    try:
        # Create email
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = f"{auditor_login}@amazon.com"
        if manager_email:
            mail.CC = manager_email
        mail.Subject = "RDR Productivity & Quality Metrics Report"

        # Create complete HTML body
        html_body = f"""
        <html>
        <head>
          <style>
            body {{
              font-family: Arial, sans-serif;
              color: #333;
              margin: 20px;
            }}
            .footer {{
              margin-top: 20px;
              padding: 10px;
              background-color: #f8f9fa;
            }}
            .percentile-legend {{
              margin-top: 10px;
              font-style: italic;
            }}
            table {{
              border-collapse: collapse;
              margin: 10px 0;
              width: 100%;
            }}
            th {{
              background-color: #2c3e50;
              color: white;
              padding: 12px 15px;
              text-align: left;
              font-weight: bold;
              border: 1px solid #34495e;
            }}
            td {{
              padding: 12px 15px;
              border: 1px solid #dee2e6;
            }}
            tr:nth-child(even) {{
              background-color: #f8f9fa;
            }}
            tr:nth-child(odd) {{
              background-color: white;
            }}
            tr:first-child {{
              background-color: #2c3e50;
            }}
          </style>
        </head>
        <body>
          <h2>RDR Productivity & Quality Metrics Report</h2>
          <p>Hi {auditor_login},</p>
          <p>Please find your RDR Productivity & Quality Metrics report below.</p>

          <table style="width: 100%; border-spacing: 10px 10px;">
            <tr>
              <td style="vertical-align: top; width: 50%;">
                <h3>Productivity</h3>
                {productivity_table}
              </td>
              <td style="vertical-align: top; width: 50%;">
                <h3>Quality</h3>
                {percentile_table}
              </td>
            </tr>
            <tr>
              <td style="vertical-align: top; width: 50%;">
                {quality_table}
              </td>
              <td style="vertical-align: top; width: 50%;">
                {quality_percentile_table}
              </td>
            </tr>
            <tr>
              <td style="vertical-align: top; width: 50%;">
                {quality_table}
              </td>
              <td style="vertical-align: top; width: 50%;">
                {quality_percentile_table}
              </td>
            </tr>

          </table>

          <div class="percentile-legend">
            <p>Percentiles are calculated based on the latest week's productivity across all auditors for each pipeline.</p>
          </div>

          <div class="footer">
            <p><strong>Note:</strong> If you have any questions, please reach out to your manager.</p>
            <p>Best regards,<br>System Generated Report</p>
          </div>
        </body>
        </html>
        """
        mail.HTMLBody = html_body

        # Send or display
        if test_mode:
            mail.Display()
        else:
            mail.Send()

        return True

    except Exception as e:
        print(f"Error sending email to {auditor_login}: {str(e)}")
        return False


if __name__ == "__main__":
    app = EnhancedEmailSenderApp()
    app.mainloop()