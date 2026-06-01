"""
Central Error Logging Utility for Validation Tools

This module provides functions for all processors to log errors
in a consistent format to a central Excel file.

Usage:
    from utils.data_logger import log_errors

    log_errors(
        hall="JPB15",
        rack_type="T0-to-Host",
        building="B11",
        error_category="Mismatch",
        count=14,
        source_file="JPB15_T0_Host_B11.xlsx"
    )
"""

import pandas as pd
from datetime import datetime
import os
import time
import shutil
from pathlib import Path

# Central log file location (relative to repo root)
LOG_FILE = Path(__file__).parent.parent / "data" / "validation_error_log.xlsx"

def ensure_log_exists():
    """Create the log file with proper columns if it doesn't exist."""
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    if not LOG_FILE.exists():
        df = pd.DataFrame(columns=[
            "timestamp",
            "hall",
            "rack_type",
            "building",
            "error_category",
            "count",
            "source_file",
            "processed_by"
        ])
        df.to_excel(LOG_FILE, index=False)
        print(f"Created new error log at: {LOG_FILE}")

def log_errors(
    hall: str,
    rack_type: str,
    building: str,
    error_category: str,
    count: int,
    source_file: str = "",
    processed_by: str = "system"
):
    """
    Log error counts from a validation processor.
    
    Args:
        hall: e.g. "JPB15", "JPB19", "SYD20"
        rack_type: e.g. "T0-to-Host", "T1-to-T0", "CFAB", "QFAB", "HOPS"
        building: Building or rack identifier (e.g. "B11", "R12101")
        error_category: Type of error (e.g. "Mismatch", "Downlink", "Optics", "FEC")
        count: Number of occurrences
        source_file: Original filename processed
        processed_by: User or system that processed it
    """
    ensure_log_exists()
    
    new_row = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "hall": hall,
        "rack_type": rack_type,
        "building": building,
        "error_category": error_category,
        "count": count,
        "source_file": source_file,
        "processed_by": processed_by
    }
    
    # Reliable append using openpyxl with retries (Windows-friendly)
    from openpyxl import load_workbook, Workbook

    max_retries = 5
    for attempt in range(max_retries):
        wb = None
        try:
            if LOG_FILE.exists():
                wb = load_workbook(LOG_FILE)
                ws = wb.active
            else:
                wb = Workbook()
                ws = wb.active
                ws.title = "Error Log"
                headers = ["timestamp", "hall", "rack_type", "building", 
                           "error_category", "count", "source_file", "processed_by"]
                ws.append(headers)

            # Append the new row
            ws.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                hall,
                rack_type,
                building,
                error_category,
                count,
                source_file,
                processed_by
            ])

            wb.save(LOG_FILE)
            print(f"Logged error → {LOG_FILE}")
            return True

        except PermissionError:
            if wb:
                try:
                    wb.close()
                except:
                    pass
            print(f"Log file locked, retrying... ({attempt + 1}/{max_retries})")
            time.sleep(0.8 * (attempt + 1))  # increasing backoff

        except Exception as e:
            if wb:
                try:
                    wb.close()
                except:
                    pass
            print(f"Failed to write to error log: {e}")
            return False

    print("Failed to write to error log after multiple retries (file probably locked by Excel).")
    return False

def get_error_log():
    """Return the full error log as a DataFrame."""
    ensure_log_exists()
    return pd.read_excel(LOG_FILE)

def get_summary_by_hall_and_type():
    """Return aggregated summary grouped by Hall and Rack Type."""
    df = get_error_log()
    if df.empty:
        return pd.DataFrame()
    
    summary = df.groupby(["hall", "rack_type"]).agg({
        "count": "sum",
        "building": "nunique"
    }).reset_index()
    summary.columns = ["Hall", "Rack Type", "Total Errors", "Unique Buildings"]
    return summary.sort_values("Total Errors", ascending=False)