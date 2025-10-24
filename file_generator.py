"""
ITBI Data Collection - File Generator
Generates DATA and META Excel files for each ISIN
"""

import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import os

from config import (
    TIME_SERIES_ORDER, METADATA_COLUMNS, DATA_FILE_PATTERN, 
    META_FILE_PATTERN, OUTPUT_DIR, TIME_SERIES_MAPPING
)


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def log_debug(message: str, prefix: str = "INFO"):
    """Log debug messages with timestamp"""
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    print(f"[{timestamp}] [{prefix}] {message}")


def ensure_output_directory():
    """Create output directory if it doesn't exist"""
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    log_debug(f"Output directory ensured: {OUTPUT_DIR}")


def create_meta_file(isin, metadata_rows, current_month):
    """
    Create META file for an ISIN
    
    Args:
        isin (str): ISIN code
        metadata_rows (list): List of metadata dictionaries
        current_month (str): Current month string (YYYY-MM)
        
    Returns:
        str: Path to created META file
    """
    # Create a DataFrame from metadata rows
    meta_df = pd.DataFrame(metadata_rows)
    
    # Ensure all required columns are present
    for col in METADATA_COLUMNS:
        if col not in meta_df.columns:
            meta_df[col] = ""
    
    # Reorder columns to match the required format
    meta_df = meta_df[METADATA_COLUMNS]
    
    # Generate filename - use .xlsx instead of .xls to avoid engine issues
    timestamp = current_month.replace("-", "")
    filename = META_FILE_PATTERN.format(isin=isin, timestamp=timestamp)
    filepath = os.path.join(OUTPUT_DIR, filename)
    
    # Write to Excel file with openpyxl engine
    meta_df.to_excel(filepath, index=False, engine='openpyxl')
    
    log_debug(f"Created META file: {filepath}")
    return filepath


def create_data_file(isin, time_series, description, current_month):
    """
    Create DATA file for an ISIN
    
    Args:
        isin (str): ISIN code
        time_series (dict): Dictionary of time series DataFrames
        description (str): ISIN description
        current_month (str): Current month string (YYYY-MM)
        
    Returns:
        str: Path to created DATA file
    """
    # Create a new DataFrame for the data file
    # First, get all unique dates and sort them
    all_dates = set()
    for ts_type, ts_df in time_series.items():
        all_dates.update(ts_df['date'].tolist())
    
    sorted_dates = sorted(all_dates)
    
    # Create the main data structure
    data_rows = []
    
    # Add header row (empty for column A)
    header_row = [""]
    for ts_type in TIME_SERIES_ORDER:
        code = f"{isin}.{TIME_SERIES_MAPPING[ts_type]['suffix']}"
        header_row.append(code)
    data_rows.append(header_row)
    
    # Add description row
    desc_row = [""]
    for ts_type in TIME_SERIES_ORDER:
        desc = f"ISIN:{isin};{description}:{TIME_SERIES_MAPPING[ts_type]['description']}"
        desc_row.append(desc)
    data_rows.append(desc_row)
    
    # Add data rows
    for date in sorted_dates:
        row = [date]
        for ts_type in TIME_SERIES_ORDER:
            ts_df = time_series[ts_type]
            # Find the value for this date
            value_row = ts_df[ts_df['date'] == date]
            if not value_row.empty:
                value = value_row.iloc[0]['value']
                # Handle NaN values
                if pd.isna(value):
                    row.append("")
                else:
                    row.append(value)
            else:
                row.append("")
        data_rows.append(row)
    
    # Create DataFrame
    columns = ["DATE"] + [f"{isin}.{TIME_SERIES_MAPPING[ts_type]['suffix']}" for ts_type in TIME_SERIES_ORDER]
    data_df = pd.DataFrame(data_rows, columns=columns)
    
    # Generate filename - use .xlsx instead of .xls to avoid engine issues
    timestamp = current_month.replace("-", "")
    filename = DATA_FILE_PATTERN.format(isin=isin, timestamp=timestamp)
    filepath = os.path.join(OUTPUT_DIR, filename)
    
    # Write to Excel file with openpyxl engine
    data_df.to_excel(filepath, index=False, engine='openpyxl')
    
    log_debug(f"Created DATA file: {filepath}")
    return filepath


def generate_files_for_isin(isin, isin_data):
    """
    Generate DATA and META files for a single ISIN
    
    Args:
        isin (str): ISIN code
        isin_data (dict): Dictionary containing ISIN data
        
    Returns:
        tuple: (meta_file_path, data_file_path)
    """
    description = isin_data['description']
    time_series = isin_data['time_series']
    metadata = isin_data['metadata']
    current_month = isin_data['current_month']
    
    log_debug(f"\nGenerating files for ISIN: {isin}")
    
    # Create META file
    meta_file = create_meta_file(isin, metadata, current_month)
    
    # Create DATA file
    data_file = create_data_file(isin, time_series, description, current_month)
    
    return meta_file, data_file


def generate_all_files(parsed_data):
    """
    Generate DATA and META files for all ISINs
    
    Args:
        parsed_data (dict): Dictionary containing parsed data for all ISINs
        
    Returns:
        dict: Dictionary with file paths for each ISIN
    """
    log_debug("\n" + "="*80)
    log_debug("STARTING FILE GENERATION")
    log_debug("="*80 + "\n")
    
    ensure_output_directory()
    
    file_paths = {}
    
    for isin, isin_data in parsed_data.items():
        meta_file, data_file = generate_files_for_isin(isin, isin_data)
        file_paths[isin] = {
            'meta_file': meta_file,
            'data_file': data_file
        }
    
    log_debug("\n" + "="*80)
    log_debug("FILE GENERATION COMPLETE")
    log_debug("="*80)
    log_debug(f"Generated files for {len(file_paths)} ISINs")
    
    return file_paths


# =============================================================================
# TEST FUNCTION
# =============================================================================

def test_file_generator():
    """Test the file generator with parsed data"""
    # Import the parser to get test data
    from parser import test_parser
    
    # Get parsed data
    parsed_data = test_parser()
    
    if not parsed_data:
        log_debug("No parsed data available for testing", "ERROR")
        return None
    
    # Generate files
    file_paths = generate_all_files(parsed_data)
    
    if file_paths:
        # Print summary
        log_debug("\nGENERATED FILES SUMMARY:")
        for isin, paths in file_paths.items():
            log_debug(f"\nISIN: {isin}")
            log_debug(f"  META file: {paths['meta_file']}")
            log_debug(f"  DATA file: {paths['data_file']}")
    
    return file_paths


# =============================================================================
# VERIFICATION FUNCTION
# =============================================================================

def verify_generated_files(file_paths):
    """Verify that the generated files match the expected format"""
    log_debug("\n" + "="*80)
    log_debug("VERIFYING GENERATED FILES")
    log_debug("="*80 + "\n")
    
    for isin, paths in file_paths.items():
        log_debug(f"\nVerifying files for ISIN: {isin}")
        
        # Verify META file
        try:
            meta_df = pd.read_excel(paths['meta_file'])
            log_debug(f"  META file: {len(meta_df)} rows, {len(meta_df.columns)} columns")
            
            # Check if it has the expected structure
            if 'CODE' in meta_df.columns and 'DESCRIPTION' in meta_df.columns:
                log_debug("  META file structure: OK")
            else:
                log_debug("  META file structure: MISSING REQUIRED COLUMNS", "WARNING")
        except Exception as e:
            log_debug(f"  Error reading META file: {str(e)}", "ERROR")
        
        # Verify DATA file
        try:
            data_df = pd.read_excel(paths['data_file'])
            log_debug(f"  DATA file: {len(data_df)} rows, {len(data_df.columns)} columns")
            
            # Check if it has the expected structure
            if len(data_df.columns) >= 6:  # DATE + 5 time series
                log_debug("  DATA file structure: OK")
            else:
                log_debug("  DATA file structure: INSUFFICIENT COLUMNS", "WARNING")
        except Exception as e:
            log_debug(f"  Error reading DATA file: {str(e)}", "ERROR")
    
    log_debug("\nFile verification complete")


# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Main execution function"""
    print("\n" + "="*80)
    print("ITBI FILE GENERATOR")
    print("Banca d'Italia - Government Bond Auction Results")
    print("="*80 + "\n")
    
    file_paths = test_file_generator()
    
    if file_paths:
        verify_generated_files(file_paths)
        
        print("\n[SUCCESS] File generation completed!")
        print(f"Generated files for {len(file_paths)} ISINs")
        print(f"Total files created: {len(file_paths) * 2} (1 META + 1 DATA per ISIN)")
        return file_paths
    else:
        print("\n[FAILED] File generation failed!")
        return None


if __name__ == "__main__":
    result = main()
    if result:
        print("\nAll files have been generated successfully!")
        print(f"Check the '{OUTPUT_DIR}' directory for the output files.")