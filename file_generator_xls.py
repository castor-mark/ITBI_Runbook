"""
ITBI Data Collection - XLS File Generator (Fixed Version)
Generates DATA and META .xls files for each ISIN with proper formatting and ZIP packaging

FIXES:
1. Uses .xls format (xlwt library) instead of .xlsx
2. Proper file naming: ITBI_{ISIN}_TYPE_{YYYYMMDD}.xls
3. Correct DATA file structure (no "DATE" header in Row 1, Column A)
4. FREQUENCY set to 'M' (Monthly)
5. Creates ZIP file per ISIN
"""

import xlwt
from datetime import datetime, timedelta
from pathlib import Path
import os
import zipfile

from config import (
    TIME_SERIES_ORDER, METADATA_COLUMNS, DATA_FILE_PATTERN,
    META_FILE_PATTERN, ZIP_FILE_PATTERN, OUTPUT_DIR, TIME_SERIES_MAPPING,
    METADATA_DEFAULTS, FILENAME_DATE_FORMAT, BASE_URL
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


def get_timestamp_from_month(year_month_str):
    """
    Convert YYYY-MM format to YYYYMMDD format (using last day of month)

    Args:
        year_month_str (str): Date in YYYY-MM format

    Returns:
        str: Date in YYYYMMDD format
    """
    year, month = map(int, year_month_str.split('-'))
    # Get the last day of the month
    if month == 12:
        next_month = datetime(year + 1, 1, 1)
    else:
        next_month = datetime(year, month + 1, 1)

    last_day = next_month - timedelta(days=1)
    return last_day.strftime(FILENAME_DATE_FORMAT)


# =============================================================================
# META FILE CREATION
# =============================================================================

def create_meta_file_xls(isin, metadata_rows, current_month):
    """
    Create META file in .xls format using xlwt

    Args:
        isin (str): ISIN code
        metadata_rows (list): List of metadata dictionaries
        current_month (str): Current month string (YYYY-MM)

    Returns:
        str: Path to created META file
    """
    # Create workbook
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('META')

    # Write header row
    for col_idx, col_name in enumerate(METADATA_COLUMNS):
        ws.write(0, col_idx, col_name)

    # Write data rows
    for row_idx, metadata in enumerate(metadata_rows, start=1):
        for col_idx, col_name in enumerate(METADATA_COLUMNS):
            value = metadata.get(col_name, '')
            # Handle boolean values specially (keep as boolean, not string)
            if isinstance(value, bool):
                ws.write(row_idx, col_idx, value)
            # Convert other values to string for compatibility
            elif value or value == 0:
                ws.write(row_idx, col_idx, str(value))
            else:
                ws.write(row_idx, col_idx, '')

    # Generate filename with proper format
    timestamp = get_timestamp_from_month(current_month)
    filename = META_FILE_PATTERN.format(isin=isin, timestamp=timestamp)
    filepath = os.path.join(OUTPUT_DIR, filename)

    # Save file
    wb.save(filepath)

    log_debug(f"Created META file: {filepath}")
    return filepath


# =============================================================================
# DATA FILE CREATION
# =============================================================================

def create_data_file_xls(isin, time_series, description, current_month):
    """
    Create DATA file in .xls format using xlwt with CORRECT structure

    CORRECT FORMAT:
    Row 1: (blank) | CODE1 | CODE2 | CODE3 | CODE4 | CODE5
    Row 2: (blank) | DESC1 | DESC2 | DESC3 | DESC4 | DESC5
    Row 3: DATE1   | val1  | val2  | val3  | val4  | val5
    Row 4: DATE2   | val1  | val2  | val3  | val4  | val5

    Args:
        isin (str): ISIN code
        time_series (dict): Dictionary of time series DataFrames
        description (str): ISIN description
        current_month (str): Current month string (YYYY-MM)

    Returns:
        str: Path to created DATA file
    """
    # Create workbook
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('DATA')

    # Get all unique dates and sort them
    all_dates = set()
    for ts_type, ts_df in time_series.items():
        all_dates.update(ts_df['date'].tolist())

    sorted_dates = sorted(all_dates)

    # Row 1: TIME SERIES CODES (Column A is blank!)
    # CRITICAL FIX: Do NOT write "DATE" in column A, row 1
    ws.write(0, 0, '')  # Column A, Row 1 is BLANK

    for col_idx, ts_type in enumerate(TIME_SERIES_ORDER, start=1):
        code = f"{isin}.{TIME_SERIES_MAPPING[ts_type]['suffix']}"
        ws.write(0, col_idx, code)

    # Row 2: TIME SERIES DESCRIPTIONS (Column A is blank!)
    ws.write(1, 0, '')  # Column A, Row 2 is BLANK

    for col_idx, ts_type in enumerate(TIME_SERIES_ORDER, start=1):
        desc = f"ISIN:{isin};{description}:{TIME_SERIES_MAPPING[ts_type]['description']}"
        ws.write(1, col_idx, desc)

    # Row 3+: DATA ROWS (Column A has dates)
    for row_idx, date in enumerate(sorted_dates, start=2):
        # Write date in Column A
        ws.write(row_idx, 0, date)

        # Write values for each time series
        for col_idx, ts_type in enumerate(TIME_SERIES_ORDER, start=1):
            ts_df = time_series[ts_type]
            # Find the value for this date
            value_rows = ts_df[ts_df['date'] == date]

            if not value_rows.empty:
                value = value_rows.iloc[0]['value']
                # Handle NaN/None values
                if value is None or (isinstance(value, float) and value != value):  # NaN check
                    ws.write(row_idx, col_idx, '')
                else:
                    ws.write(row_idx, col_idx, value)
            else:
                ws.write(row_idx, col_idx, '')

    # Generate filename with proper format
    timestamp = get_timestamp_from_month(current_month)
    filename = DATA_FILE_PATTERN.format(isin=isin, timestamp=timestamp)
    filepath = os.path.join(OUTPUT_DIR, filename)

    # Save file
    wb.save(filepath)

    log_debug(f"Created DATA file: {filepath}")
    return filepath


# =============================================================================
# ZIP PACKAGING
# =============================================================================

def create_zip_package(isin, meta_file, data_file, current_month):
    """
    Create a ZIP file containing both META and DATA files for an ISIN

    Args:
        isin (str): ISIN code
        meta_file (str): Path to META file
        data_file (str): Path to DATA file
        current_month (str): Current month string (YYYY-MM)

    Returns:
        str: Path to created ZIP file
    """
    # Generate ZIP filename
    timestamp = get_timestamp_from_month(current_month)
    zip_filename = ZIP_FILE_PATTERN.format(isin=isin, timestamp=timestamp)
    zip_filepath = os.path.join(OUTPUT_DIR, zip_filename)

    # Create ZIP file
    with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Add META file
        zipf.write(meta_file, os.path.basename(meta_file))

        # Add DATA file
        zipf.write(data_file, os.path.basename(data_file))

    log_debug(f"Created ZIP package: {zip_filepath}")

    # Optional: Delete the individual files after zipping
    # Uncomment the following lines if you want to keep only ZIP files
    # os.remove(meta_file)
    # os.remove(data_file)
    # log_debug(f"Removed individual files (kept in ZIP)")

    return zip_filepath


# =============================================================================
# MAIN GENERATION FUNCTIONS
# =============================================================================

def generate_files_for_isin(isin, isin_data):
    """
    Generate DATA and META files for a single ISIN, then package into ZIP

    Args:
        isin (str): ISIN code
        isin_data (dict): Dictionary containing ISIN data

    Returns:
        tuple: (meta_file_path, data_file_path, zip_file_path)
    """
    description = isin_data['description']
    time_series = isin_data['time_series']
    metadata = isin_data['metadata']
    current_month = isin_data['current_month']

    log_debug(f"\nGenerating files for ISIN: {isin}")

    # Create META file (.xls)
    meta_file = create_meta_file_xls(isin, metadata, current_month)

    # Create DATA file (.xls)
    data_file = create_data_file_xls(isin, time_series, description, current_month)

    # Create ZIP package
    zip_file = create_zip_package(isin, meta_file, data_file, current_month)

    return meta_file, data_file, zip_file


def generate_all_files(parsed_data):
    """
    Generate DATA and META files for all ISINs

    Args:
        parsed_data (dict): Dictionary containing parsed data for all ISINs
                           Keys can be either ISIN or ISIN_MONTH format

    Returns:
        dict: Dictionary with file paths for each key
    """
    log_debug("\n" + "="*80)
    log_debug("STARTING FILE GENERATION (XLS FORMAT)")
    log_debug("="*80 + "\n")

    ensure_output_directory()

    file_paths = {}

    for data_key, isin_data in parsed_data.items():
        # Extract ISIN from key (handles both "ISIN" and "ISIN_MONTH" formats)
        if 'isin' in isin_data:
            # New format with separate 'isin' field
            isin = isin_data['isin']
        else:
            # Old format where key is the ISIN
            isin = data_key
        meta_file, data_file, zip_file = generate_files_for_isin(isin, isin_data)
        file_paths[data_key] = {
            'isin': isin,
            'meta_file': meta_file,
            'data_file': data_file,
            'zip_file': zip_file
        }

    log_debug("\n" + "="*80)
    log_debug("FILE GENERATION COMPLETE")
    log_debug("="*80)
    log_debug(f"Generated files for {len(file_paths)} data combinations")
    log_debug(f"Total files: {len(file_paths) * 3} (META + DATA + ZIP per combination)")

    return file_paths


# =============================================================================
# VERIFICATION FUNCTION
# =============================================================================

def verify_generated_files(file_paths):
    """Verify that the generated files exist and have proper naming"""
    log_debug("\n" + "="*80)
    log_debug("VERIFYING GENERATED FILES")
    log_debug("="*80 + "\n")

    for isin, paths in file_paths.items():
        log_debug(f"\nVerifying files for ISIN: {isin}")

        # Check META file
        if os.path.exists(paths['meta_file']):
            size = os.path.getsize(paths['meta_file'])
            log_debug(f"  META file: OK ({size} bytes)")
        else:
            log_debug(f"  META file: MISSING", "ERROR")

        # Check DATA file
        if os.path.exists(paths['data_file']):
            size = os.path.getsize(paths['data_file'])
            log_debug(f"  DATA file: OK ({size} bytes)")
        else:
            log_debug(f"  DATA file: MISSING", "ERROR")

        # Check ZIP file
        if os.path.exists(paths['zip_file']):
            size = os.path.getsize(paths['zip_file'])
            log_debug(f"  ZIP file: OK ({size} bytes)")

            # Verify ZIP contents
            try:
                with zipfile.ZipFile(paths['zip_file'], 'r') as zipf:
                    files_in_zip = zipf.namelist()
                    log_debug(f"    Contains: {', '.join(files_in_zip)}")
            except Exception as e:
                log_debug(f"    Error reading ZIP: {e}", "ERROR")
        else:
            log_debug(f"  ZIP file: MISSING", "ERROR")

    log_debug("\nFile verification complete")


# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Main execution function"""
    print("\n" + "="*80)
    print("ITBI FILE GENERATOR (XLS FORMAT)")
    print("Banca d'Italia - Government Bond Auction Results")
    print("="*80 + "\n")

    # Note: This requires parsed data from the parser
    print("This module is designed to be imported and used by the main script.")
    print("It cannot run standalone without parsed data.")
    print("\nTo test, run: python main.py")


if __name__ == "__main__":
    main()
