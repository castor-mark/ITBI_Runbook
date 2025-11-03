"""
ITBI Data Collection - Parser
Parses the downloaded XLS file and extracts auction data for processing
"""

import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import re

from config import (
    SOURCE_COLUMNS_IT, SOURCE_COLUMNS_EN, TIME_SERIES_MAPPING,
    TIME_SERIES_ORDER, DATE_FORMAT_INPUT, DATE_FORMAT_OUTPUT,
    METADATA_DEFAULTS, METADATA_COLUMNS, DATETIME_FORMAT_META,
    TARGET_MONTH, PROCESS_ALL_MONTHS
)


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def log_debug(message: str, prefix: str = "INFO"):
    """Log debug messages with timestamp"""
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    print(f"[{timestamp}] [{prefix}] {message}")


def clean_isin_column(df):
    """Clean the ISIN column to filter valid ISIN codes"""
    # Check if 'ISIN' column exists, if not, try to find it
    if 'ISIN' not in df.columns:
        # Print available columns for debugging
        log_debug(f"Available columns: {list(df.columns)}")
        
        # Try to find the ISIN column by looking for common patterns
        isin_col = None
        for col in df.columns:
            if 'isin' in str(col).lower():
                isin_col = col
                log_debug(f"Found ISIN column: {isin_col}")
                break
        
        # If still not found, try to use the third column (index 2) which should be ISIN
        if isin_col is None and len(df.columns) > 2:
            isin_col = df.columns[2]
            log_debug(f"Using third column as ISIN: {isin_col}")
        
        if isin_col is None:
            log_debug("Could not find ISIN column in the Excel file", "ERROR")
            return df.copy()  # Return original df if ISIN column not found
        
        # Rename the column to 'ISIN' for consistency
        df = df.rename(columns={isin_col: 'ISIN'})
    
    # Debug: Show some sample values from the ISIN column
    log_debug(f"Sample ISIN values: {df['ISIN'].head(10).tolist()}")
    
    # Filter out rows where ISIN is not a valid ISIN code (starts with IT and has 12 characters)
    # First, convert to string to handle any non-string values
    df['ISIN'] = df['ISIN'].astype(str)
    
    # Check for valid ISIN patterns
    valid_isin_mask = df['ISIN'].str.match(r'^IT\d{11}$', na=False)
    valid_count = valid_isin_mask.sum()
    log_debug(f"Found {valid_count} valid ISIN rows")
    
    # If no valid ISINs found, try a more relaxed pattern
    if valid_count == 0:
        log_debug("No valid ISINs found with strict pattern, trying relaxed pattern")
        # Try to find any values that look like ISINs (contain 'IT' and numbers)
        relaxed_mask = df['ISIN'].str.contains(r'IT\d+', na=False)
        relaxed_count = relaxed_mask.sum()
        log_debug(f"Found {relaxed_count} rows with relaxed ISIN pattern")
        
        if relaxed_count > 0:
            valid_isin_mask = relaxed_mask
    
    # If still no valid ISINs, check if there are any non-empty values
    if valid_isin_mask.sum() == 0:
        log_debug("No ISINs found with relaxed pattern, checking for any non-empty values")
        non_empty_mask = df['ISIN'].notna() & (df['ISIN'] != '') & (df['ISIN'] != 'nan')
        non_empty_count = non_empty_mask.sum()
        log_debug(f"Found {non_empty_count} non-empty ISIN rows")
        
        if non_empty_count > 0:
            valid_isin_mask = non_empty_mask
    
    return df[valid_isin_mask].copy()


def convert_date_format(date_series):
    """Convert date series to ISO format (YYYY-MM-DD)"""
    return date_series.dt.strftime(DATE_FORMAT_OUTPUT)


def get_month_data(df, target_month=None, process_all=False):
    """
    Extract data for specified month(s)

    Args:
        df: DataFrame with auction data
        target_month: Tuple of (year, month) or None for most recent
        process_all: If True, return data grouped by all months

    Returns:
        If process_all=True: dict of {month_str: month_data}
        If process_all=False: (month_data, month_str)
    """
    # Check if 'data asta' column exists, if not, try to find it
    if 'data asta' not in df.columns:
        # Try to find the date column by looking for common patterns
        date_col = None
        for col in df.columns:
            if 'data' in str(col).lower() or 'date' in str(col).lower():
                date_col = col
                log_debug(f"Found date column: {date_col}")
                break

        # If still not found, try to use the first column (index 0) which should be date
        if date_col is None and len(df.columns) > 0:
            date_col = df.columns[0]
            log_debug(f"Using first column as date: {date_col}")

        if date_col is None:
            log_debug("Could not find date column in the Excel file", "ERROR")
            if process_all:
                return {}
            return df.copy(), ""

        # Rename the column to 'data asta' for consistency
        df = df.rename(columns={date_col: 'data asta'})

    # Debug: Show some sample values from the date column
    log_debug(f"Sample date values: {df['data asta'].head(10).tolist()}")

    # Convert date column to datetime if it's not already
    if not pd.api.types.is_datetime64_any_dtype(df['data asta']):
        df['data asta'] = pd.to_datetime(df['data asta'], dayfirst=True, errors='coerce')

    # Remove rows with invalid dates
    valid_date_mask = df['data asta'].notna()
    df = df[valid_date_mask].copy()
    log_debug(f"Rows with valid dates: {len(df)}")

    if len(df) == 0:
        log_debug("No valid dates found", "ERROR")
        if process_all:
            return {}
        return df.copy(), ""

    # Add month column for filtering
    df['month'] = df['data asta'].dt.to_period('M')

    # Process all months
    if process_all:
        log_debug("Processing ALL months in the dataset")
        all_months = sorted(df['month'].unique())
        log_debug(f"Found {len(all_months)} months: {[str(m) for m in all_months]}")

        result = {}
        for month in all_months:
            month_data = df[df['month'] == month].copy()
            month_str = str(month)
            result[month_str] = month_data
            log_debug(f"  {month_str}: {len(month_data)} rows")

        return result

    # Process specific month
    if target_month:
        year, month = target_month
        target_period = pd.Period(f"{year}-{month:02d}", freq='M')
        log_debug(f"Processing SPECIFIC month: {target_period}")

        if target_period in df['month'].values:
            month_data = df[df['month'] == target_period].copy()
            log_debug(f"Rows for {target_period}: {len(month_data)}")
            return month_data, str(target_period)
        else:
            log_debug(f"No data found for {target_period}", "WARNING")
            return df.iloc[:0].copy(), str(target_period)

    # Default: get most recent month
    current_month = df['month'].max()
    log_debug(f"Processing MOST RECENT month: {current_month}")

    current_month_data = df[df['month'] == current_month].copy()
    log_debug(f"Rows for current month: {len(current_month_data)}")

    return current_month_data, str(current_month)


def get_unique_isins(df):
    """Get unique ISINs from the dataset"""
    if len(df) == 0:
        return []
    
    unique_isins = df['ISIN'].unique()
    log_debug(f"Unique ISINs in current month: {len(unique_isins)}")
    return unique_isins


def prepare_isin_data(df, isin):
    """Prepare data for a specific ISIN"""
    # Filter data for the specific ISIN
    isin_data = df[df['ISIN'] == isin].copy()
    
    # Sort by date (ascending order as required)
    isin_data = isin_data.sort_values('data asta')
    
    # Convert date to string format
    isin_data['date_str'] = convert_date_format(isin_data['data asta'])
    
    # Get the description from the appropriate column
    description_col = None
    for col in isin_data.columns:
        if 'descrizione' in str(col).lower() or 'description' in str(col).lower():
            description_col = col
            log_debug(f"Found description column: {description_col}")
            break
    
    # If still not found, try to use the sixth column (index 5) which should be description
    if description_col is None and len(isin_data.columns) > 5:
        description_col = isin_data.columns[5]
        log_debug(f"Using sixth column as description: {description_col}")
    
    description = isin_data.iloc[0][description_col] if description_col and description_col in isin_data.columns else ""
    
    log_debug(f"Prepared data for ISIN {isin}: {len(isin_data)} rows")
    
    return isin_data, description


def create_time_series_data(isin_data, description):
    """Create time series data for all 5 metrics"""
    # Initialize result dictionary
    time_series = {}
    
    # Based on the Excel structure, we need to map the columns correctly
    # From the error log, we can see the columns are:
    # ['data asta', 'data regolamento', 'ISIN', 'numero tranche', 
    #  'ordinaria (O) / supplementare (S)', 'descrizione', 'data scadenza', 
    #  'coefficiente di indicizzazione', 'tipo titolo', 'Unnamed: 9', 
    #  'Unnamed: 10', 'importi (mln euro)', 'Unnamed: 12', 'Unnamed: 13', 
    #  'prezzo aggiudicazione', 'rendimento aggiudicazione (BOT)', 
    #  'rendimento aggiudicazione (altri titoli)', 'numero  operatori partecipanti']
    
    # Create a mapping from our expected column names to the actual column indices
    # CORRECT mapping verified from source XLS data:
    # Col 9 (Unnamed: 9) = offerto (offered)
    # Col 10 (Unnamed: 10) = minimo offerto (minimum offered)
    # Col 11 (importi mln euro) = massimo offerto (maximum offered)
    # Col 12 (Unnamed: 12) = richiesto (required)
    # Col 13 (Unnamed: 13) = assegnato (assigned)
    column_mapping = {
        'offerto': 9,           # 'Unnamed: 9' - offered
        'minimo offerto': 10,   # 'Unnamed: 10' - minimum offered
        'massimo offerto': 11,  # 'importi (mln euro)' - maximum offered
        'richiesto': 12,        # 'Unnamed: 12' - required
        'assegnato': 13         # 'Unnamed: 13' - assigned
    }
    
    # Create data for each time series type
    for ts_type in TIME_SERIES_ORDER:
        mapping = TIME_SERIES_MAPPING[ts_type]
        source_column = mapping['source_column']
        
        # Get the column index from our mapping
        if source_column in column_mapping:
            col_index = column_mapping[source_column]
            if col_index < len(isin_data.columns):
                actual_column = isin_data.columns[col_index]
                log_debug(f"Using column {actual_column} (index {col_index}) for {source_column}")
                
                # Create a DataFrame with date and value
                ts_df = pd.DataFrame({
                    'date': isin_data['date_str'],
                    'value': isin_data[actual_column]
                })
                
                time_series[ts_type] = ts_df
            else:
                log_debug(f"Column index {col_index} out of range for {source_column}", "WARNING")
                # Create an empty DataFrame with the required structure
                ts_df = pd.DataFrame({
                    'date': isin_data['date_str'],
                    'value': [np.nan] * len(isin_data)
                })
                time_series[ts_type] = ts_df
        else:
            log_debug(f"No mapping found for {source_column}", "WARNING")
            # Create an empty DataFrame with the required structure
            ts_df = pd.DataFrame({
                'date': isin_data['date_str'],
                'value': [np.nan] * len(isin_data)
            })
            time_series[ts_type] = ts_df
    
    return time_series


def create_metadata_rows(isin, description, current_month):
    """Create metadata rows for all 5 time series of an ISIN"""
    metadata_rows = []

    # Create a row for each time series type
    for ts_type in TIME_SERIES_ORDER:
        mapping = TIME_SERIES_MAPPING[ts_type]

        # Create the code
        code = f"{isin}.{mapping['suffix']}"

        # Create the description (no spaces after colons/semicolons)
        desc = f"ISIN:{isin};{description}:{mapping['description']}"

        # Create metadata row
        metadata_row = METADATA_DEFAULTS.copy()
        metadata_row['CODE'] = code
        metadata_row['DESCRIPTION'] = desc

        # Ensure all required columns are present
        for col in METADATA_COLUMNS:
            if col not in metadata_row:
                metadata_row[col] = ""

        metadata_rows.append(metadata_row)

    return metadata_rows


# =============================================================================
# MAIN PARSER FUNCTION
# =============================================================================

def parse_auction_data(xls_file_path, target_month=None, process_all=False):
    """
    Parse the auction data XLS file and extract data for processing

    Args:
        xls_file_path (str): Path to the XLS file
        target_month (tuple): Optional (year, month) tuple for specific month
        process_all (bool): If True, process all months in the file

    Returns:
        dict: Dictionary containing parsed data for each ISIN (optionally grouped by month)
    """
    log_debug("\n" + "="*80)
    log_debug("STARTING AUCTION DATA PARSING")
    log_debug("="*80 + "\n")

    try:
        # Read the XLS file with proper header row (row 2, index 1)
        log_debug(f"Reading XLS file: {xls_file_path}")
        df = pd.read_excel(xls_file_path, header=1)
        log_debug(f"Original rows: {len(df)}")
        log_debug(f"Columns: {list(df.columns)}")

        # Skip the first 2 rows (empty header rows)
        df = df.iloc[2:].reset_index(drop=True)
        log_debug(f"Rows after skipping header: {len(df)}")

        # Remove completely empty rows
        df = df.dropna(how='all')
        log_debug(f"Rows after removing empty rows: {len(df)}")

        # Clean the ISIN column
        df_clean = clean_isin_column(df)
        log_debug(f"Valid ISIN rows: {len(df_clean)}")

        if len(df_clean) == 0:
            log_debug("No valid ISIN rows found", "ERROR")
            return None

        # Get data for specified month(s)
        month_data_result = get_month_data(df_clean, target_month, process_all)

        if process_all:
            # Processing all months - month_data_result is a dict
            if not month_data_result:
                log_debug("No month data found", "ERROR")
                return None

            # Process each month separately
            all_parsed_data = {}

            for month_str, month_df in month_data_result.items():
                log_debug(f"\n" + "-"*80)
                log_debug(f"PROCESSING MONTH: {month_str}")
                log_debug("-"*80)

                # Get unique ISINs for this month
                unique_isins = get_unique_isins(month_df)

                if len(unique_isins) == 0:
                    log_debug(f"No unique ISINs found for {month_str}", "WARNING")
                    continue

                log_debug(f"Found {len(unique_isins)} unique ISINs for {month_str}")

                # Process each ISIN
                for isin in unique_isins:
                    # Create unique key: ISIN_MONTH
                    data_key = f"{isin}_{month_str}"

                    log_debug(f"  Processing ISIN: {isin}")

                    # Prepare data for this ISIN
                    isin_data, description = prepare_isin_data(month_df, isin)

                    # Create time series data
                    time_series = create_time_series_data(isin_data, description)

                    # Create metadata rows
                    metadata_rows = create_metadata_rows(isin, description, month_str)

                    # Store in result dictionary with unique key
                    all_parsed_data[data_key] = {
                        'isin': isin,
                        'description': description,
                        'time_series': time_series,
                        'metadata': metadata_rows,
                        'current_month': month_str
                    }

            log_debug("\n" + "="*80)
            log_debug("PARSING COMPLETE (ALL MONTHS)")
            log_debug("="*80)
            log_debug(f"Processed {len(all_parsed_data)} ISIN-month combinations")

            return all_parsed_data

        else:
            # Processing single month - month_data_result is a tuple
            current_month_data, current_month_str = month_data_result

            if len(current_month_data) == 0:
                log_debug("No data for specified month found", "ERROR")
                return None

            # Get unique ISINs
            unique_isins = get_unique_isins(current_month_data)

            if len(unique_isins) == 0:
                log_debug("No unique ISINs found", "ERROR")
                return None

            # Prepare data for each ISIN
            parsed_data = {}

            for isin in unique_isins:
                log_debug(f"\nProcessing ISIN: {isin}")

                # Prepare data for this ISIN
                isin_data, description = prepare_isin_data(current_month_data, isin)

                # Create time series data
                time_series = create_time_series_data(isin_data, description)

                # Create metadata rows
                metadata_rows = create_metadata_rows(isin, description, current_month_str)

                # Store in result dictionary
                parsed_data[isin] = {
                    'description': description,
                    'time_series': time_series,
                    'metadata': metadata_rows,
                    'current_month': current_month_str
                }

            log_debug("\n" + "="*80)
            log_debug("PARSING COMPLETE")
            log_debug("="*80)
            log_debug(f"Processed {len(parsed_data)} unique ISINs")

            return parsed_data

    except Exception as e:
        log_debug(f"Error parsing auction data: {str(e)}", "ERROR")
        import traceback
        traceback.print_exc()
        return None


# =============================================================================
# TEST FUNCTION
# =============================================================================

def test_parser():
    """Test the parser with the downloaded file"""
    from pathlib import Path
    
    # Find the extracted XLS file
    extracted_dir = Path('./extracted')
    xls_files = list(extracted_dir.glob('*.xls'))
    
    if not xls_files:
        log_debug("No XLS files found in extracted directory", "ERROR")
        return None
    
    xls_file = xls_files[0]
    log_debug(f"Using XLS file: {xls_file}")
    
    # Parse the data
    parsed_data = parse_auction_data(str(xls_file))
    
    if parsed_data:
        # Print summary
        log_debug("\nPARSED DATA SUMMARY:")
        for isin, data in parsed_data.items():
            log_debug(f"\nISIN: {isin}")
            log_debug(f"  Description: {data['description']}")
            log_debug(f"  Current Month: {data['current_month']}")
            log_debug(f"  Time Series: {list(data['time_series'].keys())}")
            
            # Show sample data for each time series
            for ts_type, ts_df in data['time_series'].items():
                log_debug(f"    {ts_type}: {len(ts_df)} data points")
                if len(ts_df) > 0:
                    log_debug(f"      First date: {ts_df.iloc[0]['date']}, value: {ts_df.iloc[0]['value']}")
            
            # Show metadata
            log_debug(f"  Metadata: {len(data['metadata'])} rows")
            for i, row in enumerate(data['metadata']):
                log_debug(f"    Row {i+1}: {row['CODE']}")
    
    return parsed_data


# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Main execution function"""
    print("\n" + "="*80)
    print("ITBI AUCTION DATA PARSER")
    print("Banca d'Italia - Government Bond Auction Results")
    print("="*80 + "\n")
    
    result = test_parser()
    
    if result:
        print("\n[SUCCESS] Parsing completed!")
        print(f"Processed {len(result)} unique ISINs")
        return result
    else:
        print("\n[FAILED] Parsing failed!")
        return None


if __name__ == "__main__":
    result = main()
    if result:
        print("\nNext step: Generate DATA and META files for each ISIN")