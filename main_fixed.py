"""
ITBI Data Collection - Main Automation Script (FIXED VERSION)
Combines all steps: download, parse, and generate files with proper .xls format

FEATURES:
- Uses xlwt for .xls format files
- Proper file naming: ITBI_{ISIN}_TYPE_{YYYYMMDD}.xls
- Creates ZIP packages per ISIN
- Supports month selection via config
- Can process all months in the year
"""

import sys
import os
from datetime import datetime
from pathlib import Path

from downloader import download_latest_auction_file
from parser import parse_auction_data
from file_generator_xls import generate_all_files, verify_generated_files
from config import (
    OUTPUT_DIR, DOWNLOAD_DIR, EXTRACTED_DIR,
    TARGET_MONTH, PROCESS_ALL_MONTHS
)


def log_debug(message: str, prefix: str = "INFO"):
    """Log debug messages with timestamp"""
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    print(f"[{timestamp}] [{prefix}] {message}")


def print_header(title: str):
    """Print a formatted header"""
    print("\n" + "="*80)
    print(f"{title:^80}")
    print("="*80)


def print_step(step_num: int, title: str):
    """Print a formatted step header"""
    print(f"\n{'-'*80}")
    print(f"STEP {step_num}: {title}")
    print(f"{'-'*80}")


def check_directories():
    """Check if required directories exist and show their status"""
    print_step(0, "Checking Directories")

    dirs_to_check = [
        ("Downloads", DOWNLOAD_DIR),
        ("Extracted", EXTRACTED_DIR),
        ("Output", OUTPUT_DIR)
    ]

    for name, path in dirs_to_check:
        if os.path.exists(path):
            file_count = len([f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))])
            log_debug(f"{name} directory: {path} ({file_count} files)")
        else:
            log_debug(f"{name} directory: {path} (will be created)")


def show_configuration():
    """Show current configuration"""
    print_step(0.5, "Configuration")

    if PROCESS_ALL_MONTHS:
        log_debug("Processing Mode: ALL MONTHS (entire year)")
    elif TARGET_MONTH:
        year, month = TARGET_MONTH
        log_debug(f"Processing Mode: SPECIFIC MONTH ({year}-{month:02d})")
    else:
        log_debug("Processing Mode: AUTO (most recent month)")


def show_summary(parsed_data, file_paths, elapsed_time):
    """Show a summary of the processing results"""
    print("\n" + "="*80)
    print("AUTOMATION COMPLETE")
    print("="*80)

    # Count total files
    total_isins = len(file_paths)
    total_files = total_isins * 3  # META + DATA + ZIP per ISIN

    print(f"[OK] Successfully processed {total_isins} unique ISINs")
    print(f"[OK] Generated {total_files} files (1 META + 1 DATA + 1 ZIP per ISIN)")
    print(f"[OK] Total processing time: {elapsed_time:.2f} seconds")

    # Show output directory
    abs_output = os.path.abspath(OUTPUT_DIR)
    print(f"\nOutput directory: {abs_output}")

    # List generated ZIP files
    print(f"\nZIP files generated:")
    for data_key, paths in file_paths.items():
        zip_file = os.path.basename(paths['zip_file'])
        isin = paths.get('isin', data_key.split('_')[0] if '_' in data_key else data_key)
        print(f"  - {zip_file}")

    # Show ISINs processed
    print(f"\nData combinations processed:")
    for data_key in sorted(parsed_data.keys()):
        data = parsed_data[data_key]
        description = data['description']
        month = data['current_month']
        isin = data.get('isin', data_key.split('_')[0] if '_' in data_key else data_key)
        print(f"  - {isin} ({month}): {description}")

    print(f"\n[SUCCESS] ITBI data collection completed successfully!")

    # Next steps
    print(f"\nNext steps:")
    print(f"1. Review the generated ZIP files in the output directory")
    print(f"2. Upload the ZIP files to your destination system")
    if PROCESS_ALL_MONTHS:
        print(f"3. To process single month only, set PROCESS_ALL_MONTHS = False in config.py")
    else:
        print(f"3. To process all months, set PROCESS_ALL_MONTHS = True in config.py")


def main():
    """Main execution function"""
    start_time =datetime.now()

    print_header("ITBI DATA COLLECTION AUTOMATION")
    print("Banca d'Italia - Government Bond Auction Results")

    # Check directories
    check_directories()

    # Show configuration
    show_configuration()

    try:
        # STEP 1: Download latest auction file
        print_step(1, "Downloading Latest Auction File")
        log_debug("Navigating to Banca d'Italia website and downloading the latest auction results...")

        xls_file = download_latest_auction_file()

        if not xls_file:
            log_debug("Failed to download auction file", "ERROR")
            return None

        log_debug(f"Successfully downloaded and extracted: {os.path.basename(xls_file)}")

        # STEP 2: Parse auction data
        print_step(2, "Parsing Auction Data")
        log_debug("Extracting auction data for the configured month(s)...")

        parsed_data = parse_auction_data(
            xls_file,
            target_month=TARGET_MONTH,
            process_all=PROCESS_ALL_MONTHS
        )

        if not parsed_data:
            log_debug("No data was parsed", "ERROR")
            return None

        log_debug(f"Successfully parsed data for {len(parsed_data)} unique ISINs")

        # STEP 3: Generate DATA and META files
        print_step(3, "Generating DATA and META Files")
        log_debug("Creating Excel files for each ISIN...")

        file_paths = generate_all_files(parsed_data)

        if not file_paths:
            log_debug("File generation failed", "ERROR")
            return None

        log_debug(f"Successfully generated {len(file_paths) * 3} files")

        # STEP 4: Verify generated files
        print_step(4, "Verifying Generated Files")
        log_debug("Checking file structure and content...")

        verify_generated_files(file_paths)

        # Calculate elapsed time
        elapsed_time = (datetime.now() - start_time).total_seconds()

        # Show summary
        show_summary(parsed_data, file_paths, elapsed_time)

        return {
            'parsed_data': parsed_data,
            'file_paths': file_paths,
            'xls_file': xls_file,
            'elapsed_time': elapsed_time
        }

    except KeyboardInterrupt:
        print("\n\n[WARNING] Process interrupted by user")
        return None
    except Exception as e:
        print(f"\n\n[ERROR] CRITICAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    result = main()

    if result:
        print("\n" + "="*80)
        print("[SUCCESS] All operations completed")
        print("="*80)
        sys.exit(0)
    else:
        print("\n" + "="*80)
        print("[FAILED] Please check the logs above")
        print("="*80)
        sys.exit(1)
