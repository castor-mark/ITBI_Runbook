"""
ITBI Data Collection - Main Automation Script
Combines all steps: download, parse, and generate files
"""

import sys
import os
from datetime import datetime
from pathlib import Path

from downloader import download_latest_auction_file
from parser import parse_auction_data
from file_generator import generate_all_files, verify_generated_files
from config import OUTPUT_DIR, DOWNLOAD_DIR, EXTRACTED_DIR


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
    print(f"\n{'─'*80}")
    print(f"STEP {step_num}: {title}")
    print(f"{'─'*80}")


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


def show_summary(parsed_data, file_paths, elapsed_time):
    """Show a summary of the processing results"""
    print("\n" + "="*80)
    print("AUTOMATION COMPLETE")
    print("="*80)
    
    print(f"✅ Successfully processed {len(parsed_data)} unique ISINs")
    print(f"✅ Generated {len(file_paths) * 2} files (1 META + 1 DATA per ISIN)")
    print(f"✅ Total processing time: {elapsed_time.total_seconds():.2f} seconds")
    
    print(f"\n📁 Output directory: {os.path.abspath(OUTPUT_DIR)}")
    print(f"📄 Files generated:")
    
    for isin, paths in file_paths.items():
        meta_file = os.path.basename(paths['meta_file'])
        data_file = os.path.basename(paths['data_file'])
        print(f"  • {isin}: {meta_file}, {data_file}")
    
    print(f"\n📊 ISINs processed:")
    for isin, data in parsed_data.items():
        description = data['description'][:50] + "..." if len(data['description']) > 50 else data['description']
        print(f"  • {isin}: {description}")


def main():
    """Main execution function"""
    print_header("ITBI DATA COLLECTION AUTOMATION")
    print("Banca d'Italia - Government Bond Auction Results")
    
    start_time = datetime.now()
    
    try:
        # Check directories
        check_directories()
        
        # Step 1: Download the latest auction file
        print_step(1, "Downloading Latest Auction File")
        log_debug("Navigating to Banca d'Italia website and downloading the latest auction results...")
        
        xls_file_path = download_latest_auction_file()
        
        if not xls_file_path:
            log_debug("Failed to download auction file", "ERROR")
            print("\n❌ DOWNLOAD FAILED!")
            print("Please check the error messages above for details.")
            return False
        
        log_debug(f"Successfully downloaded and extracted: {os.path.basename(xls_file_path)}")
        
        # Step 2: Parse the auction data
        print_step(2, "Parsing Auction Data")
        log_debug("Extracting auction data for the current month...")
        
        parsed_data = parse_auction_data(xls_file_path)
        
        if not parsed_data:
            log_debug("Failed to parse auction data", "ERROR")
            print("\n❌ PARSING FAILED!")
            print("Please check the error messages above for details.")
            return False
        
        log_debug(f"Successfully parsed data for {len(parsed_data)} unique ISINs")
        
        # Step 3: Generate DATA and META files
        print_step(3, "Generating DATA and META Files")
        log_debug("Creating Excel files for each ISIN...")
        
        file_paths = generate_all_files(parsed_data)
        
        if not file_paths:
            log_debug("Failed to generate files", "ERROR")
            print("\n❌ FILE GENERATION FAILED!")
            print("Please check the error messages above for details.")
            return False
        
        log_debug(f"Successfully generated {len(file_paths) * 2} files")
        
        # Step 4: Verify generated files
        print_step(4, "Verifying Generated Files")
        log_debug("Checking file structure and content...")
        
        verify_generated_files(file_paths)
        
        # Calculate elapsed time
        elapsed_time = datetime.now() - start_time
        
        # Show summary
        show_summary(parsed_data, file_paths, elapsed_time)
        
        return True
        
    except KeyboardInterrupt:
        log_debug("Process interrupted by user", "WARNING")
        print("\n\n⚠️ PROCESS INTERRUPTED!")
        print("The automation was stopped by the user.")
        return False
        
    except Exception as e:
        log_debug(f"Critical error in automation: {str(e)}", "ERROR")
        import traceback
        traceback.print_exc()
        
        print("\n" + "="*80)
        print("❌ AUTOMATION FAILED!")
        print("="*80)
        print(f"Error: {str(e)}")
        print("Please check the error messages above for details.")
        
        return False


def check_dependencies():
    """Check if all required dependencies are installed"""
    try:
        import pandas as pd
        import selenium
        import undetected_chromedriver
        return True
    except ImportError as e:
        print(f"❌ Missing dependency: {str(e)}")
        print("Please install all required dependencies with:")
        print("pip install pandas selenium undetected-chromedriver openpyxl")
        return False


if __name__ == "__main__":
    # Check dependencies first
    if not check_dependencies():
        sys.exit(1)
    
    # Run the main automation
    success = main()
    
    if success:
        print("\n🎉 ITBI data collection completed successfully!")
        print("\nNext steps:")
        print("1. Review the generated files in the output directory")
        print("2. Upload the files to your destination system")
        print("3. Schedule this script to run monthly (e.g., on the 1st of each month)")
        sys.exit(0)
    else:
        print("\n❌ ITBI data collection failed!")
        print("Please resolve the issues above and try again.")
        sys.exit(1)