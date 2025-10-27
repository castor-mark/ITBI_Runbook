# config.py
# ITBI (Italian Treasury Bond Index) Data Collection Configuration

import os
from datetime import datetime

# =============================================================================
# DATA SOURCE CONFIGURATION
# =============================================================================

BASE_URL = 'https://www.bancaditalia.it/compiti/operazioni-mef/risultati-aste/index.html'
PROVIDER_NAME = 'Banca d\'Italia'
DATASET_NAME = 'ITBI'
COUNTRY = 'ITA'

# =============================================================================
# TIMESTAMPED FOLDERS CONFIGURATION
# =============================================================================

# Generate timestamp for this run (format: YYYYMMDD_HHMMSS)
RUN_TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')

# Use timestamped folders to avoid conflicts between runs
USE_TIMESTAMPED_FOLDERS = True

# =============================================================================
# PROCESSING CONFIGURATION
# =============================================================================

# Target month for data extraction
# Set to None to auto-detect most recent month from data
# Format: (year, month) e.g., (2025, 10) for October 2025
# Or 'all' to process all months in the year
TARGET_MONTH = None  # Options: None (auto), (2025, 10), 'all'

# When set to True, process all months from January to the most recent month
PROCESS_ALL_MONTHS = False  # Set to True to process entire year

# =============================================================================
# WEB SCRAPING SELECTORS
# =============================================================================

SELECTORS = {
    # Cookie banner
    'cookie_accept_button': 'button#cb-AcceptAll',
    'cookie_banner': 'div#banner-privacy',

    # Auction results section
    'auction_results_section': 'div.accordion-date',
    'auction_results_heading': 'h2.sr-only.bdi-no-anchor',  # Contains "Auction results"
    'auction_links': 'a.accordion-link-download',
    'auction_link_text': 'font',  # Contains download link text

    # Alternative selectors
    'download_links_container': 'ul.accordion-list-link',
    'pdf_link': 'a.accordion-link-pdf',
}

# =============================================================================
# FILE PATTERNS
# =============================================================================

# Expected file name patterns
CURRENT_YEAR_PATTERN = r'Current year auctions as of (.+)'
AUCTION_FILE_PATTERN = r'aste_corrente.*\.zip'
DOWNLOAD_DIR = './downloads'
EXTRACTED_DIR = './extracted'

# =============================================================================
# REQUIRED COLUMNS FROM SOURCE XLS
# =============================================================================

# Italian column names in source file
SOURCE_COLUMNS_IT = {
    'auction_date': 'data asta',
    'isin': 'ISIN',
    'offered': 'offerto',
    'min_offered': 'minimo offerto',
    'max_offered': 'massimo offerto',
    'required': 'richiesto',
    'assigned': 'assegnato'
}

# English column names (after translation or mapping)
SOURCE_COLUMNS_EN = {
    'auction_date': 'Auction date',
    'isin': 'ISIN',
    'offered': 'Offered',
    'min_offered': 'Minimum offered',
    'max_offered': 'Maximum offered',
    'required': 'Required',
    'assigned': 'Assigned'
}

# =============================================================================
# TIME SERIES MAPPING
# =============================================================================

# For each ISIN, we create 5 time series
TIME_SERIES_MAPPING = {
    'ASGN': {
        'suffix': 'ASGN.ITBI.D',
        'description': 'amounts:assigned',
        'source_column': 'assegnato'
    },
    'MAX': {
        'suffix': 'MAX.ITBI.D',
        'description': 'amounts:maximum offered',
        'source_column': 'massimo offerto'
    },
    'MIN': {
        'suffix': 'MIN.ITBI.D',
        'description': 'amounts:minimum offered',
        'source_column': 'minimo offerto'
    },
    'OFR': {
        'suffix': 'OFR.ITBI.D',
        'description': 'amounts:offered',
        'source_column': 'offerto'
    },
    'REQ': {
        'suffix': 'REQ.ITBI.D',
        'description': 'amounts:required',
        'source_column': 'richiesto'
    }
}

# Order for columns in output files (CORRECTED: OFR, MIN, MAX, REQ, ASGN)
TIME_SERIES_ORDER = ['OFR', 'MIN', 'MAX', 'REQ', 'ASGN']

# =============================================================================
# METADATA STANDARD FIELDS
# =============================================================================

METADATA_DEFAULTS = {
    'FREQUENCY': 'D',  # Daily
    'MULTIPLIER': 6.0,
    'AGGREGATION_TYPE': 'END_OF_PERIOD',
    'UNIT_TYPE': 'LEVEL',
    'DATA_TYPE': 'CURRENCY',
    'DATA_UNIT': 'EUR',
    'SEASONALLY_ADJUSTED': 'NSA',
    'ANNUALIZED': False,
    'PROVIDER_MEASURE_URL': BASE_URL,
    'PROVIDER': 'AfricaAI',
    'SOURCE': 'BdIt',
    'SOURCE_DESCRIPTION': PROVIDER_NAME,
    'COUNTRY': COUNTRY,
    'DATASET': DATASET_NAME
}

# Metadata file columns (removed NEXT_RELEASE_DATE and LAST_RELEASE_DATE)
METADATA_COLUMNS = [
    'CODE',
    'DESCRIPTION',
    'FREQUENCY',
    'MULTIPLIER',
    'AGGREGATION_TYPE',
    'UNIT_TYPE',
    'DATA_TYPE',
    'DATA_UNIT',
    'SEASONALLY_ADJUSTED',
    'ANNUALIZED',
    'PROVIDER_MEASURE_URL',
    'PROVIDER',
    'SOURCE',
    'SOURCE_DESCRIPTION',
    'COUNTRY',
    'DATASET'
]

# =============================================================================
# DATE FORMATS
# =============================================================================

DATE_FORMAT_INPUT = '%d/%m/%Y'  # Italian date format from XLS
DATE_FORMAT_OUTPUT = '%Y-%m-%d'  # ISO format for DATA files
DATETIME_FORMAT_META = '%Y-%m-%d %H:%M:%S'  # Format for NEXT_RELEASE_DATE

# =============================================================================
# BROWSER CONFIGURATION
# =============================================================================

HEADLESS_MODE = False
DEBUG_MODE = True
WAIT_TIMEOUT = 15
PAGE_LOAD_DELAY = 3

# =============================================================================
# OUTPUT CONFIGURATION
# =============================================================================

# Base directories
BASE_DOWNLOAD_DIR = './downloads'
BASE_EXTRACTED_DIR = './extracted'
BASE_OUTPUT_DIR = './output'

# Apply timestamping if enabled
if USE_TIMESTAMPED_FOLDERS:
    DOWNLOAD_DIR = os.path.join(BASE_DOWNLOAD_DIR, RUN_TIMESTAMP)
    EXTRACTED_DIR = os.path.join(BASE_EXTRACTED_DIR, RUN_TIMESTAMP)
    OUTPUT_DIR = os.path.join(BASE_OUTPUT_DIR, RUN_TIMESTAMP)
else:
    DOWNLOAD_DIR = BASE_DOWNLOAD_DIR
    EXTRACTED_DIR = BASE_EXTRACTED_DIR
    OUTPUT_DIR = BASE_OUTPUT_DIR

# File naming patterns (ITBI_ISIN_TYPE_YYYYMMDD.xls)
DATA_FILE_PATTERN = 'ITBI_{isin}_DATA_{timestamp}.xls'
META_FILE_PATTERN = 'ITBI_{isin}_META_{timestamp}.xls'
ZIP_FILE_PATTERN = 'ITBI_{isin}_{timestamp}.zip'

# Date format for filenames (YYYYMMDD)
FILENAME_DATE_FORMAT = '%Y%m%d'