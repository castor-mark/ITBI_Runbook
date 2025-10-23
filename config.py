# config.py
# ITBI (Italian Treasury Bond Index) Data Collection Configuration

# =============================================================================
# DATA SOURCE CONFIGURATION
# =============================================================================

BASE_URL = 'https://www.bancaditalia.it/compiti/operazioni-mef/risultati-aste/index.html'
PROVIDER_NAME = 'Banca d\'Italia'
DATASET_NAME = 'ITBI'
COUNTRY = 'ITA'

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
        'suffix': 'ASGN.ITBI.M',
        'description': 'amounts: assigned',
        'source_column': 'assigned'
    },
    'MAX': {
        'suffix': 'MAX.ITBI.M',
        'description': 'amounts: maximum offered',
        'source_column': 'max_offered'
    },
    'MIN': {
        'suffix': 'MIN.ITBI.M',
        'description': 'amounts: minimum offered',
        'source_column': 'min_offered'
    },
    'OFR': {
        'suffix': 'OFR.ITBI.M',
        'description': 'amounts: offered',
        'source_column': 'offered'
    },
    'REQ': {
        'suffix': 'REQ.ITBI.M',
        'description': 'amounts: required',
        'source_column': 'required'
    }
}

# Order for columns in output files
TIME_SERIES_ORDER = ['ASGN', 'MAX', 'MIN', 'OFR', 'REQ']

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
    'ANNUALIZED': '',
    'PROVIDER_MEASURE_URL': BASE_URL,
    'PROVIDER': 'AfricaAI',
    'SOURCE': 'BdIt',
    'SOURCE_DESCRIPTION': PROVIDER_NAME,
    'COUNTRY': COUNTRY,
    'DATASET': DATASET_NAME
}

# Metadata file columns
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
    'DATASET',
    'NEXT_RELEASE_DATE',
    'LAST_RELEASE_DATE'
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

OUTPUT_DIR = './output'
DATA_FILE_PATTERN = '{isin}_DATA_{timestamp}.xls'
META_FILE_PATTERN = '{isin}_META_{timestamp}.xls'
