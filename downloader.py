"""
ITBI Data Collection - File Downloader
Downloads the latest auction results ZIP file from Banca d'Italia
"""

import os
import time
import re
import zipfile
from datetime import datetime
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import undetected_chromedriver as uc

from config import (
    BASE_URL, SELECTORS, DOWNLOAD_DIR, EXTRACTED_DIR,
    HEADLESS_MODE, DEBUG_MODE, WAIT_TIMEOUT, PAGE_LOAD_DELAY,
    CURRENT_YEAR_PATTERN, AUCTION_FILE_PATTERN
)


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def log_debug(message: str, prefix: str = "INFO"):
    """Log debug messages with timestamp"""
    if DEBUG_MODE:
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"[{timestamp}] [{prefix}] {message}")


def ensure_directories():
    """Create necessary directories if they don't exist"""
    Path(DOWNLOAD_DIR).mkdir(parents=True, exist_ok=True)
    Path(EXTRACTED_DIR).mkdir(parents=True, exist_ok=True)
    log_debug(f"Directories ensured: {DOWNLOAD_DIR}, {EXTRACTED_DIR}")


def setup_driver():
    """Set up Chrome WebDriver with download preferences"""
    log_debug("Setting up Chrome WebDriver...")

    # Get absolute path for downloads
    download_path = os.path.abspath(DOWNLOAD_DIR)

    options = uc.ChromeOptions()

    # Download preferences
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0
    }
    options.add_experimental_option("prefs", prefs)

    if HEADLESS_MODE:
        options.add_argument("--headless=new")

    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--lang=en-US")

    try:
        driver = uc.Chrome(options=options)
        log_debug("WebDriver initialized successfully", "SUCCESS")
        return driver
    except Exception as e:
        log_debug(f"Error creating driver: {str(e)}", "ERROR")
        raise


def wait_for_element(driver, by: By, selector: str, timeout: int = WAIT_TIMEOUT):
    """Wait for element to be present"""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
        return element
    except TimeoutException:
        log_debug(f"Timeout waiting for element: {selector}", "WARNING")
        return None


def wait_for_clickable(driver, by: By, selector: str, timeout: int = WAIT_TIMEOUT):
    """Wait for element to be clickable"""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((by, selector))
        )
        return element
    except TimeoutException:
        log_debug(f"Timeout waiting for clickable element: {selector}", "WARNING")
        return None


def safe_click(driver, element, description: str = "element"):
    """Safely click an element"""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.5)
        element.click()
        log_debug(f"Clicked {description}", "SUCCESS")
        return True
    except Exception as e:
        log_debug(f"Error clicking {description}: {str(e)}", "ERROR")
        return False


# =============================================================================
# PAGE INTERACTION FUNCTIONS
# =============================================================================

def handle_cookie_banner(driver):
    """Accept cookie banner if present"""
    log_debug("Checking for cookie banner...")

    try:
        # Wait a bit for banner to appear
        time.sleep(2)

        # Try to find the accept button
        accept_button = wait_for_clickable(
            driver,
            By.CSS_SELECTOR,
            SELECTORS['cookie_accept_button'],
            timeout=5
        )

        if accept_button:
            log_debug("Cookie banner found, accepting cookies...")
            safe_click(driver, accept_button, "Cookie accept button")
            time.sleep(2)
            log_debug("Cookies accepted", "SUCCESS")
            return True
        else:
            log_debug("No cookie banner found or already accepted")
            return False
    except Exception as e:
        log_debug(f"Error handling cookie banner: {str(e)}", "WARNING")
        return False


def trigger_page_translation(driver):
    """
    Trigger Google Translate on the page
    Note: This may require manual intervention or browser extension
    """
    log_debug("Page translation step...")
    log_debug("NOTE: If page is in Italian, right-click and select 'Translate to English'", "INFO")

    # Check current language
    try:
        html_tag = driver.find_element(By.TAG_NAME, "html")
        current_lang = html_tag.get_attribute("lang")
        log_debug(f"Current page language: {current_lang}")

        if current_lang and current_lang.startswith('en'):
            log_debug("Page is already in English", "SUCCESS")
            return True
        else:
            log_debug("Page is in Italian - waiting for translation...", "WARNING")
            # Wait a bit for manual translation if needed
            time.sleep(5)
            return True
    except Exception as e:
        log_debug(f"Could not determine page language: {str(e)}", "WARNING")
        return True


def find_latest_auction_file_link(driver):
    """
    Find and return the download link for the latest auction file
    Works with both Italian and English page
    """
    log_debug("Looking for latest auction file link...")

    try:
        # Wait for page to fully load
        time.sleep(PAGE_LOAD_DELAY)

        # Print page source for debugging (first 500 chars of body)
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text[:500]
            log_debug(f"Page body preview: {body_text[:200]}...")
        except:
            pass

        # Method 1: Find all links with .zip extension
        log_debug("Method 1: Searching all links for .zip files...")
        all_links = driver.find_elements(By.TAG_NAME, "a")
        log_debug(f"Found {len(all_links)} total links on page")

        zip_links = []
        for link in all_links:
            href = link.get_attribute('href')
            if href and href.endswith('.zip'):
                link_text = link.text.strip()
                log_debug(f"  ZIP link found: {link_text[:50]} -> {href}")
                zip_links.append((link, href, link_text))

        # Filter for current year auctions (anno_corrente or aste_corrente)
        for link, href, link_text in zip_links:
            if 'anno_corrente' in href.lower() or 'aste_corrente' in href.lower() or 'corrente' in href.lower():
                log_debug(f"[OK] Found current year auction file!", "SUCCESS")
                log_debug(f"  Text: {link_text}")
                log_debug(f"  URL: {href}")
                return link, href, link_text

        # If we found any zip links, return the first one
        if zip_links:
            link, href, link_text = zip_links[0]
            log_debug(f"Using first ZIP file found: {link_text}", "WARNING")
            return link, href, link_text

        # Method 2: Look by CSS selectors
        log_debug("Method 2: Searching by CSS selectors...")
        selectors_to_try = [
            "a.accordion-link-download",
            "a[href*='.zip']",
            "a[href*='aste']",
            "a.accordion-link",
        ]

        for selector in selectors_to_try:
            try:
                links = driver.find_elements(By.CSS_SELECTOR, selector)
                log_debug(f"  Selector '{selector}': found {len(links)} links")
                for link in links:
                    href = link.get_attribute('href')
                    if href and href.endswith('.zip'):
                        link_text = link.text.strip()
                        log_debug(f"[OK] Found via selector: {link_text}", "SUCCESS")
                        return link, href, link_text
            except:
                continue

        # Method 3: Look inside accordion containers
        log_debug("Method 3: Searching accordion containers...")
        try:
            accordions = driver.find_elements(By.CSS_SELECTOR, "div.accordion-date, div[class*='accordion']")
            log_debug(f"  Found {len(accordions)} accordion containers")

            for accordion in accordions:
                links = accordion.find_elements(By.TAG_NAME, "a")
                for link in links:
                    href = link.get_attribute('href')
                    if href and href.endswith('.zip'):
                        link_text = link.text.strip()
                        log_debug(f"[OK] Found in accordion: {link_text}", "SUCCESS")
                        return link, href, link_text
        except Exception as e:
            log_debug(f"  Accordion search failed: {e}", "WARNING")

        log_debug("Could not find auction file link", "ERROR")
        log_debug("Saving page source for debugging...", "INFO")

        # Save page source for debugging
        try:
            with open("page_source_debug.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            log_debug("Page source saved to: page_source_debug.html", "INFO")
        except:
            pass

        return None, None, None

    except Exception as e:
        log_debug(f"Error finding auction link: {str(e)}", "ERROR")
        import traceback
        traceback.print_exc()
        return None, None, None


def wait_for_download_complete(download_dir: str, timeout: int = 60):
    """
    Wait for download to complete by checking for .crdownload files
    """
    log_debug("Waiting for download to complete...")

    start_time = time.time()
    while time.time() - start_time < timeout:
        # Check for .crdownload files (Chrome partial download)
        crdownload_files = list(Path(download_dir).glob("*.crdownload"))

        if not crdownload_files:
            # Check if we have any .zip files
            zip_files = list(Path(download_dir).glob("*.zip"))
            if zip_files:
                latest_file = max(zip_files, key=lambda p: p.stat().st_mtime)
                log_debug(f"Download complete: {latest_file.name}", "SUCCESS")
                return str(latest_file)

        time.sleep(1)

    log_debug("Download timeout", "ERROR")
    return None


def extract_zip_file(zip_path: str, extract_to: str):
    """
    Extract ZIP file and return path to extracted XLS file
    """
    log_debug(f"Extracting ZIP file: {zip_path}")

    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)

            # List extracted files
            extracted_files = zip_ref.namelist()
            log_debug(f"Extracted {len(extracted_files)} files")

            # Find XLS file
            xls_files = [f for f in extracted_files if f.endswith('.xls') or f.endswith('.xlsx')]

            if xls_files:
                xls_file = xls_files[0]
                xls_path = os.path.join(extract_to, xls_file)
                log_debug(f"Found XLS file: {xls_file}", "SUCCESS")
                return xls_path
            else:
                log_debug("No XLS file found in ZIP", "ERROR")
                return None

    except Exception as e:
        log_debug(f"Error extracting ZIP: {str(e)}", "ERROR")
        return None


# =============================================================================
# MAIN DOWNLOAD FUNCTION
# =============================================================================

def download_latest_auction_file():
    """
    Main function to download and extract the latest auction file

    Returns:
        str: Path to extracted XLS file, or None if failed
    """
    log_debug("\n" + "="*80)
    log_debug("STARTING AUCTION FILE DOWNLOAD")
    log_debug("="*80 + "\n")

    ensure_directories()
    driver = None

    try:
        # Setup driver
        driver = setup_driver()

        # Navigate to page
        log_debug(f"Navigating to: {BASE_URL}")
        driver.get(BASE_URL)
        time.sleep(PAGE_LOAD_DELAY)

        # Handle cookie banner
        handle_cookie_banner(driver)

        # Trigger translation (manual step noted)
        trigger_page_translation(driver)

        # Find download link
        link_element, download_url, link_text = find_latest_auction_file_link(driver)

        if not link_element:
            log_debug("Failed to find download link", "ERROR")
            return None

        # Click download link
        log_debug("Clicking download link...")
        safe_click(driver, link_element, "Download link")

        # Wait for download to complete
        downloaded_file = wait_for_download_complete(DOWNLOAD_DIR, timeout=60)

        if not downloaded_file:
            log_debug("Download failed or timed out", "ERROR")
            return None

        log_debug(f"Downloaded file: {downloaded_file}", "SUCCESS")

        # Extract ZIP file
        xls_file_path = extract_zip_file(downloaded_file, EXTRACTED_DIR)

        if xls_file_path:
            log_debug("\n" + "="*80)
            log_debug("DOWNLOAD AND EXTRACTION COMPLETE")
            log_debug("="*80)
            log_debug(f"XLS file ready at: {xls_file_path}")
            return xls_file_path
        else:
            return None

    except Exception as e:
        log_debug(f"Critical error in download process: {str(e)}", "ERROR")
        import traceback
        traceback.print_exc()
        return None

    finally:
        if driver:
            log_debug("Closing browser...")
            driver.quit()


# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Main execution function"""
    start_time = time.time()

    print("\n" + "="*80)
    print("ITBI AUCTION FILE DOWNLOADER")
    print("Banca d'Italia - Government Bond Auction Results")
    print("="*80 + "\n")

    xls_file = download_latest_auction_file()

    elapsed_time = time.time() - start_time

    if xls_file:
        print("\n[SUCCESS] File downloaded and extracted!")
        print(f"XLS file location: {xls_file}")
        print(f"Total time: {elapsed_time:.2f} seconds")
        return xls_file
    else:
        print("\n[FAILED] Download failed!")
        print("Please check the logs above for details.")
        return None


if __name__ == "__main__":
    result = main()
    if result:
        print("\nNext step: Parse XLS file and extract auction data")
