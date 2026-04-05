import os
import time
import logging
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- Configuration ---
EXCEL_PATH = "bookings.xlsx"
OUTPUT_PATH = "bookings_updated.xlsx"
BASE_URL = "https://www.trackon.in"

# Logging Setup
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class TrackonAutomation:
    def __init__(self, excel_path):
        self.excel_path = os.path.join(os.getcwd(), excel_path)
        self.driver = None
        self.wait = None

    def setup_driver(self):
        logger.info("Initializing Chrome Driver...")
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        # Try to use standard service setup
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        self.wait = WebDriverWait(self.driver, 20)

    def login_manual(self):
        """Navigates to login and waits for manual intervention."""
        self.driver.get(BASE_URL)
        
        try:
            # wait for overlay or main page
            time.sleep(3) 
            
            # 1. Click Login Dropdown (Top Right)
            logger.info("Locating Login dropdown...")
            # Based on screenshot, 'Login' is a text link with a caret
            login_menu = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Login')]")))
            login_menu.click()
            
            # 2. Click Franchisee Login
            logger.info("Clicking Franchisee Login...")
            franchise_link = self.wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Franchisee Login")))
            franchise_link.click()
            
            print("\n" + "!" * 60)
            print("ACTION REQUIRED: PLEASE LOG IN MANUALLY IN THE BROWSER WINDOW.")
            print("1. Enter your Franchisee credentials.")
            print("2. Solve the CAPTCHA correctly.")
            print("3. Click 'SIGN IN'.")
            print("4. Wait until you see the 'Dashboards' screen.")
            print("!" * 60 + "\n")
            
            # Wait for dashboard sidebar element to appear to confirm login success
            # The sidebar has 'Reports' as seen in the screenshot
            self.wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Reports')] | //a[contains(., 'Reports')]")))
            logger.info("Login confirmed. Proceeding with automation...")

        except Exception as e:
            logger.error(f"Error during login navigation: {e}")
            raise

    def navigate_to_reports(self):
        """Navigates to Reports -> Booking Information."""
        logger.info("Navigating to Reports > Booking Information...")
        try:
            # Click Reports (Side menu)
            reports_menu = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Reports')] | //a[contains(., 'Reports')]")))
            reports_menu.click()
            time.sleep(1) # Small pause for submenu expansion
            
            # Click Booking Information
            booking_info = self.wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Booking Information")))
            booking_info.click()
            
            # Wait for the search form to load
            self.wait.until(EC.presence_of_element_located((By.ID, "txtFromDate")))
            logger.info("Report page loaded.")
        except Exception as e:
            logger.error(f"Failed to navigate to reports: {e}")
            raise

    def perform_search(self, from_date, to_date):
        """Inputs dates and performs search."""
        from_str = from_date.strftime("%d-%m-%Y")
        to_str = to_date.strftime("%d-%m-%Y")
        logger.info(f"==> Searching Range: {from_str} to {to_str}")
        
        try:
            # Clear and Type From Date
            from_input = self.driver.find_element(By.ID, "txtFromDate")
            # Clear doesn't always work on date pickers, so use JS
            self.driver.execute_script("arguments[0].value = '';", from_input)
            from_input.send_keys(from_str)
            
            # Clear and Type To Date
            to_input = self.driver.find_element(By.ID, "txtToDate")
            self.driver.execute_script("arguments[0].value = '';", to_input)
            to_input.send_keys(to_str)
            
            # Select Product Type = "ALL"
            # It's an orange dropdown in the screenshot
            product_select_elem = self.driver.find_element(By.ID, "ddlProductType")
            product_select = Select(product_select_elem)
            product_select.select_by_visible_text("ALL")
            
            # Click SEARCH
            search_btn = self.driver.find_element(By.ID, "btnSearch")
            search_btn.click()
            
            # Wait for the table show message or table rows
            time.sleep(4) 
            logger.info("Search button clicked. Waiting for results...")
            
        except Exception as e:
            logger.warning(f"Search input failed: {e}")

    def extract_table_data(self):
        """Scrapes the visible table with dynamic column detection."""
        extracted_data = {}
        try:
            time.sleep(2) # Ensure table is rendered
            if "No data available" in self.driver.page_source:
                logger.info("No records found for this range.")
                return extracted_data

            while True:
                # Find the table and headers
                headers = self.driver.find_elements(By.XPATH, "//table[@id='example']/thead/tr/th")
                col_map = {th.text.strip().upper(): i for i, th in enumerate(headers)}
                
                awb_idx = col_map.get("CNONO", col_map.get("AWB NO", 1))
                weight_idx = col_map.get("WEIGHT", 10)
                
                rows = self.driver.find_elements(By.XPATH, "//table[@id='example']/tbody/tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) > max(awb_idx, weight_idx):
                        awb = cells[awb_idx].text.strip()
                        weight = cells[weight_idx].text.strip()
                        if awb:
                            extracted_data[awb] = weight
                
                # Pagination
                try:
                    next_btn = self.driver.find_element(By.ID, "example_next")
                    if "disabled" in next_btn.get_attribute("class"):
                        break
                    self.driver.execute_script("arguments[0].scrollIntoView();", next_btn)
                    next_btn.click()
                    time.sleep(2)
                except NoSuchElementException:
                    break
                    
        except Exception as e:
            logger.error(f"Error during data extraction: {e}")
            
        logger.info(f"Extracted {len(extracted_data)} unique AWB weights.")
        return extracted_data

    def run(self):
        pass

if __name__ == "__main__":
    app = TrackonAutomation(EXCEL_PATH)
    app.run()
