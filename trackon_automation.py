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

    def run(self):
        pass

if __name__ == "__main__":
    app = TrackonAutomation(EXCEL_PATH)
    app.run()
