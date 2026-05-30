import os
import time
import logging
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import openpyxl

# --- Configuration ---
EXCEL_PATH = "bookings.xlsx"
OUTPUT_PATH = "bookings_updated.xlsx"
BASE_URL = "https://www.trackon.in"

# Logging Setup
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class TrackonAutomation:
    def __init__(self, excel_path, username=None, password=None, captcha_api_key=None, headless=False, run_id=None):
        self.excel_path = os.path.join(os.getcwd(), excel_path)
        self.username = username
        self.password = password
        self.captcha_api_key = captcha_api_key
        self.headless = headless
        
        self.driver = None
        self.wait = None
        
        # State tracking for API status queries
        self.state = {
            "run_id": run_id,
            "status": "queued",
            "progress": 0,
            "total_rows": 0,
            "processed_rows": 0,
            "weights_updated": 0,
            "logs": [],
            "error": None,
            "start_time": None,
            "end_time": None
        }
        self.log_progress("Automation job initialized.")

    def log_progress(self, message):
        """Helper to log messages and store them in the run state."""
        logger.info(message)
        self.state["logs"].append(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")
        if self.state.get("run_id"):
            try:
                import json
                run_dir = os.path.join(os.getcwd(), "runs", self.state["run_id"])
                os.makedirs(run_dir, exist_ok=True)
                with open(os.path.join(run_dir, "metadata.json"), "w") as f:
                    json.dump(self.state, f, indent=4)
            except Exception as meta_err:
                logger.error(f"Failed to write run metadata: {meta_err}")

    def setup_driver(self):
        self.log_progress("Initializing Chrome Driver...")
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        
        if self.headless:
            self.log_progress("Running Chrome in Headless Mode...")
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            # Set a standard User Agent to act like a real browser
            options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
            
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        self.wait = WebDriverWait(self.driver, 20)

    def login_automated(self):
        """Fills franchisee login credentials, captures CAPTCHA, solves it using 2Captcha, and submits."""
        import captcha_solver
        
        # 1. Fill Username
        self.log_progress("Entering Franchisee ID...")
        user_field = self.wait.until(EC.presence_of_element_located((By.ID, "txtuserId")))
        user_field.clear()
        user_field.send_keys(self.username)
        
        # 2. Fill Password
        self.log_progress("Entering password...")
        pass_field = self.wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))
        pass_field.clear()
        pass_field.send_keys(self.password)
        
        # 3. Capture CAPTCHA Image screenshot
        self.log_progress("Capturing CAPTCHA image element...")
        captcha_img = self.wait.until(EC.presence_of_element_located((By.ID, "captchaImage")))
        time.sleep(1) # Small buffer for rendering
        captcha_bytes = captcha_img.screenshot_as_png
        
        # 4. Solve CAPTCHA via 2Captcha
        self.log_progress("Requesting CAPTCHA solution from 2Captcha service...")
        captcha_text = captcha_solver.solve_captcha(captcha_bytes, self.captcha_api_key)
        self.log_progress(f"CAPTCHA solved! Result: '{captcha_text}'")
        
        # 5. Fill CAPTCHA input field
        captcha_field = self.wait.until(EC.presence_of_element_located((By.NAME, "CaptchaCode")))
        captcha_field.clear()
        captcha_field.send_keys(captcha_text)
        
        # 6. Click SIGN IN
        self.log_progress("Clicking SIGN IN...")
        signin_btn = self.wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'btn-signin')] | //button[contains(text(), 'SIGN IN')]")))
        self.driver.execute_script("arguments[0].click();", signin_btn)
        
        # 7. Confirm Login Success
        time.sleep(2)
        try:
            self.log_progress("Verifying dashboard load...")
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Reports')] | //a[contains(., 'Reports')]"))
            )
            self.log_progress("Automated Headless Login Verified successfully!")
        except TimeoutException:
            # Check for rejected credentials/captcha errors in the page source
            page_text = self.driver.page_source
            if any(x in page_text.lower() for x in ["invalid", "wrong", "error", "incorrect", "captcha"]):
                raise RuntimeError("Login failed: Invalid CAPTCHA code or incorrect franchisee credentials.")
            raise RuntimeError("Dashboard page took too long to load.")

    def login_manual(self):
        """Navigates directly to the Franchisee Login portal and handles automated or manual login."""
        login_url = "https://ba.trackon.in/"
        self.log_progress(f"Navigating directly to Franchisee Login: {login_url}")
        self.driver.get(login_url)
        
        try:
            # Check if we should execute fully automated headless login
            if self.headless and self.username and self.password and self.captcha_api_key:
                self.log_progress("Automated login credentials found. Attempting headless 2Captcha login...")
                try:
                    self.login_automated()
                    return
                except Exception as e:
                    self.log_progress(f"Automated login attempt failed: {e}. Retrying login...")
                    # Try reloading page and solving captcha again (up to 2 retry attempts)
                    for attempt in range(2):
                        try:
                            self.driver.get(login_url)
                            time.sleep(3)
                            self.login_automated()
                            return
                        except Exception as retry_err:
                            self.log_progress(f"Login retry attempt {attempt + 1} failed: {retry_err}")
                    raise RuntimeError("Headless automated login failed after multiple attempts.")
            
            # Manual fallback mode (if running visual local mode or key is missing)
            self.log_progress("Falling back to manual browser login...")
            print("\n" + "!" * 60)
            print("ACTION REQUIRED: PLEASE LOG IN MANUALLY IN THE BROWSER WINDOW.")
            print("1. Enter your Franchisee credentials.")
            print("2. Solve the CAPTCHA correctly.")
            print("3. Click 'SIGN IN'.")
            print("4. Wait until you see the 'Dashboards' screen.")
            print("!" * 60 + "\n")
            
            # Wait up to 300 seconds (5 minutes) for the user to complete the manual login
            self.log_progress("Waiting for manual login (up to 300 seconds)...")
            WebDriverWait(self.driver, 300).until(
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Reports')] | //a[contains(., 'Reports')]"))
            )
            self.log_progress("Login confirmed. Proceeding with automation...")

        except Exception as e:
            self.log_progress(f"Error during login navigation: {e}")
            raise

    def navigate_to_reports(self):
        """Navigates to Reports -> Booking Information."""
        self.log_progress("Navigating to Reports > Booking Information...")
        try:
            # Wait for customPreloader to fade out/become invisible if it exists
            try:
                self.log_progress("Waiting for dashboard preloader to disappear...")
                WebDriverWait(self.driver, 15).until(
                    EC.invisibility_of_element_located((By.ID, "customPreloader"))
                )
                self.log_progress("Preloader is gone.")
            except Exception:
                self.log_progress("Preloader check timed out or not present, proceeding...")
                
            # Click Reports (Side menu)
            reports_menu = self.wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Reports')] | //a[contains(., 'Reports')]")))
            self.driver.execute_script("arguments[0].click();", reports_menu)
            time.sleep(1) # Small pause for submenu expansion
            
            # Click Booking Information
            booking_info = self.wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Booking Information")))
            self.driver.execute_script("arguments[0].click();", booking_info)
            
            # Wait for the search form to load using the new ID tbfromDate
            time.sleep(3) # Give page transition some time
            self.wait.until(EC.presence_of_element_located((By.ID, "tbfromDate")))
            self.log_progress("Report page loaded successfully.")
        except Exception as e:
            self.log_progress(f"Failed to navigate to reports: {e}")
            try:
                inputs = self.driver.find_elements(By.TAG_NAME, "input")
                self.log_progress(f"Current page URL: {self.driver.current_url}")
                self.log_progress(f"Input elements found: {[i.get_attribute('id') for i in inputs]}")
            except Exception:
                pass
            raise

    def perform_search(self, from_date, to_date):
        """Inputs dates and performs search."""
        from_str = from_date.strftime("%Y-%m-%d")
        to_str = to_date.strftime("%Y-%m-%d")
        self.log_progress(f"==> Searching Range: {from_date.strftime('%d-%m-%Y')} to {to_date.strftime('%d-%m-%Y')}")
        
        try:
            # Set Date values via JavaScript
            from_input = self.driver.find_element(By.ID, "tbfromDate")
            self.driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", from_input, from_str)
            
            to_input = self.driver.find_element(By.ID, "tbtoDate")
            self.driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", to_input, to_str)
            
            # Select Product Type = "ALL"
            product_select_elem = self.driver.find_element(By.ID, "ddlproductype")
            product_select = Select(product_select_elem)
            try:
                product_select.select_by_value("All")
                self.log_progress("Selected Product Type: 'All Product'")
            except Exception:
                try:
                    product_select.select_by_visible_text("All Product")
                    self.log_progress("Selected Product Type: 'All Product'")
                except Exception as p_err:
                    self.log_progress(f"Could not select 'All Product', choosing first: {p_err}")
                    product_select.select_by_index(0)
                    
            # Select Report Type = Booking Details (Value: 'D')
            try:
                report_select_elem = self.driver.find_element(By.ID, "ddlreporttype")
                report_select = Select(report_select_elem)
                report_select.select_by_value("D")
                self.log_progress("Selected Report Type: 'Booking Details - 7 Days'")
            except Exception as report_err:
                self.log_progress(f"Could not select ddlreporttype: {report_err}")
            
            # Click SEARCH via Javascript
            search_btn = self.driver.find_element(By.ID, "btnSearch")
            self.driver.execute_script("arguments[0].click();", search_btn)
            
            # Wait for search results to fetch
            time.sleep(5) 
            self.log_progress("Search triggered. Waiting for results...")
            
        except Exception as e:
            self.log_progress(f"Search input failed: {e}")

    def extract_table_data(self):
        """Scrapes the visible table with dynamic column detection and page navigation."""
        extracted_data = {}
        try:
            self.log_progress("Waiting for table results to load (up to 15 seconds)...")
            start_time = time.time()
            table_loaded = False
            
            while time.time() - start_time < 15:
                try:
                    rows = self.driver.find_elements(By.XPATH, "//table[@id='tbldetails']/tbody/tr")
                    if rows:
                        first_row_text = rows[0].text.strip()
                        # Check for placeholder rows
                        if any(x in first_row_text for x in ["No data available", "No records found", "No booking found"]):
                            self.log_progress(f"Placeholder detected: '{first_row_text}'. No records for this range.")
                            return extracted_data
                        
                        # Verify the first row actually has cells
                        cells = rows[0].find_elements(By.TAG_NAME, "td")
                        if len(cells) > 2:
                            self.log_progress(f"Loaded row detected! First cell content: '{cells[0].text.strip()}'. Scraping...")
                            table_loaded = True
                            break
                except Exception:
                    pass
                time.sleep(0.5)

            if not table_loaded:
                self.log_progress("Table loading timed out. Checking page source as fallback...")
                if "No data available" in self.driver.page_source or "No records found" in self.driver.page_source:
                    self.log_progress("No records found for this range.")
                    return extracted_data

            while True:
                # Find the table and headers
                headers = self.driver.find_elements(By.XPATH, "//table[@id='tbldetails']/thead/tr/th")
                col_map = {th.text.strip().upper(): i for i, th in enumerate(headers)}
                
                # Dynamically locate columns based on substrings
                awb_idx = None
                weight_idx = None
                for col_name, idx in col_map.items():
                    if any(x in col_name for x in ["AWB", "CNO", "CNONO", "DOCKET", "C.NOTE"]):
                        awb_idx = idx
                    elif "WEIGHT" in col_name:
                        weight_idx = idx

                # Defaults if column header detection was imperfect
                if awb_idx is None:
                    awb_idx = col_map.get("CNONO", col_map.get("AWB NO", 1))
                if weight_idx is None:
                    weight_idx = col_map.get("WEIGHT", 4)
                
                rows = self.driver.find_elements(By.XPATH, "//table[@id='tbldetails']/tbody/tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) > max(awb_idx, weight_idx):
                        awb = cells[awb_idx].text.strip()
                        weight = cells[weight_idx].text.strip()
                        if awb:
                            extracted_data[awb] = weight
                
                # Pagination logic for the DataTable structure
                try:
                    next_btn = None
                    for possible_id in ["tbldetails_next", "example_next"]:
                        try:
                            next_btn = self.driver.find_element(By.ID, possible_id)
                            break
                        except NoSuchElementException:
                            continue
                    
                    if not next_btn:
                        possible_nexts = self.driver.find_elements(By.XPATH, "//a[contains(@class, 'next') or contains(text(), 'Next')] | //button[contains(@class, 'next') or contains(text(), 'Next')]")
                        if possible_nexts:
                            next_btn = possible_nexts[0]
                    
                    if not next_btn or "disabled" in next_btn.get_attribute("class") or next_btn.get_attribute("disabled") is not None:
                        break
                        
                    self.driver.execute_script("arguments[0].scrollIntoView();", next_btn)
                    self.driver.execute_script("arguments[0].click();", next_btn)
                    time.sleep(2)
                except Exception as pag_err:
                    self.log_progress(f"No more pages or pagination completed: {pag_err}")
                    break
                    
        except Exception as e:
            self.log_progress(f"Error during data extraction: {e}")
            
        self.log_progress(f"Extracted {len(extracted_data)} unique AWB weights in this batch.")
        return extracted_data

    def run(self):
        self.state["status"] = "running"
        self.state["start_time"] = datetime.now().isoformat()
        
        try:
            # 0. Validate Excel
            if not os.path.exists(self.excel_path):
                raise FileNotFoundError(f"{self.excel_path} not found.")

            self.log_progress(f"Loading data from {self.excel_path} using openpyxl...")
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
            
            # Find header row dynamically (supports shifting tables due to header metadata)
            header_row_idx = None
            headers = None
            for r_idx in range(1, 30):
                row_vals = [cell.value for cell in ws[r_idx]]
                normalized_vals = [str(v).strip().upper() for v in row_vals if v is not None]
                has_date = any("DATE" in v for v in normalized_vals)
                has_cno = any(any(x in v for x in ["CNO", "C.NOTE", "AWB", "DOCKET", "C.NO", "C NO"]) for v in normalized_vals)
                
                if has_date and has_cno:
                    header_row_idx = r_idx
                    headers = row_vals
                    self.log_progress(f"Dynamically detected header row at Row {header_row_idx}: {headers}")
                    break
                    
            if not header_row_idx:
                self.log_progress("Could not dynamically find header row, falling back to Row 11.")
                header_row_idx = 11
                headers = [cell.value for cell in ws[11]]

            header_map = {}
            for col_idx, val in enumerate(headers, start=1):
                if val:
                    header_map[str(val).strip().upper()] = col_idx
            
            # Find exact keys
            cno_col_idx = None
            date_col_idx = None
            weight_col_idx = None
            
            for k, idx in header_map.items():
                if any(x in k for x in ["CNO", "C.NOTE", "AWB", "DOCKET", "C.NO", "C NO"]):
                    cno_col_idx = idx
                elif "DATE" in k:
                    date_col_idx = idx
                elif "WEIGHT" in k:
                    weight_col_idx = idx
            
            if not cno_col_idx or not date_col_idx:
                raise ValueError(f"Missing required columns (C.NOTE.NO / CNO. or DATE). Headers found: {headers}")

            if not weight_col_idx:
                weight_col_idx = 5
                self.log_progress(f"WEIGHT column not found in headers, defaulting to Column {weight_col_idx}")

            # Read all rows starting from row after headers dynamically
            rows_data = []
            valid_dates = []
            
            for row_idx in range(header_row_idx + 1, ws.max_row + 1):
                cno_val = ws.cell(row=row_idx, column=cno_col_idx).value
                date_val = ws.cell(row=row_idx, column=date_col_idx).value
                weight_val = ws.cell(row=row_idx, column=weight_col_idx).value
                
                if cno_val is None:
                    continue # Skip empty row
                
                # Parse date_val
                parsed_date = None
                if isinstance(date_val, datetime):
                    parsed_date = date_val
                elif isinstance(date_val, str):
                    date_str = date_val.strip()
                    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d"):
                        try:
                            parsed_date = datetime.strptime(date_str, fmt)
                            break
                        except ValueError:
                            continue
                
                if parsed_date:
                    valid_dates.append(parsed_date)
                
                rows_data.append({
                    'row_idx': row_idx,
                    'cno': cno_val,
                    'date': parsed_date,
                    'weight': weight_val
                })

            if not rows_data:
                raise ValueError("Excel data is empty.")

            if not valid_dates:
                raise ValueError("No valid dates found in the DATE column.")

            min_date = min(valid_dates)
            max_date = max(valid_dates)
            self.state["total_rows"] = len(rows_data)
            
            self.log_progress(f"Loaded {len(rows_data)} rows. Date range found: {min_date.strftime('%Y-%m-%d')} to {max_date.strftime('%Y-%m-%d')}")

            # Initialize selenium
            self.setup_driver()
            self.login_manual()
            self.navigate_to_reports()

            all_extracted_weights = {}

            # Execute searches in chunks
            temp_start = min_date
            while temp_start <= max_date:
                temp_end = min(temp_start + timedelta(days=6), max_date)
                
                self.perform_search(temp_start, temp_end)
                batch_results = self.extract_table_data()
                all_extracted_weights.update(batch_results)
                
                temp_start = temp_end + timedelta(days=1)

            # Write results back to Excel
            self.log_progress("Writing weights back to Excel...")
            count = 0
            for row in rows_data:
                cno_id = str(row['cno']).strip().split('.')[0]
                
                self.state["processed_rows"] += 1
                # Progress percentage
                self.state["progress"] = int((self.state["processed_rows"] / self.state["total_rows"]) * 100)
                
                if cno_id in all_extracted_weights:
                    if row['weight'] is not None and str(row['weight']).strip() != "":
                        continue
                    
                    # Formats weight as "{weight} KG"
                    raw_weight = str(all_extracted_weights[cno_id]).strip()
                    if raw_weight:
                        cleaned_weight = raw_weight.upper().replace("KG", "").strip()
                        formatted_weight = f"{cleaned_weight} KG"
                        ws.cell(row=row['row_idx'], column=weight_col_idx).value = formatted_weight
                        count += 1

            self.state["weights_updated"] = count

            # Save the updated workbook
            # If a custom output path is provided in metadata, write there
            output_dir = os.path.dirname(self.excel_path)
            output_file = os.path.join(output_dir, "bookings_updated.xlsx")
            wb.save(output_file)
            
            self.log_progress(f"Update Complete! {count} weights updated and formatting preserved.")
            self.log_progress(f"File saved successfully to: {output_file}")
            
            self.state["status"] = "completed"
            self.state["progress"] = 100

        except Exception as e:
            self.state["status"] = "failed"
            self.state["error"] = str(e)
            self.log_progress(f"FATAL ERROR ENCOUNTERED: {e}")
        finally:
            if self.driver:
                self.log_progress("Closing browser...")
                self.driver.quit()
            self.state["end_time"] = datetime.now().isoformat()
            self.log_progress("--- Session Finished ---")

if __name__ == "__main__":
    from dotenv import load_dotenv
    load_dotenv()
    
    username = os.getenv("TRACKON_USERNAME")
    password = os.getenv("TRACKON_PASSWORD")
    captcha_key = os.getenv("TWOCAPTCHA_API_KEY")
    
    # Run headlessly if credentials and 2Captcha key exist, otherwise fall back to manual interactive mode
    run_headless = bool(username and password and captcha_key)
    if run_headless:
        logger.info("Credentials and CAPTCHA key detected in .env. Running in headless automated mode!")
    else:
        logger.info("No CAPTCHA key or credentials found. Running in interactive manual fallback mode!")
        
    app = TrackonAutomation(
        excel_path=EXCEL_PATH,
        username=username,
        password=password,
        captcha_api_key=captcha_key,
        headless=run_headless
    )
    app.run()
