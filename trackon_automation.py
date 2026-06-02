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
        try:
            # Locate input elements
            from_input = self.wait.until(EC.presence_of_element_located((By.ID, "tbfromDate")))
            to_input = self.wait.until(EC.presence_of_element_located((By.ID, "tbtoDate")))
            
            # Dynamically detect portal date format from pre-populated value
            default_val = from_input.get_attribute("value") or ""
            input_type = from_input.get_attribute("type") or "text"
            self.log_progress(f"Detected default date value on portal: '{default_val}', input type: '{input_type}'")
            
            if input_type == "date":
                fmt = "%Y-%m-%d"
            elif "/" in default_val:
                fmt = "%d/%m/%Y"
            elif "-" in default_val:
                parts = default_val.split("-")
                if len(parts) == 3 and len(parts[0]) == 4:
                    fmt = "%Y-%m-%d"
                else:
                    fmt = "%d-%m-%Y"
            else:
                # Default fallback for Trackon/Indian format
                fmt = "%d-%m-%Y"
                
            from_str = from_date.strftime(fmt)
            to_str = to_date.strftime(fmt)
            self.log_progress(f"==> Searching Range (Detected Format: {fmt}): {from_str} to {to_str}")
            
            # Set Date values via JavaScript
            self.driver.execute_script(
                "arguments[0].value = arguments[1]; "
                "arguments[0].setAttribute('value', arguments[1]); "
                "arguments[0].dispatchEvent(new Event('change')); "
                "arguments[0].dispatchEvent(new Event('input'));", 
                from_input, from_str
            )
            self.driver.execute_script(
                "arguments[0].value = arguments[1]; "
                "arguments[0].setAttribute('value', arguments[1]); "
                "arguments[0].dispatchEvent(new Event('change')); "
                "arguments[0].dispatchEvent(new Event('input'));", 
                to_input, to_str
            )
            
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
            
            # Click SEARCH via Javascript with AJAX staleness syncing
            search_btn = self.driver.find_element(By.ID, "btnSearch")
            
            try:
                old_table = self.driver.find_element(By.ID, "tbldetails")
            except Exception:
                old_table = None
                
            self.driver.execute_script("arguments[0].click();", search_btn)
            self.log_progress("Search triggered. Waiting for results...")
            
            # Wait for AJAX reload to complete (staleness of the old table element)
            if old_table:
                try:
                    self.log_progress("Waiting for old results to clear...")
                    WebDriverWait(self.driver, 10).until(EC.staleness_of(old_table))
                    self.log_progress("Table refreshed. Loading new records...")
                except Exception as sync_err:
                    self.log_progress(f"Table refresh wait completed/skipped: {sync_err}")
            else:
                time.sleep(5)
            
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

            previous_first_awb = None
            
            while True:
                # Scrape current page with retry logic for StaleElementReferenceExceptions
                scraped_page_successfully = False
                for attempt in range(3):
                    try:
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
                        if not rows:
                            scraped_page_successfully = True
                            break
                            
                        # Robust pagination refresh check
                        first_row_cells = rows[0].find_elements(By.TAG_NAME, "td")
                        if len(first_row_cells) > awb_idx:
                            current_first_awb = first_row_cells[awb_idx].text.strip()
                            if previous_first_awb is not None and current_first_awb == previous_first_awb:
                                # Table has not updated to next page yet, sleep and wait for change
                                time.sleep(1.0)
                                rows = self.driver.find_elements(By.XPATH, "//table[@id='tbldetails']/tbody/tr")
                        
                        if rows:
                            first_row_cells = rows[0].find_elements(By.TAG_NAME, "td")
                            if len(first_row_cells) > awb_idx:
                                previous_first_awb = first_row_cells[awb_idx].text.strip()
                        
                        # Fetch text values inside this try-catch to prevent stale element crashes
                        page_data = {}
                        for row in rows:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            if len(cells) > max(awb_idx, weight_idx):
                                awb = cells[awb_idx].text.strip()
                                weight = cells[weight_idx].text.strip()
                                if awb:
                                    page_data[awb] = weight
                                    
                        # Update global results on success
                        extracted_data.update(page_data)
                        scraped_page_successfully = True
                        break
                        
                    except Exception as scrape_err:
                        self.log_progress(f"Temporary page scrape error (attempt {attempt + 1}/3): {scrape_err}")
                        time.sleep(1.5) # Sleep and let DOM settle before retrying
                
                if not scraped_page_successfully:
                    self.log_progress("Failed to scrape current page after 3 attempts due to DOM instability. Moving on...")
                    break
                
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
                    time.sleep(2.5) # Increased sleep slightly for safety
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
            
            # 1. Scan all sheets and analyze candidates
            sheet_candidates = []
            date_sheets = []
            
            for sheet_name in wb.sheetnames:
                temp_ws = wb[sheet_name]
                header_row_idx = None
                headers = None
                
                # Check for standard header row in this sheet
                for r_idx in range(1, 30):
                    row_vals = [cell.value for cell in temp_ws[r_idx]]
                    normalized_vals = [str(v).strip().upper() for v in row_vals if v is not None]
                    has_date = any("DATE" in v for v in normalized_vals)
                    has_cno = any(any(x in v for x in ["CNO", "C.NOTE", "AWB", "DOCKET", "C.NO", "C NO"]) for v in normalized_vals)
                    
                    if has_date and has_cno:
                        header_row_idx = r_idx
                        headers = row_vals
                        break
                
                if header_row_idx:
                    # Map columns in this sheet candidate
                    header_map = {str(val).strip().upper(): col_idx for col_idx, val in enumerate(headers, start=1) if val}
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
                            
                    if not weight_col_idx:
                        weight_col_idx = 5
                        
                    if not cno_col_idx or not date_col_idx:
                        continue # Required columns missing in this candidate, skip it
                        
                    # Count total rows and rows with missing weights
                    total_data_rows = 0
                    missing_weight_count = 0
                    
                    for row_idx in range(header_row_idx + 1, temp_ws.max_row + 1):
                        cno_val = temp_ws.cell(row=row_idx, column=cno_col_idx).value
                        if cno_val is not None:
                            total_data_rows += 1
                            weight_val = temp_ws.cell(row=row_idx, column=weight_col_idx).value
                            if weight_val is None or str(weight_val).strip() == "":
                                missing_weight_count += 1
                    
                    # Try to parse sheet name as a date format
                    parsed_sheet_date = None
                    sheet_name_clean = sheet_name.strip()
                    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%y", "%d-%m-%y"):
                        try:
                            parsed_sheet_date = datetime.strptime(sheet_name_clean, fmt)
                            break
                        except ValueError:
                            continue
                            
                    candidate = {
                        "name": sheet_name,
                        "ws": temp_ws,
                        "header_row_idx": header_row_idx,
                        "headers": headers,
                        "cno_col_idx": cno_col_idx,
                        "date_col_idx": date_col_idx,
                        "weight_col_idx": weight_col_idx,
                        "total_rows": total_data_rows,
                        "missing_weights": missing_weight_count,
                        "parsed_date": parsed_sheet_date
                    }
                    sheet_candidates.append(candidate)
                    if parsed_sheet_date:
                        date_sheets.append(candidate)

            if not sheet_candidates:
                raise ValueError("Could not find any sheet containing the required DATE and C.NOTE.NO columns.")

            # Log sheet analysis transparently
            self.log_progress("Analyzing sheets in workbook:")
            for cand in sheet_candidates:
                date_str_log = f"parsed date: {cand['parsed_date'].strftime('%Y-%m-%d')}" if cand['parsed_date'] else "no date name"
                self.log_progress(f" - Sheet '{cand['name']}' ({date_str_log}): {cand['total_rows']} rows, {cand['missing_weights']} missing weights")

            selected_sheets = []
            
            # Scenario A: If some sheet names are valid dates, select the one with the latest date!
            if date_sheets:
                # Sort by parsed date descending
                date_sheets.sort(key=lambda x: x["parsed_date"], reverse=True)
                selected_sheet = date_sheets[0]
                self.log_progress(f"==> Selected Sheet '{selected_sheet['name']}' because it has the latest date name ({selected_sheet['parsed_date'].strftime('%d-%m-%Y')}).")
                selected_sheets = [selected_sheet]
            else:
                # Scenario B: Generic sheet names - scan and process ALL valid data sheets!
                self.log_progress("==> Generic sheet names detected. Selecting ALL valid sheets to process together.")
                selected_sheets = sheet_candidates
            
            # 2. Compile rows from selected worksheets
            rows_data = []
            valid_dates = []
            
            for target_sheet in selected_sheets:
                t_ws = target_sheet["ws"]
                t_header_row_idx = target_sheet["header_row_idx"]
                t_cno_col = target_sheet["cno_col_idx"]
                t_date_col = target_sheet["date_col_idx"]
                t_weight_col = target_sheet["weight_col_idx"]
                
                self.log_progress(f"Parsing rows from Sheet '{target_sheet['name']}'...")
                
                for row_idx in range(t_header_row_idx + 1, t_ws.max_row + 1):
                    cno_val = t_ws.cell(row=row_idx, column=t_cno_col).value
                    date_val = t_ws.cell(row=row_idx, column=t_date_col).value
                    weight_val = t_ws.cell(row=row_idx, column=t_weight_col).value
                    
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
                        # Ignore future outlier dates (e.g. year typos) for global boundary calculations
                        if parsed_date <= datetime.now() + timedelta(days=1):
                            valid_dates.append(parsed_date)
                        else:
                            self.log_progress(f"Skipping future outlier date in global range: {parsed_date.strftime('%d-%m-%Y')} on Row {row_idx} in Sheet '{target_sheet['name']}'")
                    
                    rows_data.append({
                        'ws': t_ws,
                        'weight_col_idx': t_weight_col,
                        'row_idx': row_idx,
                        'cno': cno_val,
                        'date': parsed_date,
                        'weight': weight_val
                    })

            if not rows_data:
                raise ValueError("Excel data is empty.")

            if not valid_dates:
                raise ValueError("No valid dates found in the DATE column.")

            # Add a 2-day padding buffer to dates to handle manifest delays
            min_date = min(valid_dates) - timedelta(days=2)
            max_date = min(max(valid_dates) + timedelta(days=2), datetime.now() + timedelta(days=1))
            self.state["total_rows"] = len(rows_data)
            
            self.log_progress(f"Loaded {len(rows_data)} rows. Date range found (with 2-day padding buffer): {min_date.strftime('%d-%m-%Y')} to {max_date.strftime('%d-%m-%Y')}")

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
                # 1. Secure integer string cleaning for AWBs (prevents float & scientific format mismatches)
                cno_raw = row['cno']
                cno_id = ""
                if isinstance(cno_raw, (int, float)):
                    cno_id = f"{int(cno_raw)}"
                elif cno_raw is not None:
                    cno_id = str(cno_raw).strip().split('.')[0]
                
                self.state["processed_rows"] += 1
                # Progress percentage
                self.state["progress"] = int((self.state["processed_rows"] / self.state["total_rows"]) * 100)
                
                # 2. Advanced Fuzzy Matcher (Exact -> Substring Fallback)
                found_weight = None
                if cno_id in all_extracted_weights:
                    found_weight = all_extracted_weights[cno_id]
                else:
                    # Fallback to substring matching (useful for prefix/suffix differences or dropped zeroes)
                    matched_key = next((k for k in all_extracted_weights if cno_id in k or k in cno_id), None)
                    if matched_key:
                        found_weight = all_extracted_weights[matched_key]
                
                if found_weight is not None:
                    if row['weight'] is not None and str(row['weight']).strip() != "":
                        continue
                    
                    # Formats weight as "{weight} KG"
                    raw_weight = str(found_weight).strip()
                    if raw_weight:
                        cleaned_weight = raw_weight.upper().replace("KG", "").strip()
                        formatted_weight = f"{cleaned_weight} KG"
                        row['ws'].cell(row=row['row_idx'], column=row['weight_col_idx']).value = formatted_weight
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
