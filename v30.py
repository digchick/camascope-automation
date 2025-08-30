import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, ElementNotInteractableException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import json
import os
from datetime import datetime
import traceback
import shutil

class FixedDropdownAutomator:
    def __init__(self, download_path=None, driver_path=None):
        """Initialize the automation tool with enhanced element selection"""
        print("Starting browser setup...")

        # Set up Chrome driver
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        # This line suppresses the "DevTools listening on..." message and other verbose logs
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        if download_path:
            prefs = {
                "download.default_directory": download_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True
            }
            options.add_experimental_option("prefs", prefs)

        try:
            if driver_path:
                self.driver = webdriver.Chrome(executable_path=driver_path, options=options)
            else:
                self.driver = webdriver.Chrome(options=options)

            print("Browser started successfully!")

        except Exception as e:
            print(f"Error starting browser: {e}")
            print("Trying to install webdriver-manager...")
            try:
                import subprocess
                subprocess.check_call(['pip', 'install', 'webdriver-manager'])
                from webdriver_manager.chrome import ChromeDriverManager
                self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
                print("Browser started with webdriver-manager!")
            except Exception as e2:
                print(f"Failed to start browser: {e2}")
                raise

        self.wait = WebDriverWait(self.driver, 20)
        self.short_wait = WebDriverWait(self.driver, 5)
        self.progress_file = "chunking_progress.json"
    def automated_login(self, target_url):
        """
        Automates the 3-step login process and navigates to the reports page.
        """
        print("=== AUTOMATED LOGIN PROCESS ===")
        print("Reading credentials from 'camascope login.txt'...")
        try:
            with open("camascope login.txt", "r") as f:
                lines = f.readlines()
                if len(lines) >= 2:
                    plaintext_username = lines[0].strip()
                    plaintext_password = lines[1].strip()
                    print("Credentials loaded successfully.")
                else:
                    raise ValueError("Error: 'camascope login.txt' must contain username and password.")
        except FileNotFoundError:
            print("Error: 'camascope login.txt' not found.")
            raise

        try:
            print(f"Navigating to {target_url}...")
            self.driver.get(target_url)

            # --- Step 1: Click "Another User" Button ---
            print("Navigating to the 'Another User' screen...")
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'otherUserDetailsBox')]"))
            )
            another_user_button = self.driver.find_element(By.XPATH, "//div[contains(@class, 'otherUserDetailsBox')]")
            another_user_button.click()
            print("Clicked 'Another User'.")

            # --- Step 2: Enter Username and Sign In ---
            print("Entering username...")
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "signInName"))
            )
            username_input = self.driver.find_element(By.ID, "signInName")
            username_input.send_keys(plaintext_username)

            sign_in_username_button = self.driver.find_element(By.ID, "continue")
            sign_in_username_button.click()
            print("Entered username and clicked 'Sign In'.")

            # --- Step 3: Enter Password and Sign In ---
            print("Entering password...")
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "password"))
            )
            password_input = self.driver.find_element(By.ID, "password")
            password_input.send_keys(plaintext_password)

            sign_in_password_button = self.driver.find_element(By.ID, "continue")
            sign_in_password_button.click()
            print("Entered password and clicked 'Sign In'.")

            print("Login sequence complete. Waiting for the dashboard to load.")
            time.sleep(3) # Short pause for pop-ups to appear

            print("Please manually clear any pop-ups and press Enter to continue...")
            input()

            # --- Step 4: Click the 'Reports' menu option ---
            print("Clicking 'Reports' menu option...")
            reports_menu_link = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Reports"))
            )
            reports_menu_link.click()
            print("Clicked 'Reports'.")

            # --- Step 5: Click the 'MAR Report' sub-menu option ---
            print("Clicking 'MAR Report' sub-menu option...")
            mar_report_link = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "MAR Report"))
            )
            mar_report_link.click()
            print("Clicked 'MAR Report'.")

            time.sleep(2) # Allow the page to render

            # --- Get date input from user and fill fields ---
            date_choice = input("Enter 'D' to use default dates (01/09/2024 to today's date), or 'C' to enter custom dates: ").strip().lower()

            if date_choice == 'd':
                from_date = "01/09/2024"
                to_date = datetime.now().strftime("%d/%m/%Y")
                print("Using default dates.")
            elif date_choice == 'c':
                from_date = input("Please enter the 'From' date in DD/MM/YYYY format: ")
                to_date = input("Please enter the 'To' date in DD/MM/YYYY format: ")
            else:
                print("Invalid choice. Using default dates.")
                from_date = "01/09/2024"
                to_date = datetime.now().strftime("%d/%m/%Y")

            # Clear and enter 'From' date
            print(f"Entering '{from_date}' into the 'From' date field.")
            from_date_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Start Date']"))
            )
            ActionChains(self.driver).click(from_date_input).perform()
            self.driver.execute_script("arguments[0].value = '';", from_date_input)
            from_date_input.send_keys(from_date)
            from_date_input.send_keys(Keys.ENTER)
            print("Successfully entered 'From' date.")

            time.sleep(1)

            # Clear and enter 'To' date
            print(f"Entering '{to_date}' into the 'To' date field.")
            to_date_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='End Date']"))
            )
            ActionChains(self.driver).click(to_date_input).perform()
            self.driver.execute_script("arguments[0].value = '';", to_date_input)
            to_date_input.send_keys(to_date)
            to_date_input.send_keys(Keys.ENTER)
            print("Successfully entered 'To' date.")

            print("Login and navigation complete. Proceeding with main script logic.")

        except Exception as e:
            print(f"An error occurred during login or navigation: {e}")
            raise

    def get_current_first_selection(self):
        """Dynamically get the current first/top selected item from the dropdown"""
        try:
            # Strategy 1: Look for "All units" first (appears when many selections are made)
            try:
                all_units_element = self.driver.find_element(By.XPATH, "//span[text()='All units']")
                return "All units"
            except:
                pass

            # Strategy 2: Look for the first selected item badge/span
            try:
                # Look for the first span inside the badge container
                first_selection = self.driver.find_element(By.CSS_SELECTOR, ".badge.border.border-primary.text-primary")
                # Get the text content, excluding the X button
                text_content = first_selection.text.strip()
                # Remove the Ã— character if present
                if text_content.endswith('Ã—'):
                    text_content = text_content[:-1].strip()
                if text_content:
                    return text_content
            except:
                pass

            # Strategy 3: Alternative selector for selected items
            try:
                selected_spans = self.driver.find_elements(By.CSS_SELECTOR, ".css-1pahdxg-control .badge span")
                if selected_spans:
                    first_text = selected_spans[0].text.strip()
                    if first_text and first_text != 'Ã—':
                        return first_text
            except:
                pass

            # Strategy 4: Look for any selected item in the control container
            try:
                control_container = self.driver.find_element(By.CSS_SELECTOR, ".css-1pahdxg-control, .css-yk16xz-control")
                spans = control_container.find_elements(By.TAG_NAME, "span")
                for span in spans:
                    text = span.text.strip()
                    if text and text != 'Ã—' and 'Select' not in text:
                        return text
            except:
                pass

            print("Could not detect current first selection")
            return None

        except Exception as e:
            print(f"Error detecting current first selection: {e}")
            return None

    def normalize_text(self, text):
        """Normalize text to handle encoding issues"""
        if not text:
            return text

        # Common character replacements for HTML entities and encoding issues
        replacements = {
            '&amp;': '&',  # ampersand
            '&lt;': '<',   # less than
            '&gt;': '>',   # greater than
            '&quot;': '"', # quotation mark
            '&#39;': "'",  # apostrophe
            '\u2013': '-', # en dash
            '\u2014': '-', # em dash
            '\u2018': "'", # left single quotation mark
            '\u2019': "'", # right single quotation mark
            '\u201c': '"', # left double quotation mark
            '\u201d': '"', # right double quotation mark
            '\u2026': '...', # horizontal ellipsis
        }

        normalized = text
        for old, new in replacements.items():
            normalized = normalized.replace(old, new)

        return normalized.strip()

    def load_names_from_file(self, file_path, column_name="Location Name"):
        """Load names from CSV or Excel file with text normalization"""
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, encoding='utf-8-sig') # Handle BOM
        elif file_path.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path)
        else:
            raise ValueError("File must be CSV or Excel format")

        print(f"Columns in your file: {list(df.columns)}")

        # Use the Location Name column specifically
        if column_name in df.columns:
            names = df[column_name].tolist()
            print(f"Using column: '{column_name}'")
        else:
            # Fallback to second column if "Location Name" not found
            names = df.iloc[:, 1].tolist()
            print(f"Using second column: '{df.columns[1]}'")

        # Remove any NaN values, convert to strings, and normalize text
        names = [self.normalize_text(str(name)) for name in names if pd.notna(name)]

        # Show which names were normalized
        original_names = [str(name) for name in df[column_name if column_name in df.columns else df.columns[1]].tolist() if pd.notna(name)]
        for i, (orig, norm) in enumerate(zip(original_names, names)):
            if orig != norm:
                print(f"Normalized: '{orig}' â†' '{norm}'")

        return names, df

    def find_clickable_dropdown_option(self, target_text):
        """
        Find the actual clickable dropdown option element using multiple strategies
        """
        print(f"Finding clickable element for: '{target_text}'")

        # Strategy 1: Target dropdown menu options specifically
        specific_selectors = [
            f"//div[contains(@class, 'option')]//span[@class='me-2' and text()='{target_text}']",
            f"//div[contains(@class, 'option')]//span[text()='{target_text}']",
            f"//div[contains(@class, 'option') and contains(text(), '{target_text}')]",
            f"//div[contains(@class, 'menu')]//span[text()='{target_text}']",
            f"//*[@role='option' and contains(text(), '{target_text}')]",
            f"//div[contains(@class, 'option')]//text()[.='{target_text}']/parent::*"
        ]

        # Try specific selectors first
        for selector in specific_selectors:
            try:
                elements = self.driver.find_elements(By.XPATH, selector)
                if elements:
                    element = elements[0]
                    if element.is_displayed() and element.is_enabled():
                        print(f"Found clickable element using: {selector[:50]}...")
                        return element
            except:
                continue

        # Strategy 2: Fallback to general selector but filter for clickable elements
        print("Trying fallback general selector...")
        try:
            general_elements = self.driver.find_elements(By.XPATH, f"//*[text()='{target_text}']")
            for element in general_elements:
                try:
                    if element.is_displayed() and element.is_enabled():
                        # Check if it's in a dropdown context
                        parent_classes = element.find_element(By.XPATH, "./ancestor::div[1]").get_attribute("class") or ""
                        if "option" in parent_classes.lower() or "menu" in parent_classes.lower():
                            print(f"Found element via fallback method")
                            return element
                except:
                    continue
        except:
            pass

        print(f"Could not find clickable element for: '{target_text}'")
        return None

    def click_dropdown_and_select(self, target_text):
        """
        Enhanced dropdown selection with better element targeting
        """
        print(f"Attempting to select: '{target_text}'")

        # Get the current first selection dynamically
        current_anchor = self.get_current_first_selection()

        if current_anchor:
            print(f"Using current anchor: '{current_anchor}'")
            dropdown_xpath = f"//div[./span[text()='{current_anchor}']]"
        else:
            # No selections yet - look for 'Select...' text
            print("No selections found - using 'Select...' as anchor")
            dropdown_xpath = "//div[text()='Select...']"

        try:
            # Step 1: Click the dropdown to open it with retries for stale elements
            print("Waiting for dropdown button to be clickable...")

            # Retry logic for stale elements
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    dropdown_button = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
                    )

                    print("Clicking dropdown button...")
                    dropdown_button.click()
                    print("Successfully opened dropdown")
                    break # Success, exit retry loop

                except Exception as dropdown_error:
                    if "stale element" in str(dropdown_error).lower() and attempt < max_retries - 1:
                        print(f"Stale element on attempt {attempt + 1}, refreshing anchor...")
                        # Refresh the anchor and try again
                        current_anchor = self.get_current_first_selection()
                        if current_anchor:
                            dropdown_xpath = f"//div[./span[text()='{current_anchor}']]"
                        else:
                            dropdown_xpath = "//div[text()='Select...']"
                        time.sleep(0.08) # Brief pause before retry
                        continue
                    else:
                        raise dropdown_error # Re-raise if not stale element or max retries reached

            # Step 2: Wait for dropdown to load and find the clickable element
            print(f"Looking for clickable option: '{target_text}'...")
            time.sleep(0.08) # Give dropdown time to fully load

            # Use enhanced element finding
            option_element = self.find_clickable_dropdown_option(target_text)

            if not option_element:
                print(f"Could not find clickable element for '{target_text}'")
                return False

            print(f"Found clickable element - attempting click...")

            # Step 3: Try JavaScript click first (solves the Apple House issue)
            try:
                self.driver.execute_script("arguments[0].click();", option_element)
                print(f"Successfully selected using JavaScript click: '{target_text}'")
                return True
            except Exception as js_error:
                print(f"JavaScript click failed: {js_error}")

                # Fallback to standard click
                try:
                    option_element.click()
                    print(f"Successfully selected using standard click: '{target_text}'")
                    return True
                except Exception as std_error:
                    print(f"Standard click also failed: {std_error}")
                    return False

        except Exception as e:
            print(f"Failed to select '{target_text}': {str(e)}")
            print(f"   Attempted anchor: '{current_anchor if current_anchor else 'Select...'}'")
            return False

    def clear_all_selections(self):
        """Clear all selections using the Select All double-click method"""
        try:
            print("Clearing all selections...")

            # Get current anchor dynamically
            current_anchor = self.get_current_first_selection()

            if not current_anchor:
                print("No selections to clear - dropdown is already empty")
                return True

            print(f"Current anchor detected: '{current_anchor}'")

            # Use the proven Select All double-click method
            try:
                print("Using Select All double-click method...")

                # Step 1: Open dropdown using current anchor
                print(f"Opening dropdown using '{current_anchor}' as anchor...")
                dropdown_xpath = f"//div[./span[text()='{current_anchor}']]"
                dropdown_button = self.driver.find_element(By.XPATH, dropdown_xpath)
                dropdown_button.click()
                time.sleep(.05)
                print("Dropdown opened")

                # Step 2: Click Select All to select everything (this creates "All units")
                print("Looking for 'Select All' option...")
                select_all_element = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//*[text()='Select All']"))
                )

                print("Clicking 'Select All' to select everything...")
                select_all_element.click()
                time.sleep(.05) # Wait for all selections to register
                print("All items selected (including 'All units')")

                # Step 3: Get fresh anchor - should now be "All units"
                fresh_anchor = self.get_current_first_selection()
                print(f"Opening dropdown again using fresh anchor: '{fresh_anchor}'...")

                fresh_dropdown_xpath = f"//div[./span[text()='{fresh_anchor}']]"
                fresh_dropdown_button = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, fresh_dropdown_xpath))
                )
                fresh_dropdown_button.click()
                time.sleep(0.08)
                print("Dropdown reopened")

                # Step 4: Click Select All again to deselect everything
                print("Looking for 'Select All' element again to deselect...")
                time.sleep(0.08)
                select_all_elements = self.driver.find_elements(By.XPATH, "//*[text()='Select All']")

                if select_all_elements:
                    select_all_fresh = select_all_elements[0]
                    print("Clicking 'Select All' again to deselect everything...")

                    # Use JavaScript click to avoid stale element issues
                    self.driver.execute_script("arguments[0].click();", select_all_fresh)
                    time.sleep(0.08) # Wait for deselection to complete

                    print("All selections cleared!")
                    return True
                else:
                    raise Exception("Could not find fresh Select All element")

            except Exception as select_all_error:
                print(f"Select All method failed: {select_all_error}")

                # Manual fallback
                print("Automatic clearing failed. Manual intervention required.")
                print("Please manually clear the selections:")
                print("1. Click the dropdown to open it")
                print("2. Click 'Select All' to select everything")
                print("3. Click 'Select All' again to deselect everything")
                print("4. Dropdown should return to 'Select...' state")
                input("Press Enter when you've manually cleared the selections...")
                return False

        except Exception as e:
            print(f"Error during clear operation: {str(e)}")
            print("Please manually clear selections and press Enter...")
            input()
            return False

    def select_all_from_dropdown(self):
        """Use the dropdown's 'Select All' functionality to select all items"""
        try:
            print("Opening dropdown to access 'Select All' option...")

            # Open dropdown using Select... anchor (since we start with empty selection)
            dropdown_xpath = "//div[text()='Select...']"

            try:
                dropdown_button = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
                )
                dropdown_button.click()
                time.sleep(0.08)
                print("Dropdown opened successfully")
            except Exception as e:
                print(f"Failed to open dropdown: {e}")
                return False

            # Find and click Select All
            try:
                print("Looking for 'Select All' option...")
                select_all_element = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//*[text()='Select All']"))
                )

                print("Clicking 'Select All'...")
                # Use JavaScript click for reliability
                self.driver.execute_script("arguments[0].click();", select_all_element)
                time.sleep(0.08) # Wait for all selections to register

                # Verify selection worked by checking for "All units" or selected items
                current_selection = self.get_current_first_selection()
                if current_selection:
                    print(f"Success! Current selection shows: '{current_selection}'")
                    return True
                else:
                    print("Select All may not have worked - no selections detected")
                    return False

            except Exception as e:
                print(f"Failed to click Select All: {e}")
                return False

        except Exception as e:
            print(f"Error in select_all_from_dropdown: {e}")
            return False

    def find_and_click_generate_button(self):
        """
        Finds the 'Generate Report' button and clicks it.
        """
        print("Looking for 'Generate Report' button...")
        try:
            generate_button_xpath = "//button[text()='Generate Report']"

            generate_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, generate_button_xpath)))

            print("Found 'Generate Report' button. Clicking...")
            generate_button.click()
            print("Successfully clicked 'Generate Report'.")
            return True
        except TimeoutException:
            print("Timed out waiting for 'Generate Report' button to be clickable.")
            return False
        except NoSuchElementException:
            print("Could not find the 'Generate Report' button.")
            return False
        except Exception as e:
            print(f"An error occurred while clicking the generate button: {e}")
            return False

    def handle_popups_and_proceed(self):
        """
        Handles the 'Proceed Anyway' pop-up by clicking it when it appears and then
        waiting for the modal to disappear, signaling report generation has started.
        """
        print("Checking for any pop-ups to proceed...")
        try:
            proceed_button = WebDriverWait(self.driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Proceed Anyway')]"))
            )

            print("Found 'Proceed Anyway' button. Clicking to dismiss the pop-up and start report generation...")
            proceed_button.click()

            print("Successfully clicked. Now waiting for pop-up to disappear...")
            modal_xpath = "//div[contains(@class, 'modal') and contains(@class, 'show')]"
            self.wait.until(EC.invisibility_of_element_located((By.XPATH, modal_xpath)))

            print("Pop-up has disappeared. Report generation is in progress.")
            return True

        except TimeoutException:
            print("No pop-up found. Continuing...")
            return False
        except Exception as e:
            print(f"An error occurred while handling pop-up: {e}")
            return False

    def check_for_no_records_message(self):
        """
        Waits for the loading overlay to disappear, then checks for the 'No Records!' message.
        """
        print("Waiting for report loading to complete...")
        try:
            loading_overlay_xpath = "//div[contains(@class, 'ag-overlay-loading-wrapper')]"
            self.wait.until(EC.invisibility_of_element_located((By.XPATH, loading_overlay_xpath)))
            print("Loading overlay has disappeared.")
        except TimeoutException:
            print("Loading overlay did not appear or disappear in time. Continuing...")

        try:
            no_records_message_xpath = "//div[contains(@class, 'ag-center-cols-viewport') and contains(., 'No Records!')]"
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, no_records_message_xpath))
            )
            print("--- 'No Records!' message found. Skipping CSV download. ---")
            return True
        except TimeoutException:
            print("No 'No Records!' message found. Proceeding with CSV download.")
            return False

    def find_and_click_generate_csv(self):
        """
        Finds and clicks the 'Generate CSV' link, waiting for it to become clickable.
        """
        print("Waiting for 'Generate CSV' link to become clickable...")
        try:
            long_wait = WebDriverWait(self.driver, 60)

            generate_csv_link_xpath = "//a[@download and contains(text(), 'CSV')]"
            generate_csv_link = long_wait.until(EC.element_to_be_clickable((By.XPATH, generate_csv_link_xpath)))

            print(f"Found '{generate_csv_link.text}' link. Clicking...")
            generate_csv_link.click()
            print(f"Successfully clicked '{generate_csv_link.text}'.")
            return True
        except TimeoutException:
            print("Timed out waiting for 'Generate CSV' link to be clickable. Report generation may have failed or taken too long.")
            return False
        except Exception as e:
            print(f"An error occurred while clicking the generate CSV link: {e}")
            return False

    def generate_report_for_current_selections(self):
        """
        Generates and downloads report for current selections.
        Returns True if successful, False otherwise.
        """
        print("\n" + "="*50)
        print("GENERATING REPORT FOR CURRENT SELECTIONS")
        print("="*50)

        # Step 1: Click Generate Report button
        if not self.find_and_click_generate_button():
            print("Failed to click Generate Report button")
            return False

        # Step 2: Handle any popups
        self.handle_popups_and_proceed()

        # Step 3: Check for No Records message
        if self.check_for_no_records_message():
            print("No records found for current selections - skipping CSV download")
            return True # This is actually successful - just no data

        # Step 4: Download CSV if records exist
        if not self.find_and_click_generate_csv():
            print("Failed to download CSV")
            return False

        print("Report generation completed successfully!")
        return True
    def save_chunking_progress(self, session_data):
        """Save chunking progress to file"""
        try:
            with open(self.progress_file, 'w') as f:
                json.dump(session_data, f, indent=2)
            print(f"Progress saved to {self.progress_file}")
        except Exception as e:
            print(f"Error saving progress: {e}")

    def load_chunking_progress(self, progress_file):
        """Load chunking progress from file"""
        try:
            if os.path.exists(progress_file):
                with open(progress_file, 'r') as f:
                    return json.load(f)
            return None
        except Exception as e:
            print(f"Error loading progress: {e}")
            return None

    def clear_chunking_progress(self, progress_file):
        """Clear the progress file"""
        try:
            if os.path.exists(progress_file):
                os.remove(progress_file)
                print("Progress file cleared")
        except Exception as e:
            print(f"Error clearing progress: {e}")

    def create_chunks(self, items, chunk_size):
        """Split items into chunks of specified size"""
        chunks = []
        for i in range(0, len(items), chunk_size):
            chunk = items[i:i + chunk_size]
            chunks.append({
                'items': chunk,
                'start_index': i + 1,
                'end_index': min(i + chunk_size, len(items)),
                'size': len(chunk)
            })
        return chunks

    def process_chunk(self, chunk, chunk_num, total_chunks):
        """Process a single chunk of items (original manual version)"""
        print(f"\n{'='*70}")
        print(f"PROCESSING CHUNK {chunk_num} OF {total_chunks}")
        print(f"Items {chunk['start_index']}-{chunk['end_index']} ({chunk['size']} items)")
        print(f"{'='*70}")

        # Show first few items in this chunk
        print("Items in this chunk:")
        for i, item in enumerate(chunk['items'][:5], 1):
            print(f"  {i}. '{item}'")
        if len(chunk['items']) > 5:
            print(f"  ... and {len(chunk['items']) - 5} more")

        # Process each item in the chunk
        successful_selections = 0
        failed_selections = []
        consecutive_failures = 0
        max_consecutive_failures = 3

        for i, name in enumerate(chunk['items'], 1):
            print(f"\nSelecting {i}/{chunk['size']}: '{name}'")

            # Check for too many consecutive failures
            if consecutive_failures >= max_consecutive_failures:
                print(f"\nWARNING: {consecutive_failures} consecutive failures detected!")
                print("This might indicate a problem with the dropdown or remaining items.")
                print("\nOptions:")
                print("1. Continue trying")
                print("2. Skip remaining items in this chunk")
                print("3. Use 'Select All' to select everything remaining")

                failure_action = input("Select option (1, 2, or 3): ").strip()

                if failure_action == "2":
                    print("Skipping remaining items in this chunk...")
                    remaining_items = chunk['items'][i-1:]
                    failed_selections.extend(remaining_items)
                    break
                elif failure_action == "3":
                    print("Using 'Select All' to select remaining items...")
                    success = self.select_all_from_dropdown()
                    if success:
                        print("Select All completed! Assuming remaining items were selected.")
                        successful_selections += len(chunk['items']) - i + 1
                        break
                    else:
                        print("Select All failed, continuing with individual selection...")

                # Reset consecutive failures counter if user chooses to continue
                consecutive_failures = 0

            success = self.click_dropdown_and_select(name)

            if success:
                successful_selections += 1
                consecutive_failures = 0 # Reset on success
            else:
                failed_selections.append(name)
                consecutive_failures += 1

            time.sleep(0.08) # Small delay between selections

        # Chunk summary
        print(f"\n{'='*50}")
        print(f"CHUNK {chunk_num} SUMMARY")
        print(f"{'='*50}")
        print(f"Successful selections: {successful_selections}/{chunk['size']}")
        if failed_selections:
            print(f"Failed selections: {len(failed_selections)}")
            for item in failed_selections[:3]:
                print(f"  - {item}")
            if len(failed_selections) > 3:
                print(f"  ... and {len(failed_selections) - 3} more")

        return successful_selections, failed_selections

    def process_chunk_with_auto_report(self, chunk, chunk_num, total_chunks):
        """
        Process a single chunk of items and automatically generate report.
        """
        print(f"\n{'='*70}")
        print(f"PROCESSING CHUNK {chunk_num} OF {total_chunks}")
        print(f"Items {chunk['start_index']}-{chunk['end_index']} ({chunk['size']} items)")
        print(f"{'='*70}")

        # Show first few items in this chunk
        print("Items in this chunk:")
        for i, item in enumerate(chunk['items'][:5], 1):
            print(f"  {i}. '{item}'")
        if len(chunk['items']) > 5:
            print(f"  ... and {len(chunk['items']) - 5} more")

        # Process each item in the chunk
        successful_selections = 0
        failed_selections = []

        for i, name in enumerate(chunk['items'], 1):
            print(f"\nSelecting {i}/{chunk['size']}: '{name}'")

            success = self.click_dropdown_and_select(name)

            if success:
                successful_selections += 1
            else:
                failed_selections.append(name)

            time.sleep(0.08) # Small delay between selections

        # Chunk selection summary
        print(f"\n{'='*50}")
        print(f"CHUNK {chunk_num} SELECTION SUMMARY")
        print(f"{'='*50}")
        print(f"Successful selections: {successful_selections}/{chunk['size']}")
        if failed_selections:
            print(f"Failed selections: {len(failed_selections)}")
            for item in failed_selections[:3]:
                print(f"  - {item}")
            if len(failed_selections) > 3:
                print(f"  ... and {len(failed_selections) - 3} more")

        # Auto-generate report for this chunk
        print(f"\nAuto-generating report for Chunk {chunk_num}...")
        report_success = self.generate_report_for_current_selections()

        return successful_selections, failed_selections, report_success
    def consolidate_csv_files(self, directory, downloaded_files_count):
        """Consolidates multiple CSV files from a directory into a single file with Region mapping."""
        print("\n" + "="*50)
        print("CONSOLIDATING CSV FILES WITH REGION MAPPING")
        print("="*50)

        # Get the list of all CSV files in the download directory
        csv_files = [f for f in os.listdir(directory) if f.endswith('.csv')]
        
        if not csv_files:
            print("No CSV files found to consolidate.")
            return

        print(f"Found {len(csv_files)} CSV files in '{directory}'.")
        
        # Check if the downloaded count matches the found count
        if len(csv_files) != downloaded_files_count:
            print(f"WARNING: Expected {downloaded_files_count} files, but found {len(csv_files)}. "
                  f"Some files may not have finished downloading. Proceeding anyway...")

        # Load the master list for region mapping
        master_file_path = NAMES_FILE  # Use the global NAMES_FILE variable
        print(f"Loading master list from: '{master_file_path}'")
        
        try:
            if master_file_path.endswith('.csv'):
                master_df = pd.read_csv(master_file_path, encoding='utf-8-sig')
            elif master_file_path.endswith(('.xlsx', '.xls')):
                master_df = pd.read_excel(master_file_path)
            else:
                print("Master file must be CSV or Excel format. Proceeding without region mapping.")
                master_df = None
        except Exception as e:
            print(f"Error loading master file: {e}. Proceeding without region mapping.")
            master_df = None

        # Create region lookup dictionary if master file loaded successfully
        region_lookup = {}
        if master_df is not None and 'Location Name' in master_df.columns and 'Region' in master_df.columns:
            print("Creating region lookup dictionary...")
            for _, row in master_df.iterrows():
                if pd.notna(row['Location Name']) and pd.notna(row['Region']):
                    location_name = self.normalize_text(str(row['Location Name']))
                    region = str(row['Region']).strip()
                    region_lookup[location_name] = region
            print(f"Region lookup created with {len(region_lookup)} entries")
        else:
            print("Master file doesn't have required columns 'Location Name' and 'Region'. Proceeding without region mapping.")

        # Create a new directory for the merged file
        merged_dir = os.path.join(directory, "Merged_Reports")
        if not os.path.exists(merged_dir):
            os.makedirs(merged_dir)
            print(f"Created directory for merged reports: '{merged_dir}'")
        
        merged_file_path = os.path.join(merged_dir, f"Consolidated_MAR_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        
        # Initialize consolidated dataframe
        consolidated_df = pd.DataFrame()
        
        # Process each CSV file
        for i, filename in enumerate(csv_files):
            file_path = os.path.join(directory, filename)
            try:
                df = pd.read_csv(file_path)
                print(f"Processing file {i+1}/{len(csv_files)}: '{filename}' ({len(df)} rows)")
                
                # Add Region column if we have the lookup data and the file has 'Care Service' column
                if region_lookup and 'Care Service' in df.columns:
                    print("  Mapping regions based on Care Service...")
                    
                    # Create Region column by mapping Care Service to Location Name in master list
                    df['Region'] = df['Care Service'].apply(lambda x: 
                        region_lookup.get(self.normalize_text(str(x)) if pd.notna(x) else '', 'Unknown Region')
                    )
                    
                    # Count successful mappings
                    mapped_count = len(df[df['Region'] != 'Unknown Region'])
                    total_count = len(df)
                    print(f"  Successfully mapped {mapped_count}/{total_count} records to regions")
                    
                    if mapped_count < total_count:
                        unmapped_services = df[df['Region'] == 'Unknown Region']['Care Service'].unique()[:5]
                        print(f"  Sample unmapped Care Services: {list(unmapped_services)}")
                
                elif not region_lookup:
                    print("  No region lookup available - adding placeholder Region column")
                    df['Region'] = 'No Region Data'
                elif 'Care Service' not in df.columns:
                    print(f"  WARNING: 'Care Service' column not found in {filename}. Available columns: {list(df.columns)}")
                    df['Region'] = 'Missing Care Service Column'
                
                # Append to consolidated dataframe
                if consolidated_df.empty:
                    consolidated_df = df
                else:
                    # Use concat instead of append (which is deprecated)
                    consolidated_df = pd.concat([consolidated_df, df], ignore_index=True)
                    
            except Exception as e:
                print(f"  Error processing '{filename}': {e}. Skipping this file.")
        
        # Save the consolidated file
        try:
            consolidated_df.to_csv(merged_file_path, index=False)
            print(f"\nConsolidation complete!")
            print(f"Final merged report saved to: '{merged_file_path}'")
            print(f"Total consolidated records: {len(consolidated_df)}")
            
            # Show region distribution if regions were mapped
            if 'Region' in consolidated_df.columns and region_lookup:
                print(f"\nRegion distribution in consolidated file:")
                region_counts = consolidated_df['Region'].value_counts()
                for region, count in region_counts.head(10).items():
                    print(f"  {region}: {count} records")
                if len(region_counts) > 10:
                    print(f"  ... and {len(region_counts) - 10} more regions")
                    
        except Exception as e:
            print(f"Error saving consolidated file: {e}")

    def process_in_chunks(self, names_file, column_name="Location Name"):
        """Main function for chunked processing (manual report generation)"""

        # Check for existing progress
        existing_progress = self.load_chunking_progress("chunking_progress.json")
        resume_session = False

        if existing_progress:
            print(f"\n{'='*60}")
            print("PREVIOUS CHUNKING SESSION DETECTED")
            print(f"{'='*60}")
            print(f"File: {existing_progress['file_path']}")
            print(f"Total items: {existing_progress['total_items']}")
            print(f"Previous chunk size: {existing_progress['chunk_size']}")
            print(f"Total chunks: {existing_progress['total_chunks']}")
            print(f"Completed: {existing_progress['current_chunk'] - 1}/{existing_progress['total_chunks']}")
            if existing_progress.get('region_filter'):
                print(f"Region filter: {existing_progress['region_filter']}")
            print(f"Next chunk: {existing_progress['current_chunk']}")

            print(f"\nOptions:")
            print(f"1. Resume previous session (chunk size {existing_progress['chunk_size']})")
            print(f"2. Start new session (choose new chunk size)")
            print(f"3. Cancel")

            resume_choice = input("Select option (1, 2, or 3): ").strip()
            if resume_choice == "1":
                resume_session = True
                print("Resuming previous session...")
            elif resume_choice == "2":
                print("Starting new chunking session...")
                self.clear_chunking_progress("chunking_progress.json")
            else:
                print("Cancelled.")
                return

        if not resume_session:
            # Load names from file
            all_names, df = self.load_names_from_file(names_file, column_name)

            # Determine which column was actually used
            if column_name in df.columns:
                actual_column = column_name
            else:
                actual_column = df.columns[1]

            print(f"Loaded {len(all_names)} total names from file")

            # Region filtering
            print(f"\n{'='*50}")
            print("REGION FILTER FOR CHUNKING")
            print(f"{'='*50}")
            print("1. All regions (no filter)")
            print("2. Filter by specific region")

            filter_choice = input("Select option (1 or 2): ").strip()
            region_filter = None

            if filter_choice == "2":
                if "Region" not in df.columns:
                    print("No 'Region' column found in the file. Using all locations.")
                    names = all_names
                else:
                    regions = df["Region"].dropna().unique().tolist()
                    regions.sort()

                    print(f"\nAvailable regions:")
                    for i, region in enumerate(regions, 1):
                        count = len(df[df["Region"] == region])
                        print(f"  {i}. {region} ({count} locations)")

                    try:
                        region_choice = int(input(f"Select region number (1-{len(regions)}): ").strip())
                        if 1 <= region_choice <= len(regions):
                            selected_region = regions[region_choice - 1]
                            region_filter = selected_region
                            filtered_df = df[df["Region"] == selected_region]
                            names = filtered_df[actual_column].dropna().tolist()
                            names = [self.normalize_text(str(name)) for name in names]
                            print(f"Filtered to {len(names)} locations in {selected_region}")
                        else:
                            print("Invalid number. Using all locations.")
                            names = all_names
                    except ValueError:
                        print("Invalid input. Using all locations.")
                        names = all_names
            else:
                names = all_names

            # Chunk size configuration
            print(f"\n{'='*50}")
            print("CHUNK SIZE CONFIGURATION")
            print(f"{'='*50}")
            print(f"Total items to process: {len(names)}")

            default_chunk_size = 50
            chunk_input = input(f"Enter chunk size (default {default_chunk_size}): ").strip()

            try:
                chunk_size = int(chunk_input) if chunk_input else default_chunk_size
                if chunk_size <= 0:
                    chunk_size = default_chunk_size
                    print(f"Invalid chunk size. Using default: {default_chunk_size}")
            except ValueError:
                chunk_size = default_chunk_size
                print(f"Invalid input. Using default chunk size: {default_chunk_size}")

            # Create chunks
            chunks = self.create_chunks(names, chunk_size)

            print(f"\n{'='*50}")
            print("CHUNKING PLAN")
            print(f"{'='*50}")
            print(f"Total items: {len(names)}")
            print(f"Chunk size: {chunk_size}")
            print(f"Total chunks: {len(chunks)}")
            print(f"This will create {len(chunks)} separate reports")

            for i, chunk in enumerate(chunks, 1):
                print(f"  Chunk {i}: Items {chunk['start_index']}-{chunk['end_index']} ({chunk['size']} items)")
                # Show first few items in each chunk for verification
                if len(chunk['items']) > 0:
                    print(f"    First item: '{chunk['items'][0]}'")
                    if len(chunk['items']) > 1:
                        print(f"    Last item: '{chunk['items'][-1]}'")

            proceed = input(f"\nProceed with chunked processing? (y/n): ").strip().lower()
            if proceed != 'y':
                print("Chunking cancelled.")
                return

            # Initialize progress tracking
            progress_data = {
                'file_path': names_file,
                'column_name': column_name,
                'total_items': len(names),
                'chunk_size': chunk_size,
                'total_chunks': len(chunks),
                'current_chunk': 1,
                'region_filter': region_filter,
                'chunks': chunks,
                'names': names,
                'started_at': datetime.now().isoformat()
            }

            self.save_chunking_progress(progress_data)

        else:
            # Resume existing session
            progress_data = existing_progress
            chunks = progress_data['chunks']
            names = progress_data['names']
            chunk_size = progress_data['chunk_size']

            print(f"\nResuming with {len(names)} items, chunk size {chunk_size}")
        # Process chunks starting from current position
        start_chunk = progress_data['current_chunk']
        total_chunks = len(chunks)

        print(f"\n{'='*70}")
        print(f"STARTING CHUNK PROCESSING")
        print(f"Starting from chunk: {start_chunk}")
        print(f"Total chunks: {total_chunks}")
        print(f"{'='*70}")

        for chunk_num in range(start_chunk, total_chunks + 1):
            chunk = chunks[chunk_num - 1] # Convert to 0-based index

            print(f"\n{'='*70}")
            print(f"STARTING CHUNK {chunk_num} OF {total_chunks}")
            print(f"Items {chunk['start_index']}-{chunk['end_index']} ({chunk['size']} items)")
            print(f"{'='*70}")

            # Debug: Show what items are in this chunk
            print("DEBUG - Items in this chunk:")
            for i, item in enumerate(chunk['items'][:3]):
                print(f"  {i+1}. '{item}'")
            if len(chunk['items']) > 3:
                print(f"  ... and {len(chunk['items']) - 3} more")

            # Process the chunk
            successful, failed = self.process_chunk(chunk, chunk_num, total_chunks)

            # Update progress
            progress_data['current_chunk'] = chunk_num + 1
            self.save_chunking_progress(progress_data)

            # Pause for report generation
            print(f"\n{'='*70}")
            print(f"CHUNK {chunk_num} COMPLETE - GENERATE REPORT NOW")
            print(f"{'='*70}")
            print("Your selections are ready!")
            print("Please now:")
            print("1. Click the 'Generate Report' button")
            print("2. Wait for the report to generate")
            print("3. Download the report")
            print("4. Come back here when ready")
            print(f"\nProgress: {chunk_num}/{total_chunks} chunks completed")
            if chunk_num < total_chunks:
                next_chunk = chunks[chunk_num] # Next chunk (0-based)
                print(f"Next: Chunk {chunk_num + 1} (Items {next_chunk['start_index']}-{next_chunk['end_index']})")

            input(f"Press Enter after downloading Chunk {chunk_num} report...")

            if chunk_num < total_chunks:
                print(f"Moving to next chunk...")
            else:
                print("All chunks completed!")

        # Final summary
        print(f"\n{'='*70}")
        print("CHUNKING COMPLETE!")
        print(f"{'='*70}")
        print(f"Total chunks processed: {total_chunks}")
        print(f"Total items processed: {len(names)}")
        print(f"You should now have {total_chunks} separate report files")

        # Clear progress file
        self.clear_chunking_progress("chunking_progress.json")

        # Next actions menu
        print(f"\nWhat would you like to do next?")
        print("1. Start new chunking session")
        print("2. Return to main menu")
        print("3. End script and close browser")

        next_action = input("Select option (1, 2, or 3): ").strip()

        if next_action == "1":
            print("\nStarting new chunking session...")
            self.process_in_chunks(names_file, column_name)
        elif next_action == "2":
            print("\nReturning to main menu...")
            self.select_multiple_items_from_current_page(names_file, column_name)
        else:
            print("\nEnding script...")
            input("Press Enter to close browser...")
    def process_in_chunks_with_auto_reports(self, names_file, column_name="Location Name", download_directory=None):
        """
        Main function for chunked processing with automatic report generation.
        """

        # Check for existing progress
        existing_progress = self.load_chunking_progress("chunking_progress.json")
        resume_session = False

        if existing_progress:
            print(f"\n{'='*60}")
            print("PREVIOUS CHUNKING SESSION DETECTED")
            print(f"{'='*60}")
            print(f"File: {existing_progress['file_path']}")
            print(f"Total items: {existing_progress['total_items']}")
            print(f"Previous chunk size: {existing_progress['chunk_size']}")
            print(f"Total chunks: {existing_progress['total_chunks']}")
            print(f"Completed: {existing_progress['current_chunk'] - 1}/{existing_progress['total_chunks']}")
            if existing_progress.get('region_filter'):
                print(f"Region filter: {existing_progress['region_filter']}")
            print(f"Next chunk: {existing_progress['current_chunk']}")

            print(f"\nOptions:")
            print(f"1. Resume previous session (chunk size {existing_progress['chunk_size']})")
            print(f"2. Start new session (choose new chunk size)")
            print(f"3. Cancel")

            resume_choice = input("Select option (1, 2, or 3): ").strip()
            if resume_choice == "1":
                resume_session = True
                print("Resuming previous session...")
            elif resume_choice == "2":
                print("Starting new chunking session...")
                self.clear_chunking_progress("chunking_progress.json")
            else:
                print("Cancelled.")
                return

        if not resume_session:
            # Load names from file
            all_names, df = self.load_names_from_file(names_file, column_name)

            # Determine which column was actually used
            if column_name in df.columns:
                actual_column = column_name
            else:
                actual_column = df.columns[1]

            print(f"Loaded {len(all_names)} total names from file")

            # Region filtering (keep existing logic)
            print(f"\n{'='*50}")
            print("REGION FILTER FOR CHUNKING")
            print(f"{'='*50}")
            print("1. All regions (no filter)")
            print("2. Filter by specific region")

            filter_choice = input("Select option (1 or 2): ").strip()
            region_filter = None

            if filter_choice == "2":
                if "Region" not in df.columns:
                    print("No 'Region' column found in the file. Using all locations.")
                    names = all_names
                else:
                    regions = df["Region"].dropna().unique().tolist()
                    regions.sort()

                    print(f"\nAvailable regions:")
                    for i, region in enumerate(regions, 1):
                        count = len(df[df["Region"] == region])
                        print(f"  {i}. {region} ({count} locations)")

                    try:
                        region_choice = int(input(f"Select region number (1-{len(regions)}): ").strip())
                        if 1 <= region_choice <= len(regions):
                            selected_region = regions[region_choice - 1]
                            region_filter = selected_region
                            filtered_df = df[df["Region"] == selected_region]
                            names = filtered_df[actual_column].dropna().tolist()
                            names = [self.normalize_text(str(name)) for name in names]
                            print(f"Filtered to {len(names)} locations in {selected_region}")
                        else:
                            print("Invalid number. Using all locations.")
                            names = all_names
                    except ValueError:
                        print("Invalid input. Using all locations.")
                        names = all_names
            else:
                names = all_names

            # Chunk size configuration
            print(f"\n{'='*50}")
            print("CHUNK SIZE CONFIGURATION")
            print(f"{'='*50}")
            print(f"Total items to process: {len(names)}")

            default_chunk_size = 50
            chunk_input = input(f"Enter chunk size (default {default_chunk_size}): ").strip()

            try:
                chunk_size = int(chunk_input) if chunk_input else default_chunk_size
                if chunk_size <= 0:
                    chunk_size = default_chunk_size
                    print(f"Invalid chunk size. Using default: {default_chunk_size}")
            except ValueError:
                chunk_size = default_chunk_size
                print(f"Invalid input. Using default chunk size: {default_chunk_size}")

            # Create chunks
            chunks = self.create_chunks(names, chunk_size)

            print(f"\n{'='*50}")
            print("AUTOMATED CHUNKING PLAN")
            print(f"{'='*50}")
            print(f"Total items: {len(names)}")
            print(f"Chunk size: {chunk_size}")
            print(f"Total chunks: {len(chunks)}")
            print(f"This will create {len(chunks)} separate reports AUTOMATICALLY")
            print("Reports will be generated and downloaded without manual intervention")

            for i, chunk in enumerate(chunks, 1):
                print(f"  Chunk {i}: Items {chunk['start_index']}-{chunk['end_index']} ({chunk['size']} items)")

            proceed = input(f"\nProceed with AUTOMATED chunked processing? (y/n): ").strip().lower()
            if proceed != 'y':
                print("Chunking cancelled.")
                return

            # Initialize progress tracking
            progress_data = {
                'file_path': names_file,
                'column_name': column_name,
                'total_items': len(names),
                'chunk_size': chunk_size,
                'total_chunks': len(chunks),
                'current_chunk': 1,
                'region_filter': region_filter,
                'chunks': chunks,
                'names': names,
                'started_at': datetime.now().isoformat()
            }

            self.save_chunking_progress(progress_data)

        else:
            # Resume existing session
            progress_data = existing_progress
            chunks = progress_data['chunks']
            names = progress_data['names']
            chunk_size = progress_data['chunk_size']

            print(f"\nResuming with {len(names)} items, chunk size {chunk_size}")
            # Process chunks starting from current position
        start_chunk = progress_data['current_chunk']
        total_chunks = len(chunks)
        successful_reports = 0
        failed_reports = 0

        print(f"\n{'='*70}")
        print(f"STARTING AUTOMATED CHUNK PROCESSING")
        print(f"Starting from chunk: {start_chunk}")
        print(f"Total chunks: {total_chunks}")
        print(f"Reports will be generated automatically for each chunk!")
        print(f"{'='*70}")

        for chunk_num in range(start_chunk, total_chunks + 1):
            chunk = chunks[chunk_num - 1] # Convert to 0-based index

            # Process the chunk with automatic report generation
            successful, failed, report_success = self.process_chunk_with_auto_report(chunk, chunk_num, total_chunks)

            # Track report success
            if report_success:
                successful_reports += 1
            else:
                failed_reports += 1

            # Update progress
            progress_data['current_chunk'] = chunk_num + 1
            self.save_chunking_progress(progress_data)

            print(f"\n{'='*70}")
            print(f"CHUNK {chunk_num} COMPLETE")
            print(f"Progress: {chunk_num}/{total_chunks} chunks completed")
            print(f"Reports generated: {successful_reports}")
            print(f"Report failures: {failed_reports}")
            print(f"{'='*70}")

            # Brief pause between chunks
            if chunk_num < total_chunks:
                print("Brief pause before next chunk...")
                time.sleep(0.08)

        # Final summary
        print(f"\n{'='*70}")
        print("AUTOMATED CHUNKING COMPLETE!")
        print(f"{'='*70}")
        print(f"Total chunks processed: {total_chunks}")
        print(f"Total items processed: {len(names)}")
        print(f"Successful reports: {successful_reports}")
        print(f"Failed reports: {failed_reports}")

        # Add the consolidation logic here
        if download_directory:
            print("\nWaiting for a few seconds to ensure all files are downloaded before consolidation...")
            time.sleep(5)
            self.consolidate_csv_files(download_directory, successful_reports)

        print(f"You should now have a consolidated report and {successful_reports} report files downloaded")

        # Clear progress file
        self.clear_chunking_progress("chunking_progress.json")

        # Next actions menu
        print(f"\nWhat would you like to do next?")
        print("1. Start new automated chunking session")
        print("2. Return to main menu")
        print("3. End script and close browser")

        next_action = input("Select option (1, 2, or 3): ").strip()

        if next_action == "1":
            print("\nStarting new automated chunking session...")
            self.process_in_chunks_with_auto_reports(names_file, column_name, download_directory)
        elif next_action == "2":
            print("\nReturning to main menu...")
            self.select_multiple_items_from_current_page(names_file, column_name)
        else:
            print("\nEnding script...")
            input("Press Enter to close browser...")

    def wait_for_manual_setup(self, target_url):
        """Wait for user to manually login and navigate to the correct page"""
        print("=== MANUAL SETUP REQUIRED ===")
        print("1. Browser should have opened...")
        print(f"2. Please login and navigate to: {target_url}")
        print("3. Come back here and press Enter when ready")
        print("=====================================")

        try:
            print("Opening login page...")
            self.driver.get(target_url)
            print("Page loaded. Please complete login manually.")
            input("\nPress Enter when you're logged in and on the correct page with the dropdown...")

            # Verify we're on the right page (optional)
            current_url = self.driver.current_url
            print(f"Current URL: {current_url}")
            if "reports/mar" in current_url:
                print("Detected you're on the MAR reports page")
            else:
                print("Make sure you're on the correct page with the dropdown")

        except Exception as e:
            print(f"Error during setup: {e}")
            raise

    def select_multiple_items_from_current_page(self, names_file, column_name="Location Name"):
        """Enhanced main function with automated report generation option"""

        # Load names from file
        all_names, df = self.load_names_from_file(names_file, column_name)

        # Determine which column was actually used
        if column_name in df.columns:
            actual_column = column_name
        else:
            actual_column = df.columns[1]

        print(f"Loaded {len(all_names)} total names from file")

        # Enhanced main menu with automation option
        print("\n" + "="*50)
        print("PROCESSING OPTIONS")
        print("="*50)
        print("1. All regions (no filter)")
        print("2. Filter by specific region")
        print("3. Select ALL items at once (use dropdown 'Select All')")
        print("4. Process in chunks (manual report generation)")
        print("5. Process in chunks with AUTOMATED report generation")

        filter_choice = input("\nSelect option (1, 2, 3, 4, or 5): ").strip()

        if filter_choice == "5":
            # Use automated chunking functionality
            print("\nSwitching to AUTOMATED chunked processing mode...")
            self.process_in_chunks_with_auto_reports(names_file, column_name, DOWNLOAD_DIRECTORY)
            return
        elif filter_choice == "4":
            # Use manual chunking functionality
            print("\nSwitching to manual chunked processing mode...")
            self.process_in_chunks(names_file, column_name)
            return
        elif filter_choice == "2":
            # Check if Region column exists
            if "Region" not in df.columns:
                print("No 'Region' column found in the file. Using all locations.")
                names = all_names
            else:
                # Show available regions with numbers
                regions = df["Region"].dropna().unique().tolist()
                regions.sort()

                print(f"\nAvailable regions:")
                for i, region in enumerate(regions, 1):
                    count = len(df[df["Region"] == region])
                    print(f"  {i}. {region} ({count} locations)")

                try:
                    region_choice = int(input(f"Select region number (1-{len(regions)}): ").strip())
                    if 1 <= region_choice <= len(regions):
                        selected_region = regions[region_choice - 1]
                        # Filter the DataFrame by region
                        filtered_df = df[df["Region"] == selected_region]
                        # Get the names from filtered data using the correct column
                        names = filtered_df[actual_column].dropna().tolist()
                        # Normalize the names
                        names = [self.normalize_text(str(name)) for name in names]
                        print(f"Filtered to {len(names)} locations in {selected_region}")
                    else:
                        print(f"Invalid number. Using all locations.")
                        names = all_names
                except ValueError:
                    print(f"Please enter a valid number. Using all locations.")
                    names = all_names
        elif filter_choice == "3":
            # Use dropdown Select All functionality
            print("Using dropdown 'Select All' to select all items...")
            success = self.select_all_from_dropdown()
            if success:
                print("All items selected successfully using dropdown 'Select All'!")

                # Offer automated report generation
                print(f"\n" + "="*50)
                print("REPORT GENERATION OPTIONS")
                print("="*50)
                print("1. Generate report manually")
                print("2. Generate report automatically")

                report_choice = input("Select option (1 or 2): ").strip()

                if report_choice == "2":
                    print("\nGenerating report automatically...")
                    success = self.generate_report_for_current_selections()
                    if success:
                        print("Report generated and downloaded successfully!")
                    else:
                        print("Report generation failed. You may need to generate manually.")
                else:
                    print("You can now generate your report manually.")

                input("Press Enter after you've finished with the report...")

                return
            else:
                print("Select All failed. Falling back to individual selection method.")
                names = all_names
        else:
            # No filter - use all names
            print("Using all locations")
            names = all_names

        print(f"Final list: {len(names)} names to process")

        # Show first few names for verification
        print("First 5 names in your list:")
        for i, name in enumerate(names[:5]):
            print(f"  {i+1}. '{name}'")

        verify = input("Do these names look correct? (y/n): ")
        if verify.lower() != 'y':
            print("Please check your selection and restart the script")
            return
            print("\n" + "="*70)
        print("ENHANCED SELECTION MODE")
        print("Instructions:")
        print("  1. Make sure you're on the page with the dropdown")
        print("  2. Script will automatically find and click each option")
        print("  3. Uses enhanced element detection for problematic items")
        print("  4. JavaScript click as primary method, standard click as fallback")
        print(f"  5. Processing {len(names)} selections automatically")
        print("="*70)

        proceed = input("Ready to start automated selections? (y/n): ")
        if proceed.lower() != 'y':
            print("Script cancelled.")
            return

        # Process each name using enhanced selection method
        successful_selections = 0
        failed_selections = []

        for i, name in enumerate(names, 1):
            print(f"\n{'='*50}")
            print(f"Processing {i}/{len(names)}: '{name}'")
            print(f"Remaining: {len(names) - i}")
            print('='*50)

            success = self.click_dropdown_and_select(name)

            if success:
                successful_selections += 1
            else:
                failed_selections.append(name)

            # Show running totals
            print(f"Running total: {successful_selections} success, {len(failed_selections)} failed")

            # Small delay between selections
            time.sleep(0.08)

        # Summary
        print(f"\n" + "="*60)
        print(f"SELECTION SUMMARY")
        print(f"="*60)
        print(f"Total processed: {len(names)}")
        print(f"Successful selections: {successful_selections}")
        print(f"Failed selections: {len(failed_selections)}")
        if len(names) > 0:
            print(f"Success rate: {(successful_selections/len(names)*100):.1f}%")

        if failed_selections:
            print(f"\nFailed items (first 10):")
            for j, item in enumerate(failed_selections[:10]):
                print(f"  {j+1}. {item}")

            if len(failed_selections) > 10:
                print(f"  ... and {len(failed_selections) - 10} more")

        print("="*60)
        print("Selection process complete!")

        # Report generation options
        print(f"\n" + "="*50)
        print("REPORT GENERATION & NEXT ACTIONS")
        print("="*50)
        print("Your selections are now complete!")
        print("1. Generate report manually")
        print("2. Generate report automatically")

        report_choice = input("Select option (1 or 2): ").strip()

        if report_choice == "2":
            print("\nGenerating report automatically...")
            success = self.generate_report_for_current_selections()
            if success:
                print("Report generated and downloaded successfully!")
            else:
                print("Report generation failed. You may need to generate manually.")
        else:
            print("You can now generate your report manually in the browser.")

        input("Press Enter after you've finished with the report...")

        print("\nWhat would you like to do next?")
        print("1. Start new report (keep current selections)")
        print("2. End script and close browser")

        next_action = input("Select option (1 or 2): ").strip()

        if next_action == "1":
            print("\nStarting new report (keeping current selections)...")
            self.select_multiple_items_from_current_page(names_file, column_name)

        else:
            print("\nEnding script...")
            print("Browser will remain open for final manual actions.")
            input("Press Enter to close browser...")
            return

    def close(self):
        """Close the browser"""
        self.driver.quit()
        # Example usage
if __name__ == "__main__":
    # CONFIGURATION - Easy to modify
    NAMES_FILE = r"C:\Users\MarkCooper-NCG\OneDrive - National Care Group\Documents\camascope\Camascope Query\full_checkbox_list_checkbox_ID.csv"
    TARGET_URL = "https://emar.vcaresystems.co.uk/#/app/reports/mar"
    
    # Define the download directory here
    DOWNLOAD_DIRECTORY = r"C:\temp\camascope\outputs"

    # --- NEW: DIRECTORY CHECK AND CLEANUP LOGIC ---
    print(f"\nChecking and cleaning output directory: '{DOWNLOAD_DIRECTORY}'")
    
    # Check if the directory exists, if not, create it
    if not os.path.exists(DOWNLOAD_DIRECTORY):
        os.makedirs(DOWNLOAD_DIRECTORY)
        print("Directory did not exist, so it was created.")
    else:
        # If the directory exists, clean it out
        print("Directory exists. Cleaning out old files...")
        for item in os.listdir(DOWNLOAD_DIRECTORY):
            item_path = os.path.join(DOWNLOAD_DIRECTORY, item)
            # Do NOT remove the 'Merged_Reports' folder
            if item == "Merged_Reports":
                print(f"  Skipping '{item_path}' as requested.")
                continue

            try:
                if os.path.isfile(item_path):
                    os.remove(item_path)
                    print(f"  Deleted file: {item}")
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                    print(f"  Deleted directory: {item}")
            except OSError as e:
                print(f"  Error deleting {item_path}: {e}")
        print("Directory cleanup complete.")
    # --- END OF NEW LOGIC ---

    # Initialize the automator and pass the download directory
    automator = FixedDropdownAutomator(download_path=DOWNLOAD_DIRECTORY)

    try:
        # Perform the automated login and navigation
        automator.automated_login(TARGET_URL)

        # Run the automation using enhanced selection method with automation options
        automator.select_multiple_items_from_current_page(NAMES_FILE, "Location Name")

    except Exception as e:
        print(f"Error: {str(e)}")
        traceback.print_exc()

    finally:
        # Close browser
        automator.close()
