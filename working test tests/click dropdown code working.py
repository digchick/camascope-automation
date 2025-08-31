from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import time
import sys
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("dropdown_clicker_log.txt", mode='w'),
        logging.StreamHandler(sys.stdout)
    ]
)

def setup_chrome_driver():
    """Setup Chrome WebDriver with options"""
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Remove headless mode so you can see the browser
    # chrome_options.add_argument("--headless")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        driver.maximize_window()
        return driver
    except Exception as e:
        logging.error(f"Error setting up Chrome driver: {e}")
        logging.error("Make sure ChromeDriver is installed and in PATH")
        return None

def automated_login_and_navigation(driver):
    """
    Automates the login process, navigates to the reports page, and enters dates.
    """
    logging.info("=== AUTOMATED LOGIN PROCESS ===")
    logging.info("Reading credentials from 'camascope login.txt'...")
    try:
        with open("camascope login.txt", "r") as f:
            lines = f.readlines()
            if len(lines) >= 2:
                plaintext_username = lines[0].strip()
                plaintext_password = lines[1].strip()
                logging.info("Credentials loaded successfully.")
            else:
                raise ValueError("Error: 'camascope login.txt' must contain username and password.")
    except FileNotFoundError:
        logging.error("Error: 'camascope login.txt' not found.")
        raise

    try:
        target_url = "https://emar.vcaresystems.co.uk"
        logging.info(f"Navigating to {target_url}...")
        driver.get(target_url)

        # --- Step 1: Click "Another User" Button ---
        logging.info("Navigating to the 'Another User' screen...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'otherUserDetailsBox')]"))
        ).click()
        logging.info("Clicked 'Another User'.")

        # --- Step 2: Enter Username and Sign In ---
        logging.info("Entering username...")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "signInName"))
        ).send_keys(plaintext_username)

        driver.find_element(By.ID, "continue").click()
        logging.info("Entered username and clicked 'Sign In'.")

        # --- Step 3: Enter Password and Sign In ---
        logging.info("Entering password...")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "password"))
        ).send_keys(plaintext_password)

        driver.find_element(By.ID, "continue").click()
        logging.info("Entered password and clicked 'Sign In'.")

        # --- MANUAL INTERVENTION STEP ---
        logging.info("\n--- MANUAL ACTION REQUIRED ---")
        input("Login complete. Please manually clear any pop-ups and navigate to the MAR Report page. Press Enter when ready to start the dropdown test.")
        logging.info("Manual action confirmed. Resuming script.")

    except Exception as e:
        logging.error(f"An error occurred during login or navigation: {e}")
        raise

def click_dropdown_repeatedly(driver, clicks=10):
    """Click dropdown repeatedly with enhanced debugging"""
    try:
        logging.info("Waiting 5 seconds for page to fully stabilize...")
        time.sleep(5)
        
        # Enhanced debugging - let's see what we're working with
        logging.info("DEBUGGING: Analyzing page elements...")
        
        # Check all possible dropdown elements
        all_divs = driver.find_elements(By.CSS_SELECTOR, "div[class*='css-']")
        logging.info(f"Found {len(all_divs)} divs with CSS classes")
        
        # Look for the specific elements with visibility check
        svg_arrows = driver.find_elements(By.CSS_SELECTOR, "svg.css-19bqh2r")
        logging.info(f"Found {len(svg_arrows)} total SVG arrows")
        
        # Filter for visible SVG arrows
        visible_svg_arrows = [svg for svg in svg_arrows if svg.is_displayed() and svg.size['height'] > 0 and svg.size['width'] > 0]
        logging.info(f"Found {len(visible_svg_arrows)} visible SVG arrows")
        
        containers = driver.find_elements(By.CSS_SELECTOR, "div.css-1gtu0rj-indicatorContainer")
        logging.info(f"Found {len(containers)} indicator containers")
        
        controls = driver.find_elements(By.CSS_SELECTOR, "div.css-1pahdxg-control")
        logging.info(f"Found {len(controls)} control elements")
        
        # Also try broader selectors for dropdowns
        react_selects = driver.find_elements(By.CSS_SELECTOR, "[class*='css-'][class*='container']")
        logging.info(f"Found {len(react_selects)} potential React Select containers")
        
        # Try to find any clickable dropdown indicator
        clickable_indicators = driver.find_elements(By.CSS_SELECTOR, "[class*='indicator']")
        logging.info(f"Found {len(clickable_indicators)} potential indicators")
        
        # Test manual click to verify functionality
        target_element = None
        target_description = ""
        
        # Priority 1: Visible SVG arrows
        if visible_svg_arrows:
            target_element = visible_svg_arrows[0]
            target_description = "Visible SVG arrow"
        # Priority 2: Any visible indicator containers
        elif containers:
            visible_containers = [c for c in containers if c.is_displayed()]
            if visible_containers:
                target_element = visible_containers[0]
                target_description = "Indicator container"
        # Priority 3: Control elements
        elif controls:
            visible_controls = [c for c in controls if c.is_displayed()]
            if visible_controls:
                target_element = visible_controls[0]
                target_description = "Control element"
        # Priority 4: Any clickable indicator
        elif clickable_indicators:
            visible_clickable = [c for c in clickable_indicators if c.is_displayed() and c.is_enabled()]
            if visible_clickable:
                target_element = visible_clickable[0]
                target_description = "Clickable indicator"
        
        if not target_element:
            logging.error("No suitable visible dropdown elements found!")
            # Debug: Show what we found
            for i, svg in enumerate(svg_arrows[:3]):  # Show first 3
                logging.info(f"SVG {i}: displayed={svg.is_displayed()}, size={svg.size}, location={svg.location}")
            return
        
        logging.info(f"Using {target_description} as target element")
        
        # Test click
        logging.info("Testing manual click on dropdown element...")
        logging.info(f"Element location: {target_element.location}")
        logging.info(f"Element size: {target_element.size}")
        logging.info(f"Element displayed: {target_element.is_displayed()}")
        logging.info(f"Element enabled: {target_element.is_enabled()}")
        
        # Try ActionChains click
        logging.info("Testing ActionChains click...")
        actions = ActionChains(driver)
        actions.move_to_element(target_element).click().perform()
        time.sleep(2)
        
        # Check if menu appeared
        menus = driver.find_elements(By.CSS_SELECTOR, "div.css-26l3qy-menu")
        logging.info(f"Menu elements found after test click: {len(menus)}")
        
        if menus:
            logging.info("SUCCESS: Dropdown opened! Closing it...")
            # Click again to close
            actions.move_to_element(target_element).click().perform()
            time.sleep(1)
        
        logging.info("\n" + "="*30)
        input("Press ENTER to continue with automated clicking, or CTRL+C to stop...")
        
        # Start the clicking sequence
        logging.info(f"Starting to click dropdown {clicks} times...")
        
        for i in range(clicks):
            try:
                # Try to find the best target element for each click
                current_target = None
                target_type = ""
                
                # Method 1: Look for visible SVG arrows first
                svg_elements = driver.find_elements(By.CSS_SELECTOR, "svg.css-19bqh2r")
                for svg in svg_elements:
                    if svg.is_displayed() and svg.size['height'] > 0 and svg.size['width'] > 0:
                        current_target = svg
                        target_type = "visible SVG arrow"
                        break
                
                # Method 2: Look for indicator containers
                if not current_target:
                    containers = driver.find_elements(By.CSS_SELECTOR, "div.css-1gtu0rj-indicatorContainer")
                    for container in containers:
                        if container.is_displayed() and container.size['height'] > 0:
                            current_target = container
                            target_type = "indicator container"
                            break
                
                # Method 3: Look for control elements
                if not current_target:
                    controls = driver.find_elements(By.CSS_SELECTOR, "div.css-1pahdxg-control")
                    for control in controls:
                        if control.is_displayed() and control.size['height'] > 0:
                            current_target = control
                            target_type = "control element"
                            break
                
                # Method 4: Try broader selectors
                if not current_target:
                    broad_selectors = [
                        "[class*='indicator'][class*='Container']",
                        "[class*='dropdown']",
                        "[role='button']",
                        "div[class*='css-'][class*='control']"
                    ]
                    for selector in broad_selectors:
                        elements = driver.find_elements(By.CSS_SELECTOR, selector)
                        for elem in elements:
                            if elem.is_displayed() and elem.size['height'] > 0:
                                current_target = elem
                                target_type = f"element found with {selector}"
                                break
                        if current_target:
                            break
                
                if not current_target:
                    logging.error(f"Click {i+1}: Could not find any suitable target element")
                    continue
                
                logging.info(f"Click {i+1}: Using {target_type}")
                logging.info(f"Click {i+1}: Element location: {current_target.location}, size: {current_target.size}")
                
                # Scroll element into view
                driver.execute_script("arguments[0].scrollIntoView(true);", current_target)
                time.sleep(0.3)
                
                # Try multiple click methods
                click_success = False
                
                # Method 1: ActionChains
                try:
                    actions = ActionChains(driver)
                    actions.move_to_element(current_target).click().perform()
                    logging.info(f"Click {i+1}: ActionChains click performed")
                    click_success = True
                except Exception as e:
                    logging.warning(f"Click {i+1}: ActionChains failed: {e}")
                
                # Method 2: JavaScript click if ActionChains failed
                if not click_success:
                    try:
                        driver.execute_script("arguments[0].click();", current_target)
                        logging.info(f"Click {i+1}: JavaScript click performed")
                        click_success = True
                    except Exception as e:
                        logging.warning(f"Click {i+1}: JavaScript click failed: {e}")
                
                # Method 3: Regular Selenium click as last resort
                if not click_success:
                    try:
                        current_target.click()
                        logging.info(f"Click {i+1}: Regular click performed")
                        click_success = True
                    except Exception as e:
                        logging.warning(f"Click {i+1}: Regular click failed: {e}")
                
                if not click_success:
                    logging.error(f"Click {i+1}: All click methods failed!")
                    continue
                
                # Wait and check if menu appeared
                time.sleep(0.8)
                try:
                    WebDriverWait(driver, 2).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.css-26l3qy-menu"))
                    )
                    logging.info(f"Click {i+1}: Menu opened successfully")
                    
                    # Close the dropdown
                    actions = ActionChains(driver)
                    actions.move_to_element(current_target).click().perform()
                    logging.info(f"Click {i+1}: Dropdown closed")
                    
                except:
                    logging.warning(f"Click {i+1}: Menu may not have opened")
                
                # Wait before next click
                time.sleep(1.2)
                
            except Exception as e:
                logging.error(f"Error on click {i+1}: {e}")
                time.sleep(1)
                continue
        
        logging.info(f"Completed {clicks} dropdown clicks!")
        
    except Exception as e:
        logging.error(f"Error during dropdown clicking: {e}")

def main():
    logging.info("="*60)
    logging.info("STARTING AUTOMATED DROPDOWN CLICKER")
    logging.info("="*60)
    
    driver = setup_chrome_driver()
    
    if not driver:
        logging.error("Failed to setup Chrome driver. Exiting.")
        sys.exit(1)
    
    try:
        # Step 1: Automated login
        automated_login_and_navigation(driver)
        
        # Step 2: Click dropdown repeatedly
        click_dropdown_repeatedly(driver, clicks=10)
        
        # Keep browser open longer to copy console output
        logging.info("\n" + "="*50)
        logging.info("SCRIPT COMPLETED")
        logging.info("="*50)
        logging.info("You now have time to copy the console output.")
        logging.info("The browser will remain open for 60 seconds.")
        logging.info("Press CTRL+C to close immediately, or wait for auto-close.")
        logging.info("="*50)
        
        for i in range(60, 0, -5):
            logging.info(f"Browser closing in {i} seconds...")
            time.sleep(5)
        
    except KeyboardInterrupt:
        logging.info("\nScript interrupted by user")
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
    finally:
        logging.info("Closing browser...")
        driver.quit()
        logging.info("Script finished.")

if __name__ == "__main__":
    main()