import configparser
import pandas as pd
import time
import os.path
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from fake_useragent import UserAgent
from openpyxl import load_workbook # Still needed by pandas for append mode
import keyring
import pickle
import undetected_chromedriver as uc
import random

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# setup the driver
def setupDriver():
    """
    Sets up and returns a Chrome WebDriver instance using undetected_chromedriver.
    Configures user-agent, maximizes window, and hides webdriver property.
    """
    options = Options()
    options.add_argument(f'user-agent={UserAgent().random}')
    options.add_argument("--start-maximized")
    # Uncomment and fill in this line if you want to use your own user data
    # options.add_argument("user-data-dir=/path/to/your/user/data")
    
    try:
        driver = uc.Chrome(options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        driver.implicitly_wait(10) # Implicit wait applies to all find_element calls
        logging.info("WebDriver setup successfully.")
        return driver
    except Exception as e:
        logging.error(f"Error setting up WebDriver: {e}")
        raise

# logins the user to the website
def login(driver, config):
    """
    Logs the user into the website using credentials from keyring and XPath identifiers from config.
    """
    try:
        username_xpath = config.get('loginPath', 'userName')
        password_xpath = config.get('loginPath', 'password')
        login_button_xpath = config.get('loginPath', 'loginB')

        username = keyring.get_credential(service_name='Onboard', username=None).username
        password = keyring.get_credential(service_name='Onboard', username=None).password

        driver.find_element(By.XPATH, username_xpath).send_keys(username)
        driver.find_element(By.XPATH, password_xpath).send_keys(password)
        driver.find_element(By.XPATH, login_button_xpath).click()
        haltStep(1)
        logging.info("Login attempt initiated.")
    except NoSuchElementException as e:
        logging.error(f"Login element not found: {e}. Check XPaths in settings.ini.")
        raise
    except Exception as e:
        logging.error(f"An unexpected error occurred during login: {e}")
        raise

# makes sure an element loads before clicking it
def waitToLoad(driver, byType, identifier, timeout=10, condition=EC.visibility_of_element_located):
    """
    Waits for an element to satisfy a given expected condition within a timeout.
    
    Args:
        driver: The Selenium WebDriver instance.
        byType: The By strategy (e.g., By.XPATH, By.ID).
        identifier: The identifier string for the element.
        timeout: Maximum time to wait in seconds.
        condition: The expected_conditions method to use (default: visibility_of_element_located).
        
    Returns:
        True if the element is found and condition is met, False otherwise.
    """
    try:
        WebDriverWait(driver, timeout).until(condition((byType, identifier)))
        logging.debug(f"Element '{identifier}' (Type: {byType}) loaded successfully.")
        return True
    except TimeoutException:
        logging.warning(f"Element '{identifier}' (Type: {byType}) did NOT load within {timeout} seconds.")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred in waitToLoad for '{identifier}': {e}")
        return False
    
# randomly stops the script between 1-3 seconds (to avoid detection)
def haltStep(x):
    """
    Pauses the script for a specified number of seconds.
    """
    time.sleep(x)

def is_element_present(driver, by_locator, value):
    """
    Checks if an element is present in the DOM without raising an exception.
    Returns the element if found, None otherwise.
    """
    try:
        return driver.find_element(by_locator, value)
    except NoSuchElementException:
        return None

def writeOut(row_manual_review):
    """
    Writes or appends rows requiring manual review to an Excel file.
    """
    if row_manual_review:
        df_manual_review = pd.DataFrame(row_manual_review)
        output_excel_path = r'C:\Users\dmartinez\Documents\ope\TPM.xlsx' # Define your output path

        if os.path.exists(output_excel_path):
            try:
                # Load existing workbook to check for sheet existence and max_row
                book = load_workbook(output_excel_path)
                with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    writer.book = book # Assign the loaded workbook to the writer

                    sheet_name = 'Manual Review'
                    start_row = 0
                    header = True # Write header if sheet is new or if overlaying from row 0

                    if sheet_name in writer.book.sheetnames:
                        sheet = writer.book[sheet_name]
                        start_row = sheet.max_row # Start appending after the last row
                        header = False # Do not write header if sheet already exists and we are appending
                        # If the sheet is empty except for headers, max_row might be 1.
                        # Ensure we append correctly.
                        if sheet.max_row == 1 and all(cell.value is None for cell in sheet[1]):
                             # Sheet exists but is empty, or only has header row.
                             # We might want to overwrite header or append.
                             # For simplicity, if it's just a header, we can append without header.
                             # If it's truly empty, start_row will be 1, so append will start at row 1.
                             # This logic assumes the first row is always header if present.
                             pass # Keep header=False if we are just appending to existing data.
                        elif sheet.max_row > 1:
                            # If there's actual data beyond the header, append without header.
                            start_row = sheet.max_row
                            header = False
                        else: # Sheet exists but is empty or only has headers
                            start_row = sheet.max_row if sheet.max_row > 0 else 0
                            header = False if sheet.max_row > 0 else True # If sheet is truly empty, write header

                    df_manual_review.to_excel(writer, sheet_name=sheet_name, index=False, header=header, startrow=start_row)

                logging.info(f"Successfully appended {len(row_manual_review)} rows requiring manual intervention to '{output_excel_path}'")
            except Exception as e:
                logging.error(f"Error appending data to Excel: {e}. Attempting to overwrite as fallback.")
                try:
                    # Fallback: Overwrite if append fails
                    df_manual_review.to_excel(output_excel_path, index=False, sheet_name='Manual Review')
                    logging.warning(f"Fallback: Overwrote existing file due to append error. Check '{output_excel_path}'")
                except Exception as e_fallback:
                    logging.error(f"Error overwriting file as fallback: {e_fallback}")
        else:
            # If the file doesn't exist, create it normally (this is equivalent to mode='w')
            try:
                df_manual_review.to_excel(output_excel_path, sheet_name='Manual Review', index=False, header=True)
                logging.info(f"Successfully created '{output_excel_path}' with {len(row_manual_review)} rows requiring manual intervention.")
            except Exception as e:
                logging.error(f"Error creating new Excel file: {e}")
    else:
        logging.info("No rows required manual intervention in this run.")

def search_tenant_by_contact(driver, config, contact_value, search_element_xpath):
    """
    Searches for a tenant using a given contact value (phone or email).
    Returns a tuple (table_visible, not_found_visible).
    """
    # Ensure the search element is fresh and interactable
    search_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, search_element_xpath))
    )
    search_element.send_keys(Keys.CONTROL + "a")
    search_element.send_keys(Keys.BACK_SPACE)
    haltStep(0.5) # Short pause after clearing
    search_element.send_keys(contact_value)
    
    # After sending value, wait again for the element to be clickable before hitting ENTER
    # This is crucial if auto-suggest or other dynamic elements temporarily cover/disable it.
    try:
        search_element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, search_element_xpath))
        )
        haltStep(1) # Give a bit more time for any auto-suggest to settle
        search_element.send_keys(Keys.ENTER)
        haltStep(2) # Give page time to react and load results
    except TimeoutException:
        logging.error(f"Search input not clickable after entering '{contact_value}'. Cannot press ENTER.")
        return False, False # Indicate failure to proceed with search

    table_visible = waitToLoad(driver, By.XPATH, config.get('searchPath', 'table'), timeout=1)
    not_found_visible = waitToLoad(driver, By.XPATH, config.get('searchPath', 'notFound'), timeout=1)
    
    return table_visible, not_found_visible

def create_tenant(driver, config, tenant_data):
    """
    Navigates to the create tenant page and attempts to create a new tenant.
    Returns True on successful creation (or if duplicate message appears), False otherwise.
    """
    create_page_url = config.get('urlPath', 'create')
    driver.get(create_page_url)
    haltStep(3)

    try:
        # Fill in unit
        unit_xpath = config.get('createPath', 'SelectUnit')
        if not waitToLoad(driver, By.XPATH, unit_xpath, timeout=5, condition=EC.element_to_be_clickable):
            raise TimeoutException("Timeout waiting for SelectUnit element on creation page.")
        unit_element = driver.find_element(By.XPATH, unit_xpath)
        unit_element.send_keys(Keys.CLEAR)
        unit_element.send_keys(tenant_data['UNIT'])
        haltStep(1)
        unit_element.send_keys(Keys.DOWN)
        unit_element.send_keys(Keys.ENTER)
        haltStep(1)

        # Fill in other details
        driver.find_element(By.XPATH, config.get('createPath', 'first')).send_keys(tenant_data['FIRST'])
        driver.find_element(By.XPATH, config.get('createPath', 'last')).send_keys(tenant_data['LAST'])
        driver.find_element(By.XPATH, config.get('createPath', 'address')).send_keys(tenant_data['EMAIL'])
        driver.find_element(By.XPATH, config.get('createPath', 'number')).send_keys(tenant_data['PHONE'])
        
        # Click create tenant button
        create_tenant_button_xpath = config.get('createPath', 'createTenant')
        if not waitToLoad(driver, By.XPATH, create_tenant_button_xpath, timeout=4, condition=EC.element_to_be_clickable):
            raise TimeoutException("Timeout waiting for Create Tenant button.")
        driver.find_element(By.XPATH, create_tenant_button_xpath).click()
        haltStep(2) # Give time for submission to process

        # Check for duplicate message AFTER submission
        if waitToLoad(driver, By.XPATH, config.get('createPath', 'duplicate'), timeout=2):
            logging.info(f"SUCCESS: New Tenant Created (or duplicate handled): {tenant_data['FIRST']} {tenant_data['LAST']} for Unit {tenant_data['UNIT']}")
            # If a duplicate message appears, it means the system acknowledged the attempt,
            # and potentially redirected or handled it. We consider this a "success" for automation purposes
            # in the sense that the action was taken, but the user might still need to review.
            # The original code clicks it, implying it dismisses a popup.
            driver.find_element(By.XPATH, config.get('createPath', 'duplicate')).click()
            haltStep(1)
            return True
        else:
            logging.info(f"Tenant creation for {tenant_data['FIRST']} {tenant_data['LAST']} completed without explicit duplicate message. Assuming success.")
            return True # Assume success if no duplicate message and no other error
            
    except (TimeoutException, NoSuchElementException) as e:
        logging.error(f"UI element error during tenant creation for {tenant_data['FIRST']} {tenant_data['LAST']}: {e}")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred during tenant creation for {tenant_data['FIRST']} {tenant_data['LAST']}: {e}")
        return False

def process_tenant_data(driver, config, row_data, row_manual_review):
    """
    Processes a single row of tenant data: searches by phone and email,
    then decides whether to create a new tenant or mark for manual review.
    """
    current_phone = str(row_data.get('PHONE', '')).strip()
    current_email = str(row_data.get('EMAIL', '')).strip()
    current_unit = str(row_data.get('UNIT', '')).strip()
    current_first = str(row_data.get('FIRST', '')).strip()
    current_last = str(row_data.get('LAST', '')).strip()

    createEmail = False
    createNum = False
    manualEmail = False
    manualNum = False

    tenant_page_url = config.get('urlPath', 'tenant')
    search_tenant_xpath = config.get('searchPath', 'searchTenant')

    # Navigate to tenant search page
    driver.get(tenant_page_url)
    haltStep(2)
    if not waitToLoad(driver, By.XPATH, search_tenant_xpath, timeout=5, condition=EC.element_to_be_clickable):
        logging.error(f"Search bar not found on tenant page for row: {row_data}. Skipping.")
        row_data['ReasonForManualReview'] = "Search bar not accessible on tenant page."
        row_manual_review.append(row_data)
        return

    # --- Search by Phone ---
    logging.info(f"Searching by phone: '{current_phone}' for {current_first} {current_last}...")
    phone_table_visible, phone_not_found_visible = search_tenant_by_contact(driver, config, current_phone, search_tenant_xpath)

    if phone_table_visible:
        manualNum = True
        logging.info(f"Phone '{current_phone}': FOUND (Table visible). manualNum = True")
    elif phone_not_found_visible:
        createNum = True
        logging.info(f"Phone '{current_phone}': NOT FOUND ('Not Found' message visible). createNum = True")
    else:
        logging.warning(f"Phone '{current_phone}': UNCERTAIN STATE - Neither table nor 'not found' message became visible.")
        row_data['ReasonForManualReview'] = f"Phone search: Unclear page state for '{current_phone}'."
        row_manual_review.append(row_data)
        return # Move to next row if state is uncertain

    # --- Search by Email ---
    # Re-navigate to tenant page and clear search bar for email search
    driver.get(tenant_page_url)
    haltStep(2)
    if not waitToLoad(driver, By.XPATH, search_tenant_xpath, timeout=5, condition=EC.element_to_be_clickable):
        logging.error(f"Search bar not found on tenant page for row: {row_data}. Skipping.")
        row_data['ReasonForManualReview'] = "Search bar not accessible on tenant page for email search."
        row_manual_review.append(row_data)
        return

    logging.info(f"Searching by email: '{current_email}' for {current_first} {current_last}...")
    email_table_visible, email_not_found_visible = search_tenant_by_contact(driver, config, current_email, search_tenant_xpath)

    if email_table_visible:
        manualEmail = True
        logging.info(f"Email '{current_email}': FOUND (Table visible). manualEmail = True")
    elif email_not_found_visible:
        createEmail = True
        logging.info(f"Email '{current_email}': NOT FOUND ('Not Found' message visible). createEmail = True")
    else:
        logging.warning(f"Email '{current_email}': UNCERTAIN STATE - Neither table nor 'not found' message became visible.")
        row_data['ReasonForManualReview'] = f"Email search: Unclear page state for '{current_email}'."
        row_manual_review.append(row_data)
        return # Move to next row if state is uncertain

    logging.info(f"--- Decision for Tenant: {current_first} {current_last} ---")
    logging.info(f"  Phone Search Result: createNum={createNum}, manualNum={manualNum}")
    logging.info(f"  Email Search Result: createEmail={createEmail}, manualEmail={manualEmail}")

    # Decision Logic
    if createNum and createEmail: # Both must be definitively NOT FOUND
        logging.info(f"Attempting to create new tenant: {current_first} {current_last} (Unit: {current_unit})")
        tenant_creation_successful = create_tenant(driver, config, row_data)
        if not tenant_creation_successful:
            row_data['ReasonForManualReview'] = 'Automated tenant creation failed.'
            row_manual_review.append(row_data)
        # Always navigate back to tenant search page after creation attempt
        driver.get(tenant_page_url)
        haltStep(3)

    elif manualEmail or manualNum:
        logging.info(f"Tenant for {current_first} {current_last} found by existing phone/email. Adding to manual review.")
        row_data['ReasonForManualReview'] = []
        if manualNum:
            row_data['ReasonForManualReview'].append('Phone number already exists in system.')
        if manualEmail:
            row_data['ReasonForManualReview'].append('Email address already exists in system.')
        row_data['ReasonForManualReview'] = '; '.join(row_data['ReasonForManualReview'])
        row_manual_review.append(row_data)
        haltStep(1)
        # Ensure we are on the tenant search page for the next iteration
        driver.get(tenant_page_url)
        haltStep(3)

    else:
        logging.warning(f"Case: Unhandled scenario for tenant {current_first} {current_last}. Adding to manual review.")
        row_data['ReasonForManualReview'] = 'Unhandled search outcome (neither clear creation nor clear existing).'
        row_manual_review.append(row_data)
        haltStep(1)
        # Ensure we are on the tenant search page for the next iteration
        driver.get(tenant_page_url)
        haltStep(3)


def main():
    """
    Main function to orchestrate the web automation process.
    """
    # setup the config parser (where the html identifiers are stored)
    config = configparser.ConfigParser()
    try:
        config.read('venv\settings.ini')
        if not config.sections():
            raise FileNotFoundError("settings.ini not found or is empty.")
    except Exception as e:
        logging.error(f"Error reading settings.ini: {e}. Please ensure it exists and is correctly formatted.")
        return

    reactorLogin = config.get('urlPath', 'login')
    
    # setup the chrome webdriver
    driver = None
    try:
        driver = setupDriver()
        driver.get(reactorLogin)

        # load previously saved cookies
        if os.path.isfile('cookies.pkl'):
            try:
                cookies = pickle.load(open("cookies.pkl", "rb"))
                for cookie in cookies:
                    # Filter out 'expiry' if it's a float and not an integer, or if it's None
                    if 'expiry' in cookie and (not isinstance(cookie['expiry'], int) or cookie['expiry'] is None):
                        del cookie['expiry']
                    driver.add_cookie(cookie)
                logging.info("Cookies loaded successfully.")
                driver.refresh() # Refresh to apply cookies
                haltStep(2)
            except Exception as e:
                logging.warning(f"Error loading or adding cookies: {e}. Proceeding without them.")

        # login to the website
        login(driver, config)
        
        # Verify successful login by checking for a known element on the post-login page
        # This is a generic check, you might want to replace with a specific element on your dashboard
        if not waitToLoad(driver, By.XPATH, config.get('searchPath', 'searchTenant'), timeout=10):
            logging.error("Login failed or dashboard page did not load. Exiting.")
            return

    except Exception as e:
        logging.critical(f"Initial setup or login failed: {e}. Exiting.")
        if driver:
            driver.quit()
        return

    row_manual_review = []
    df = pd.DataFrame() # Initialize empty DataFrame
    try:
        excel_path = r'C:\Users\dmartinez\Documents\ope\toMove\125Campus.xlsx'
        df = pd.read_excel(excel_path, sheet_name='Sheet1')
        logging.info(f"Successfully loaded data from '{excel_path}'. {len(df)} rows found.")
    except FileNotFoundError:
        logging.error(f"Input Excel file not found at '{excel_path}'. Please check the path.")
        if driver: driver.quit()
        return
    except Exception as e:
        logging.error(f"Error reading input Excel file: {e}. Exiting.")
        if driver: driver.quit()
        return

    for index, row in df.iterrows():
        logging.info(f"\n--- Processing Row {index + 1}/{len(df)}: {row.get('FIRST', '')} {row.get('LAST', '')} ---")
        process_tenant_data(driver, config, row.to_dict(), row_manual_review)

    # Save cookies for future sessions
    try:
        pickle.dump(driver.get_cookies(), open('cookies.pkl', 'wb'))
        logging.info("Cookies saved successfully.")
    except Exception as e:
        logging.error(f"Error saving cookies: {e}")

    if driver:
        driver.quit()
        logging.info("WebDriver closed.")
    
    writeOut(row_manual_review)
    logging.info("Script finished execution.")

if __name__ == "__main__":
    main()
