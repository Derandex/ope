import configparser
import pandas as pd
import time
import os.path
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from fake_useragent import UserAgent
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
import keyring
import pickle
import re
import undetected_chromedriver as uc
import random

#setup the driver
def setupDriver():
    options = Options()
    options.add_argument(f'user-agent={UserAgent().random}')
    options.add_argument("--start-maximized")
    #fill in this line if you want to use your own user data
    #options.add_argument("user-data-dir=")
    driver = uc.Chrome(options=options, executable_path='chromedriver.exe')
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    driver.implicitly_wait(10)
    return driver

#logins the user to the website
def login(driver, config):
    driver.find_element(By.XPATH, config.get('loginPath', 'userName')).send_keys(keyring.get_credential(service_name='Onboard', username=None).username) 
    driver.find_element(By.XPATH, config.get('loginPath', 'password')).send_keys(keyring.get_credential(service_name='Onboard', username=None).password)
    driver.find_element(By.XPATH, config.get('loginPath', 'loginB')).click()
    haltStep(1)

#makes sure an element loads before clicking it
def waitToLoad(driver, byType, identifier, timeout):
    try:
        # Use visibility_of_element_located when elements are always present but hidden
        WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((byType, identifier)))
        return True # Element found and is visible
    except TimeoutException:
        # print(f"DEBUG: Element NOT visible within timeout: {identifier} (Type: {byType})") # Keep this for debugging if needed
        return False # Element not visible within timeout
    except Exception as e:
        print(f"DEBUG: An unexpected error occurred in waitToLoad for {identifier}: {e}")
        return False
    
#randomly stops the script between 1-3 seconds (to avoid detection)
def haltStep(x):
    #int = random.randint(1, 3)
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
    if row_manual_review:
        df_manual_review = pd.DataFrame(row_manual_review)
        output_excel_path = r'C:\Users\dmartinez\Documents\ope\TPM(3427).xlsx' # Define your output path

        if os.path.exists(output_excel_path):
            try:
                with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    # Check if the sheet 'Manual Review' exists in the writer's book
                    # If it exists, find the next available row. Otherwise, start from row 0.
                    start_row = 0
                    if 'Manual Review' in writer.book.sheetnames:
                        # Get the sheet object to find its max_row
                        sheet = writer.book['Manual Review']
                        start_row = sheet.max_row
                        df_manual_review.to_excel(writer, sheet_name='Manual Review', index=False, header=False, startrow=start_row)
                    else:
                        # If the sheet doesn't exist, create it with headers at the top
                        df_manual_review.to_excel(writer, sheet_name='Manual Review', index=False, header=True, startrow=0)

                print(f"\nSuccessfully appended {len(row_manual_review)} rows requiring manual intervention to '{output_excel_path}'")
            except Exception as e:
                print(f"Error appending data to Excel: {e}")
                # Fallback to overwrite if append fails, or handle as per your needs
                try:
                    df_manual_review.to_excel(output_excel_path, index=False)
                    print(f"Fallback: Overwrote existing file due to append error. Check '{output_excel_path}'")
                except Exception as e_fallback:
                    print(f"Error overwriting file as fallback: {e_fallback}")
        else:
            # If the file doesn't exist, create it normally (this is equivalent to mode='w')
            try:
                df_manual_review.to_excel(output_excel_path, sheet_name='Manual Review', index=False, header=True)
                print(f"\nSuccessfully created '{output_excel_path}' with {len(row_manual_review)} rows requiring manual intervention.")
            except Exception as e:
                print(f"Error creating new Excel file: {e}")
    else:
        print("\nNo rows required manual intervention in this run.")

def main():
    idRegex = r'id="(.*?)"'
    #setup the config parser (where the html identifiers are stored)
    config = configparser.ConfigParser()
    config.read('venv\settings.ini')

    reactorLogin = config.get('urlPath', 'login')
    tenantPage = config.get('urlPath', 'tenant')
    createPage = config.get('urlPath', '3427')

    #print(keyring.get_credential(service_name='Onboard', username=None).username)
    #setup the chrome webdriver
    driver = setupDriver()
    driver.get(reactorLogin)

    #load previously saved cookies
    if (os.path.isfile('cookies.pkl')):
        cookies = pickle.load(open("cookies.pkl", "rb"))
        for cookie in cookies:
            driver.add_cookie(cookie)

    #login to the website
    login(driver, config)
    row_manual_review = []
    try:
        df = pd.read_excel(r'C:\Users\dmartinez\Documents\ope\toMove\3427Polaris.xlsx', sheet_name='Sheet1')
    except FileNotFoundError:
        print('bruh find yo file')
    except Exception as e:
        print(f"Error reading file: {e}")
        exit

    for index, row in df.iterrows():

        current_phone = str(row.get('PHONE', ''))
        current_email = str(row.get('EMAIL', ''))
        current_unit = str(row.get('UNIT', ''))
        current_first = str(row.get('FIRST', ''))
        current_last = str(row.get('LAST', ''))


        createEmail = False
        createNum = False
        manualEmail = False
        manualNum = False


        #Navigate to tenant search page
        driver.get(tenantPage)
        haltStep(2)
        
        waitToLoad(driver, By.XPATH, config.get('searchPath', 'searchTenant'), timeout=3)
        haltStep(3)
        search_element = driver.find_element(By.XPATH, config.get('searchPath', 'searchTenant'))
        

        # --- Search by Phone ---
        print(f"Searching by phone: '{current_phone}'...")
        search_element.send_keys(Keys.CONTROL + "a")
        search_element.send_keys(Keys.BACK_SPACE)
        haltStep(2)
        search_element.send_keys(current_phone)
        waitToLoad(driver, By.XPATH, config.get('searchPath', 'searchTenant'), timeout=3)
        haltStep(4)
        search_element.send_keys(Keys.ENTER)
        haltStep(2) # Give page time to react

        table_visible = waitToLoad(driver, By.XPATH, config.get('searchPath', 'table'), timeout=1)
        not_found_visible = waitToLoad(driver, By.XPATH, config.get('searchPath', 'notFound'), timeout=1)

        if table_visible:
            manualNum = True
            print(f"Phone '{current_phone}': FOUND (Table visible). manualNum = True")
        elif not_found_visible:
            createNum = True
            print(f"Phone '{current_phone}': NOT FOUND ('Not Found' message visible). createNum = True")
        else:
            print(f"Phone '{current_phone}': UNCERTAIN STATE - Neither table nor 'not found' message became visible.")
            row_data = row.to_dict()
            row_data['ReasonForManualReview'] = f"Phone search: Unclear page state for '{current_phone}'."
            row_manual_review.append(row_data)
            continue

        # Clear search bar before next search
        search_element.send_keys(Keys.CONTROL + "a")
        search_element.send_keys(Keys.BACK_SPACE)

        search_element = driver.find_element(By.XPATH, config.get('searchPath', 'searchTenant'))
        haltStep(2)

        print(f"Searching by email: '{current_email}'...")
        search_element.send_keys(current_email)
        haltStep(2)
        search_element.send_keys(Keys.ENTER)
        haltStep(2)

        table_visible = waitToLoad(driver, By.XPATH, config.get('searchPath', 'table'), timeout=1)
        not_found_visible = waitToLoad(driver, By.XPATH, config.get('searchPath', 'notFound'), timeout=1)

        if table_visible:
            manualEmail = True
            print(f"Email '{current_email}': FOUND (Table visible). manualEmail = True")
        elif not_found_visible:
            createEmail = True
            print(f"Email '{current_email}': NOT FOUND ('Not Found' message visible). createEmail = True")
        else:
            print(f"Email '{current_email}': UNCERTAIN STATE - Neither table nor 'not found' message became visible.")
            row_data = row.to_dict()
            row_data['ReasonForManualReview'] = f"Email search: Unclear page state for '{current_email}'."
            row_manual_review.append(row_data)
            continue

        # Clear search bar before decision/next iteration
        search_element.send_keys(Keys.CONTROL + "a")
        search_element.send_keys(Keys.BACK_SPACE)
        haltStep(1)

        print(f"--- Decision for Row {index + 1} ---")
        print(f"   Phone Search Result: createNum={createNum}, manualNum={manualNum}")
        print(f"   Email Search Result: createEmail={createEmail}, manualEmail={manualEmail}")

        # Decision Logic
        if createNum and createEmail: # Both must be definitively NOT FOUND
            print(f"Attempting to create new tenant: {current_first} {current_last} (Unit: {current_unit})")
            driver.get(createPage)
            haltStep(3)

            try:
                #if not waitToLoad(driver, By.XPATH, config.get('createPath', 'SelectUnit'), timeout=2):
                #    raise TimeoutException("Timeout waiting for SelectUnit element on creation page.")
                unit_element = driver.find_element(By.XPATH, config.get('createPath', 'SelectUnit'))
                unit_element.send_keys(Keys.CLEAR)
                unit_element.send_keys(current_unit)
                haltStep(1)
                unit_element.send_keys(Keys.DOWN)
                unit_element.send_keys(Keys.ENTER)

                driver.find_element(By.XPATH, config.get('createPath', 'first')).send_keys(current_first)
                driver.find_element(By.XPATH, config.get('createPath', 'last')).send_keys(current_last)
                driver.find_element(By.XPATH, config.get('createPath', 'address')).send_keys(current_email)
                driver.find_element(By.XPATH, config.get('createPath', 'number')).send_keys(current_phone)
                haltStep(1)
                haltStep(3)
                haltStep(2)
                
                driver.find_element(By.XPATH, config.get('createPath', 'createTenant')).click()
                

                # Check for duplicate message AFTER submission
                if waitToLoad(driver, By.XPATH, config.get('createPath', 'duplicate'), timeout=2):
                    driver.find_element(By.XPATH, config.get('createPath', 'duplicate')).click()
                    haltStep(2)
                    print(f"SUCCESS: New Tenant Created: {current_first} {current_last} for Unit {current_unit}")
                else:
                    continue
            except Exception as e:
                print(f"ERROR during tenant creation for row {index + 1}: {e}")
                row_data = row.to_dict()
                row_data['ReasonForManualReview'] = f'Error during automated creation attempt: {e}'
                row_manual_review.append(row_data)

            driver.get(tenantPage)
            haltStep(3)

        elif manualEmail or manualNum:
            print(f"Tenant for row {index + 1} found by existing phone/email. Adding to manual review.")
            row_data = row.to_dict()
            row_data['ReasonForManualReview'] = []
            if manualNum:
                row_data['ReasonForManualReview'].append('Phone number already exists in system.')
            if manualEmail:
                row_data['ReasonForManualReview'].append('Email address already exists in system.')
            row_data['ReasonForManualReview'] = '; '.join(row_data['ReasonForManualReview'])
            row_manual_review.append(row_data)
            haltStep(1)
            driver.get(tenantPage)
            haltStep(3)

        else:
            print(f"Case: Unhandled scenario for row {index + 1}. Adding to manual review.")
            row_data = row.to_dict()
            row_data['ReasonForManualReview'] = 'Unhandled search outcome (neither clear creation nor clear existing).'
            row_manual_review.append(row_data)
            haltStep(1)
            driver.get(tenantPage)
            haltStep(3)

    pickle.dump(driver.get_cookies(), open('cookies.pkl', 'wb'))
    driver.quit()
    writeOut(row_manual_review)

if __name__ == "__main__":
    main()
