import time
import os
import shutil
from datetime import datetime

import PyPDF2
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, ElementNotInteractableException
from selenium.common.exceptions import NoSuchElementException

business_name = ""


# Path to the Excel file
excel_file_path = input("Enter the path to the Excel file with UBI numbers: ").strip()
default_download_dir = input("Enter the path to your download directory: ").strip()
project_folder = input("Enter the full path to your project folder: ").strip()
excel_output_path = input("Enter the path and name for the Excel file for output (e.g., C:/path/to/output.xlsx): ").strip()

# Chrome options setup
options = Options()
options.add_experimental_option("detach", True)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("window_size=1280,800")
options.add_argument("--disable-popup-blocking")
options.add_argument("--disable-save-password-bubble")

# Initialize WebDriver
driver = webdriver.Chrome(options=options)
pdf_paths=[]
# Function to wait for a new file to appear in the download directory
def wait_for_new_file(directory, old_files):
    new_file = None
    while not new_file:
        files = os.listdir(directory)
        new_files = [f for f in files if f not in old_files and not f.endswith('.crdownload')]
        if new_files:
            new_file = new_files[0]
        else:
            time.sleep(1)
    return new_file

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        print(f"Error: PDF file does not exist at {pdf_path}")
        return None
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
    return text


# Function to extract data from PDF text
def extract_data_from_text(text):
    # Check if "Annual Report" is present in the entire text
    if "annualreport" in text.lower().replace(" ", ""):
        print("Annual report found. Skipping extraction.")  # Debugging statement
        return None  # Exit function if "Annual Report" is found

    # Initialize data dictionary with default values as 'NULL'
    data = {
        'Business Status': 'NULL',
        'Principal Office Street Address': 'NULL',
        'Principal Office Mailing Address': 'NULL',
        'Principal Phone': 'NULL',
        'Principal Email': 'NULL',
        'Registered Agent': 'NULL',
        'Registered Agent Street Address': 'NULL',
        'Registered Agent Mailing Address': 'NULL',
        'Attention:': 'NULL',
        'Return Address Email': 'NULL',
        'Return Address Address': 'NULL'
    }

    # Split the text into lines
    lines = text.splitlines()

    # First loop: Extract Business Status, Principal Office Street Address, Mailing Address, and Phone
    for i, line in enumerate(lines):
        line = line.strip()

        # Check for Business Status
        if 'Business Status:' in line:
            data['Business Status'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'

        # Check for Principal Office Street Address
        elif 'Principal Office Street Address' in line or 'Street Address' in line:
            # Get the line following the "Street Address" indicator
            address_line = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'

            # Remove any leading alphabetic characters to start from the first number
            address_start = re.search(r'\d', address_line)  # Find the first number
            if address_start:
                address_line = address_line[address_start.start():].strip()

            # Continue adding lines to the address until "UNITED STATES" or "UNITED ST ATES" is encountered
            j = i + 2
            while j < len(lines):
                next_line = lines[j].strip()
                address_line += " " + next_line
                # Stop when encountering "UNITED STATES" or "UNITED ST ATES"
                if re.search(r'UNITED\s*STATES', next_line, re.IGNORECASE):
                    break
                j += 1

            # Remove any extra text after "UNITED STATES" or "UNITED ST ATES"
            match = re.search(r'(.*?UNITED\s*STATES)', address_line, re.IGNORECASE)
            if match:
                address_line = match.group(1).strip()

            data['Principal Office Street Address'] = address_line.strip()


        # Check for Principal Office Mailing Address
        elif 'Principal Office Mailing Address' in line or 'Mailing Address' in line:
            data['Principal Office Mailing Address'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'

        # Check for Phone
        elif 'Phone:' in line:
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'
            # If the next line starts with 'Email:', set phone to 'NULL'
            if next_line.startswith('Email:'):
                data['Principal Phone'] = 'NULL'
            else:
                data['Principal Phone'] = next_line if next_line else 'NULL'


    for i, line in enumerate(lines):
        line = line.strip()
        # Check if the line contains an email with '@'
        if 'Email:' in line:
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'
            # If the next line contains an email (characterized by '@')
            if '@' in next_line:
                email = re.sub(r'\s+', '', next_line)  # Remove spaces in the email
                data['Principal Email'] = email

        # Handle broken emails in lines, check for '@'
        elif '@' in line:
            # Try to reconstruct broken emails like OPERA TIONS@T AXMAKER.COM
            email_parts = re.findall(r'\S+', line)
            email = ''.join(email_parts)  # Join broken parts together
            data['Principal Email'] = email

        # Iterate through lines to extract "Registered Agent" related data

    for i, line in enumerate(lines):
        line = line.strip()

        # Check if this line hints at the start of "Registered Agent" information
        if "NameStreet" in line:
            # Step 1: Extract Registered Agent Name until a number appears
            agent_name = ""
            street_address_start = ""

            for j in range(i + 1, i + 3):  # Look at the next 1-2 lines for the name
                if j < len(lines):
                    next_line = lines[j].strip()

                    # Find the first digit in the line to separate name from address
                    num_index = next((idx for idx, char in enumerate(next_line) if char.isdigit()), None)

                    if num_index is not None:
                        agent_name += next_line[:num_index].strip()
                        street_address_start = next_line[
                                               num_index:].strip()  # Remaining part is the start of the street address
                        break
                    else:
                        agent_name += " " + next_line

            data['Registered Agent'] = agent_name.strip()

            # Step 2: Extract Registered Agent Street Address until "UNITED STATES" or "UNITED ST ATES" appears
            street_address = street_address_start
            for k in range(j + 1, len(lines)):  # Start extracting from the next line after the name
                next_line = lines[k].strip()

                if "UNITED STATES" in next_line or "UNITED ST ATES" in next_line:
                    street_address += " " + next_line.split("UNITED")[0].strip() + " UNITED STATES"
                    mailing_start_index = k  # Set index to start mailing address from this line
                    break
                else:
                    street_address += " " + next_line

            data['Registered Agent Street Address'] = street_address.strip()

            # Step 3: Extract Registered Agent Mailing Address
            mailing_address = ""
            for l in range(mailing_start_index, len(lines)):
                next_line = lines[l].strip()

                # Check if the line has numbers and contains "UNITED STATES" or "UNITED ST ATES"
                if l == mailing_start_index:  # Handle numbers concatenated with "UNITED STATES"
                    num_index = next((idx for idx, char in enumerate(next_line) if char.isdigit()), None)
                    if num_index is not None and num_index < len(next_line):
                        mailing_address = next_line[num_index:].strip()
                else:
                    # Continue until the end of mailing address
                    mailing_address += " " + next_line
                    if "UNITED STATES" in next_line or "UNITED ST ATES" in next_line:
                        break

            data['Registered Agent Mailing Address'] = mailing_address.strip() if mailing_address else 'NULL'



    # Fourth loop: Extract Attention, Email, and Address
    for i, line in enumerate(lines):
        line = line.strip()
        # Find the "RETURN ADDRESS FOR THIS FILING"
        if "RETURN ADDRESS FOR THIS FILING" in line:
            # Extract Attention:
            attention_line = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'
            if 'Attention:' in attention_line:
                next_attention_value = lines[i + 2].strip() if i + 2 < len(lines) else 'NULL'
                # If the next line contains Email:, set Attention to 'NULL'
                if 'Email:' in next_attention_value:
                    data['Attention:'] = 'NULL'
                else:
                    data['Attention:'] = next_attention_value

            # Extract Email:
            email_line = lines[i + 3].strip() if i + 3 < len(lines) else 'NULL'
            if 'Email:' in email_line:
                next_email_value = lines[i + 4].strip() if i + 4 < len(lines) else 'NULL'
                # If the next line contains Address:, set Email to 'NULL'
                if 'Address:' in next_email_value:
                    data['Return Address Email'] = 'NULL'
                else:
                    data['Return Address Email'] = next_email_value

            # Extract Address:
            address_line = lines[i + 5].strip() if i + 5 < len(lines) else 'NULL'
            if 'Address:' in address_line:
                next_address_value = lines[i + 6].strip() if i + 6 < len(lines) else 'NULL'
                # If the next line contains UPLOAD ADDITIONAL DOCUMENTS, set Address to 'NULL'
                if 'UPLOAD ADDITIONAL DOCUMENTS' in next_address_value:
                    data['Return Address Address'] = 'NULL'
                else:
                    data['Return Address Address'] = next_address_value

    return data


def clean_extracted_data(data):
    if data['Principal Office Street Address']:
        # Search for "UNITED STATES" or "UNITED ST ATES" in the `Principal Office Street Address`
        principal_end = re.search(r'UNITED STATES|UNITED ST ATES', data['Principal Office Street Address'],
                                  re.IGNORECASE)
        if principal_end:
            # Keep only the address up to "UNITED STATES"
            data['Principal Office Street Address'] = data['Principal Office Street Address'][
                                                      :principal_end.end()].strip()

    return data


# Function to append data to an Excel file
def append_to_excel(file_path, data):
    df = pd.DataFrame([data])
    if not os.path.exists(file_path):
        df.to_excel(file_path, index=False)
    else:
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            startrow = writer.sheets['Sheet1'].max_row  # Write after the last row
            df.to_excel(writer, index=False, header=False, startrow=startrow)


def find_element_with_retry(xpath, timeout=40, retry_on_failure=True, retries=2):
    """Function to handle element search with retry and refresh mechanism"""
    for attempt in range(retries + 1):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            return element  # Element found, return it
        except TimeoutException:
            if attempt < retries and retry_on_failure:
                driver.refresh()  # Refresh page and retry
                wait_for_loader_to_disappear()  # Wait for loader if it appears
                close_unexpected_dialog()
            else:
                return None  # Element not found after retries


def process_initial_report_fulfilled():
    # Click the final icon/button to download
    max_attempts = 10  # Maximum number of tbody elements to check
    for attempt in range(1, max_attempts + 1):
        try:
            close_unexpected_dialog()
            report_xpath = f'/html/body/div[1]/ng-include/div/section/div/div/div[1]/div[5]/div/div/div/div[2]/div/table/tbody[{attempt}]/tr/td[1]/span'
            report_element = find_element_with_retry(report_xpath, timeout=5, retries=1)

            if not report_element:
                print(f"No report element found for row {attempt}.")
                return False  # Exit early if no element is found

            report_text = report_element.text.strip()

            if report_text == 'INITIAL REPORT - FULFILLED':
                download_button_xpath = f'/html/body/div[1]/ng-include/div/section/div/div/div[1]/div[5]/div/div/div/div[2]/div/table/tbody[{attempt}]/tr/td[3]/i'
                download_button = find_element_with_retry(download_button_xpath, timeout=10, retries=1)

                if download_button:
                    download_button.click()
                    print("Download button clicked.")
                    return True  # Return True if the download button was successfully clicked
            else:
                print(f"Row {attempt} contains: {report_text}, moving to the next row.")
        except TimeoutException:
            print(f"No more rows found, or row {attempt} does not exist.")
            break

    # If "INITIAL REPORT - FULFILLED" is not found in any row
    print(f"'INITIAL REPORT - FULFILLED' not found. Skipping to the next.")
    return False  # Return False if the report is not found after all attempts


# Function to wait for the page to fully load
def wait_for_page_load(driver, timeout=40):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )

def wait_for_loader_to_disappear(timeout=15):
    try:
        WebDriverWait(driver, timeout).until_not(
            EC.presence_of_element_located((By.XPATH, '//*[@id="loaderDiv"]'))
        )
        #print("Loader has disappeared.")
        return True
    except TimeoutException:
        print("Loader did not disappear in time.")
        return False


def close_unexpected_dialog():
    try:
        # Wait for the dialog message to appear
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ngdialog1-aria-describedby"]'))
        )

        # Extract the message
        message = driver.find_element(By.XPATH, '//*[@id="ngdialog1-aria-describedby"]').text

        # Print the dialog message for debugging
        print(f"Unexpected dialog detected with message: {message}")

        # Locate the OK button using the button class 'ngdialog-button btn-success'
        close_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'ngdialog-button btn-success')]"))
        )

        # Use JavaScript to click the OK button
        driver.execute_script("arguments[0].click();", close_button)

        time.sleep(0.5)  # Small delay to ensure the dialog is closed

        # If the message is 'null', refresh the page
        if message.lower() == 'null':
            print("Refreshing page due to 'null' message.")
            driver.refresh()
            time.sleep(2)

    except TimeoutException:
        #print("No unexpected dialog appeared.")
        pass
    except NoSuchElementException:
        # If no dialog is found, continue the script
        print("Dialog elements not found. Proceeding with the script.")


# Read UBI numbers from Excel file
try:
    df = pd.read_excel(excel_file_path)
except FileNotFoundError:
    print("Error: The specified Excel file was not found.")
    exit()

# Locate the UBI number column
ubi_column = None
for col in df.columns:
    if col.lower() == 'ubi number':
        ubi_column = col
        break

if ubi_column is None:
    print("Error: 'UBI Number' column not found in the Excel file.")
    exit()

# Loop through the UBI numbers and interact with the website
for ubi_number in df[ubi_column].dropna().astype(str).str.strip():
    driver.get("https://ccfs.sos.wa.gov/#/Home")
    driver.refresh()
    wait_for_page_load(driver)

    # Process each UBI number
    print(f"Processing UBI Number: {ubi_number}")
    ubi_number = ubi_number.replace(" ", "")
    time.sleep(1)

    try:
        # Enter UBI number
        ubi_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="UBINumber"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", ubi_element)
        ubi_element.clear()
        ubi_element.send_keys(ubi_number)
        ubi_element.send_keys(Keys.RETURN)
    except TimeoutException:
        print("UBI Number input field was not found.")
        continue
    time.sleep(1.5)

    # Main interaction logic - Filing link
    try:
        driver.refresh()
        filing_link_xpath = '/html/body/div/ng-include/div/section/div[2]/div[1]/div/div[2]/div/div[1]/table/tbody[1]/tr/td[1]/a'

        wait_for_loader_to_disappear(10)
        close_unexpected_dialog()

        element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, filing_link_xpath))
        )
        business_name = element.text
        # Ensure the element is in view and try to click it
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        driver.execute_script("arguments[0].click();", element)


    except Exception as e:
        print(f"Business search page was not found or element was not clickable. ")
        continue
    time.sleep(1.5)

    # Filing History Button Logic
    try:
        driver.refresh()
        filing_history_button_xpath = '/html/body/div/ng-include/div/section/div/div[2]/div[2]/input[1]'

        wait_for_loader_to_disappear(10)

        # Call function to close any unexpected dialog before proceeding
        close_unexpected_dialog()

        # Ensure the button is clickable
        element_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, filing_history_button_xpath))
        )

        time.sleep(1)

        retries = 3
        while retries > 0:
            try:
                # Ensure the dialog is closed each time we retry
                close_unexpected_dialog()

                # Use JavaScript to click the button to ensure the event is triggered
                driver.execute_script("arguments[0].click();", element_btn)
                #print("Filing history button clicked!")

                # Wait for the URL to change to the expected business filings page
                WebDriverWait(driver, 10).until(
                    EC.url_contains("/BusinessSearch/BusinessFilings")
                )
                #print("Navigated to the Business Filings page.")
                break

            except Exception as e:
                retries -= 1
                print(f"Retrying... {retries} attempts left. Error: {e}")
                time.sleep(1)
    except Exception as e:
        print(f"Filing history page was not found or element was not clickable. ")
        continue

    time.sleep(1)
    report_found = False
    try:
        row_num = 1
        # Loop through the rows and check for 'INITIAL REPORT'
        while True:
            try:

                short_wait = WebDriverWait(driver, 5)

                # Construct XPaths dynamically based on the row number
                report_xpath = f'/html/body/div/ng-include/div/section/div/div/div[1]/div[4]/div[2]/div/div/table/tbody[1]/tr[{row_num}]/td[4]'
                view_documents_xpath = f'/html/body/div/ng-include/div/section/div/div/div[1]/div[4]/div[2]/div/div/table/tbody[1]/tr[{row_num}]/td[5]/a'

                short_wait.until(
                    EC.url_contains("/BusinessSearch/BusinessFilings")
                )

                # Get the text in the 4th column of the current row
                report_element = short_wait.until(
                    EC.presence_of_element_located((By.XPATH, report_xpath))
                )
                report_text = report_element.text.strip()

                # If 'INITIAL REPORT' is found, click the 'View Documents' button
                if report_text == 'INITIAL REPORT':
                    print(f"'INITIAL REPORT' found in row {row_num}. Clicking 'View Documents'.")

                    # Click the 'View Documents' button
                    view_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, view_documents_xpath))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", view_button)
                    time.sleep(1)
                    view_button.click()

                    # Wait for the loader to disappear
                    wait_for_loader_to_disappear(10)
                    report_found = process_initial_report_fulfilled()

                    if report_found:
                        print("Report processed successfully.")
                        row_num = 1
                        break  # Exit loop if report is processed
                else:
                    pass

                row_num += 1  # Increment row number after each check
            except TimeoutException:
                break
            except NoSuchElementException:
                # If the element is not found, exit the loop
                print(f"No 'INITIAL REPORT' found in row {row_num}. Skipping to the next UBI.")
                break

        # If no report was found, continue to the next UBI
        if not report_found:
            print(f"No 'INITIAL REPORT' found for UBI {ubi_number}. Continuing to next UBI.")
            continue  # Move to the next UBI

    except Exception as e:
        print(f"No INITIAL REPORT FOUND! Error: {str(e)}")

    if report_found:
        # Process the downloaded file if "ANNUAL REPORT - FULFILLED" was found
        old_files = os.listdir(default_download_dir)
        downloaded_file = wait_for_new_file(default_download_dir, old_files)

        if downloaded_file:
            latest_file_path = os.path.join(default_download_dir, downloaded_file)

            # Move the downloaded file to the project folder
            new_pdf_path = os.path.join(project_folder, downloaded_file)
            shutil.move(latest_file_path, new_pdf_path)

            # Append the file path to our list
            pdf_paths.append({
                'UBI Number': ubi_number,
                'Business Name':business_name,
                'File Name': downloaded_file,
                'File Path': new_pdf_path
            })

    # Save the list of paths to an Excel file
    if pdf_paths:
        df = pd.DataFrame(pdf_paths)
        df.to_excel(excel_output_path, index=False)
        print(f"Excel file with PDF paths created at: {excel_output_path}")

print("Done!")
driver.quit()
