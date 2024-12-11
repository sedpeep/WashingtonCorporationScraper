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


# Path to the Excel file
excel_file_path = input("Enter the path to the Excel file with UBI numbers: ").strip()
default_download_dir = input("Enter the path to your download directory: ").strip()
project_folder = input("Enter the full path to your project folder: ").strip()
excel_output_path = input("Enter the path and name for the Excel file to store downloaded pdfs paths (e.g., C:/path/to/output.xlsx): ").strip()

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
pdf_paths = []
business_name = ""

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

    # Check if "Initial report" is present in any line

    for line in lines:
        if "INITIAL REPORT" in line.upper() or "INITIAL REPOR T" in line.upper():  # Corrected condition
            #print("Initial report found. Skipping extraction.")
            return None  # Exit function and return None if "Initial report" is found

    for i, line in enumerate(lines):
        if 'NameStreet' in line:
            # Check if the next line exists
            if i + 1 < len(lines):
                name_line = lines[i + 1].strip()

                # Use regular expression to extract only the alphabetic part of the name
                # This will match until the first numeric character
                match = re.match(r'^[A-Za-z\s]+', name_line)
                if match:
                    # Extracted name from the matched portion
                    name = match.group(0).strip()
                    data['Registered Agent'] = name
                else:
                    data['Registered Agent'] = 'NULL'
            else:
                data['Registered Agent'] = 'NULL'

    for i, line in enumerate(lines):
        # Check for the first occurrence of "Amount Received:"
        if "Amount Received:" in line:
            # Extract the amount and email from the line
            match = re.search(r'\$(\d{1,5}\.\d{2})([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})', line)

            if match:
                # Amount received after the dollar sign
                amount_received = match.group(1)
                # Email immediately after the dollar amount
                principal_email = match.group(2)

                # Store values in the data dictionary
                data['Amount Received'] = amount_received
                data['Principal Email'] = principal_email

                # Debug: Print extracted values
                print(f"Extracted Amount: {amount_received}")
                print(f"Extracted Principal Email: {principal_email}")

            # Break the loop after the first occurrence is found and processed
            break

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Check for Business Status
        if 'Business Status:' in line:
            data['Business Status'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'
            i += 1

        # Check for Principal Office Street Address
        elif 'Principal Office Street Address' in line or 'Principal Of fice Street Address' in line:
            data['Principal Office Street Address'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'
            i += 1

        # Check for Principal Office Mailing Address
        elif 'Principal Office Mailing Address' in line or 'Principal Of fice Mailing Address' in line:
            # Check if the next line exists and does not contain "EXPIRATION DATE"
            if i + 1 < len(lines) and 'EXPIRATION DATE' not in lines[i + 1].upper():
                data['Principal Office Mailing Address'] = lines[i + 1].strip()
            else:
                # If the next line is "EXPIRATION DATE" or doesn't exist, set it to 'NULL'
                data['Principal Office Mailing Address'] = 'NULL'
            i += 1

        # Check for Street Address
        elif 'Street Address:' in line:
            data['Registered Agent Street Address'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'

            # Also, extract Principal Email from the line before the street address
            email_line = lines[i - 1].strip()
            match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email_line)
            if match:
                email = match.group(0).strip()
                #data['Principal Email'] = email[5:] if len(email) > 5 else 'NULL'
            i += 1

        # Check for Mailing Address
        elif 'Mailing Address:' in line:
            data['Registered Agent Mailing Address'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'
            i += 1

        # Check for Phone
        elif 'Phone:' in line:
            # Check if the next line exists
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()

                # Check if the next line is an email field instead of a phone number
                if "Email:" in next_line or "EMAIL:" in next_line:
                    data['Principal Phone'] = 'NULL'
                else:
                    data['Principal Phone'] = next_line
            else:
                data['Principal Phone'] = 'NULL'
            i += 1

        # Check for Registered Agent Name (reuse previous logic that worked)
        elif 'Register ed Agent' in line:
            # The next line contains the full name for Registered Agent
            registered_agent_line = lines[i + 2].strip() if i + 2 < len(lines) else 'NULL'
            agent_split = re.split(r'(\d+)', registered_agent_line, maxsplit=1)
            #data['Registered Agent'] = agent_split[0].strip()
            i += 4


        elif 'Attention:' in line:

            if i + 1 < len(lines) and 'Email:' not in lines[i + 1]:
                data['Attention:'] = lines[i + 1].strip()
                i += 1
            else:
                data['Attention:'] = 'NULL'



        elif 'Email:' in line:
            if i + 1 < len(lines) and 'Address:' in lines[i + 1]:
                data['Return Address Email'] = 'NULL'
            else:
                data['Return Address Email'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'

            i += 1

        # Check for Return Address Address:
        elif 'Address:' in line:
            # Check if the next line contains the value for Address or 'UPLOAD ADDITIONAL DOCUMENTS'
            if i + 1 < len(lines) and 'UPLOAD ADDITIONAL DOCUMENTS' not in lines[i + 1]:
                data['Return Address Address'] = lines[i + 1].strip()
            else:
                data['Return Address Address'] = 'NULL'
            i += 1

        i += 1
        # Check if certain critical fields are missing
    if data['Business Status'] == 'NULL' and data['Principal Office Street Address'] == 'NULL':
        print(
            "Skipping row due to missing critical data: Business Status and Principal Office Street Address are NULL.")
        return None  # Skip row by returning None

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

            if report_text == 'ANNUAL REPORT - FULFILLED':
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
                if report_text == 'ANNUAL REPORT':
                    print(f"'ANNUAL REPORT' found in row {row_num}. Clicking 'View Documents'.")

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
                print(f"No 'ANNUAL REPORT' found in row {row_num}. Skipping to the next UBI.")
                break

        # If no report was found, continue to the next UBI
        if not report_found:
            print(f"No 'ANNUAL REPORT' found for UBI {ubi_number}. Continuing to next UBI.")
            continue  # Move to the next UBI

    except Exception as e:
        print(f"No ANNUAL REPORT FOUND! ")

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
                'UBI_Number': ubi_number,
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