import os
from datetime import datetime

import pandas as pd
import PyPDF2
import re


# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        print(f"Error: PDF file does not exist at {pdf_path}")
        return None

    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()

        print("Extracted Text from PDF:\n", text)  # Debugging: print the text extracted from the PDF
        return text


def clean_extracted_data(data):
    if data['Attention:']:
        if data['Attention:'].startswith("This document is a public record"):
            data['Attention:'] = 'NULL'

    if data['Registered Agent Street Address']:
        agent_end = re.search(r'UNITED STATES|UNITED ST ATES|USA', data['Registered Agent Street Address'], re.IGNORECASE)
        if agent_end:
            data['Registered Agent Street Address'] = data['Registered Agent Street Address'][
                                                       :agent_end.end()].strip()

    return data

def remove_file_path_from_excel(excel_file_path, row_index):
    try:
        df = pd.read_excel(excel_file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    if row_index < len(df):
        df = df.drop(index=row_index).reset_index(drop=True)
        try:
            df.to_excel(excel_file_path, index=False)

        except Exception as e:
            print(f"Error saving Excel file: {e}")
    else:
        #print(f"Row index {row_index} is out of bounds for the Excel file.")
        pass

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
            print("Initial report found. Skipping extraction.")
            return None  # Exit function and return None if "Initial report" is found

    # for i, line in enumerate(lines):
    #     if 'NameStreet' in line:
    #         # Check if the next line exists
    #         if i + 1 < len(lines):
    #             name_line = lines[i + 1].strip()
    #
    #             # Use regular expression to extract only the alphabetic part of the name
    #             # This will match until the first numeric character
    #             match = re.match(r'^[A-Za-z\s]+', name_line)
    #             if match:
    #                 # Extracted name from the matched portion
    #                 name = match.group(0).strip()
    #                 data['Registered Agent'] = name
    #             else:
    #                 data['Registered Agent'] = 'NULL'
    #         else:
    #             data['Registered Agent'] = 'NULL'
    #
    for i, line in enumerate(lines):
        line = line.strip()

        # Check if this line hints at the start of "Registered Agent" information
        if "NameStreet" in line:
            # Step 1: Extract Registered Agent Name until a number appears
            agent_name = ""
            street_address_start = ""
            mailing_start_index = None  # Initialize mailing_start_index to None

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
                    mailing_start_index = k + 1  # Set index to start mailing address from the line after this one
                    break
                else:
                    street_address += " " + next_line

            data['Registered Agent Street Address'] = street_address.strip()

    for i, line in enumerate(lines):
        line = line.strip()

        # Check for the start of "NameStreet"
        if "NameStreet" in line:
            mailing_address_start = None
            mailing_address = ""

            # Search within the next 3-4 lines for the initial "UNITED STATES" or similar indicator
            for j in range(i + 1, min(i + 5, len(lines))):  # Limit search to next 3-4 lines
                next_line = lines[j].strip()

                # Look for the first occurrence of "UNITED STATES", "UNITED ST ATES", or "USA"
                match = re.search(r"(UNITED STATES|UNITED ST ATES|USA)(\s*\d+)?", next_line, re.IGNORECASE)
                if match:
                    # Identify the start of mailing address after "UNITED STATES" (with or without appended number)
                    if match.group(2):  # If a number is appended right after
                        mailing_address_start = match.start(2)  # Start from this appended number
                        mailing_address = next_line[mailing_address_start:].strip()
                    else:
                        # Find the next number after "UNITED STATES" if not appended
                        num_index = next((idx for idx, char in enumerate(next_line, match.end(1)) if char.isdigit()),
                                         None)
                        if num_index is not None:
                            mailing_address_start = num_index
                            mailing_address = next_line[mailing_address_start:].strip()

                    # Break out of this loop once the starting point for the mailing address is found
                    break

            # Start collecting the mailing address until encountering "UNITED STATES" or similar indicator again
            if mailing_address_start is not None:
                for k in range(j + 1, len(lines)):
                    current_line = lines[k].strip()

                    # Stop once the next "UNITED STATES", "UNITED ST ATES", or "USA" is encountered
                    if re.search(r"(UNITED STATES|UNITED ST ATES|USA)", current_line, re.IGNORECASE):
                        mailing_address += " " + current_line.split("UNITED")[0].strip() + " UNITED STATES"
                        break
                    else:
                        mailing_address += " " + current_line

            # Assign the extracted mailing address or set as NULL if not found
            data['Registered Agent Mailing Address'] = mailing_address.strip() if mailing_address else 'NULL'

    amount_received_processed = False

    # Check if 'Principal Phone' is not NULL
    if data['Principal Phone'] != 'NULL':
        # Loop to find "Amount Received:" and extract the amount and email if found on the same line
        for i, line in enumerate(lines):
            line = line.strip()
            print(line)
            # Check if "Amount Received:" is in line
            if "Amount Received:" in line:
                match = re.search(r'\$(\d{1,5}\.\d{2})([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})', line)

                if match:
                    # Extract amount and email
                    data['Amount Received'] = match.group(1)
                    data['Principal Email'] = match.group(2)

                    # Debug: Print extracted values
                    print(f"Extracted Amount: {data['Amount Received']}")
                    print(f"Extracted Principal Email: {data['Principal Email']}")
                # Continue to check further lines for "Email:" in case email is not found after "Amount Received"
                continue

            # If "Email:" is found and the next line doesn’t contain "This document is a public record"
            if "Email:" in line:
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if ".com" in next_line and "This document is a public record" not in next_line:
                        data['Principal Email'] = next_line  # Assign the next line as the email if valid
                        print(f"Extracted Principal Email from 'Email:' line: {next_line}")
                    else:
                        data['Principal Email'] = 'NULL'
                # Break the loop once "Email:" is processed
                break

    else:
        # If 'Principal Phone' is NULL, check for "Email:" and extract email from the next line if valid
        for i, line in enumerate(lines):
            line = line.strip()
            if "Email:" in line:
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if ".COM" in next_line and "This document is a public record" not in next_line:
                        data['Principal Email'] = next_line  # Assign the next line as the email if valid
                        print(f"Extracted Principal Email from 'Email:' line: {next_line}")
                    else:
                        data['Principal Email'] = 'NULL'
                # Break the loop once "Email:" is processed
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
            # data['Registered Agent Street Address'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'

            # Also, extract Principal Email from the line before the street address
            email_line = lines[i - 1].strip()
            match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email_line)
            if match:
                email = match.group(0).strip()
                #data['Principal Email'] = email[5:] if len(email) > 5 else 'NULL'
            i += 1

        # Check for Mailing Address
        elif 'Mailing Address:' in line:
            # data['Registered Agent Mailing Address'] = lines[i + 1].strip() if i + 1 < len(lines) else 'NULL'
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


# Function to append data to the Excel file
def append_to_excel(file_path, data):
    if data is None:
       # print("Skipping row insertion as the data extraction was skipped.")
        return

    df = pd.DataFrame([data])

    if not os.path.exists(file_path):
        df.to_excel(file_path, index=False)
    else:
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            startrow = writer.sheets['Sheet1'].max_row  # Write after the last row
            df.to_excel(writer, index=False, header=False, startrow=startrow)


# Main function to process each PDF in the paths file
def process_pdfs_from_excel(path_excel_file, output_excel_path):
    # Load the Excel file that contains the paths
    paths_df = pd.read_excel(path_excel_file)

    # Ensure the required columns exist
    if 'File Path' not in paths_df.columns:
        print("The Excel file must contain a 'File Path' column with paths to PDF files.")
        return

    # Process each PDF file listed in the Excel
    for row_index, row in paths_df.iterrows():
        pdf_path = row['File Path']

        # Extract text from the PDF file
        pdf_text = extract_text_from_pdf(pdf_path)
        if pdf_text:
            # Extract structured data from the text, passing the excel file path and row index
            pdf_data = extract_data_from_text(pdf_text)

            # Only append to the output if pdf_data is not None (i.e., Business Status wasn't ACTIVE)
            if pdf_data:
                pdf_data = clean_extracted_data(pdf_data)
                append_to_excel(output_excel_path, pdf_data)
                print(f"Data from {pdf_path} processed and added to {output_excel_path}")




paths_excel_file=input("Enter the full path to your file containing pdfs path: ").strip()
output_excel_path=input("Enter the name of the output file:").strip()
# paths_excel_file = 'annual_paths.xlsx'
# output_excel_path = 'annual_output.xlsx'

# Process all PDFs listed in the paths Excel file
process_pdfs_from_excel(paths_excel_file, output_excel_path)
