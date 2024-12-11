import os
from datetime import datetime

import pandas as pd
import PyPDF2
import re

# EXPIRATION_DATE = datetime.datetime(2024, 11, 4)  # Replace with your desired expiration date
#
# # Check if the program has expired
# if datetime.datetime.now() > EXPIRATION_DATE:
#     #print("The program has expired and is no longer available for use.")
#     exit()

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

        return text


def clean_extracted_data(data):
    if data['Principal Office Street Address']:
        principal_end = re.search(r'UNITED STATES|UNITED ST ATES', data['Principal Office Street Address'],
                                  re.IGNORECASE)
        if principal_end:

            data['Principal Office Street Address'] = data['Principal Office Street Address'][
                                                      :principal_end.end()].strip()

    if data['Principal Office Mailing Address'] and "Filed" in data['Principal Office Mailing Address']:

        data['Principal Office Mailing Address'] = "NULL"


    if data['Registered Agent Mailing Address']:

        agent_end = re.search(r'UNITED STATES|UNITED ST ATES', data['Registered Agent Mailing Address'], re.IGNORECASE)
        if agent_end:

            data['Registered Agent Mailing Address'] = data['Registered Agent Mailing Address'][
                                                       :agent_end.end()].strip()
        else:

            data['Registered Agent Mailing Address'] = "NULL"

    if data['Registered Agent Street Address']:

        agent_end = re.search(r'UNITED STATES|UNITED ST ATES', data['Registered Agent Street Address'], re.IGNORECASE)
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

def extract_data_from_text(text,excel_file_path=None, row_index=None):

    if "annualreport" in text.lower().replace(" ", ""):
           return None
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
           # Check if 'Business Status' is not NULL and has the value 'ACTIVE'
    if data['Business Status'] != 'NULL' and data['Business Status'].upper() == 'ACTIVE':
                if excel_file_path is not None and row_index is not None:
                    remove_file_path_from_excel(excel_file_path, row_index)
                    return None
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

        if "NameStreet" in line:
            agent_name = ""
            street_address_start = ""
            mailing_start_index = None  # Initialize mailing_start_index to None

            # Step 1: Extract Registered Agent name until a number appears
            for j in range(i + 1, i + 3):  # Look at the next 1-2 lines for the name
                if j < len(lines):
                    next_line = lines[j].strip()

                    # Find the first digit in the line to separate name from address
                    num_index = next((idx for idx, char in enumerate(next_line) if char.isdigit()), None)

                    if num_index is not None:
                        agent_name += next_line[:num_index].strip()
                        street_address_start = next_line[num_index:].strip()  # Start of street address
                        break
                    else:
                        agent_name += " " + next_line

            data['Registered Agent'] = agent_name.strip()

            # Step 2: Extract Registered Agent Street Address until "UNITED STATES" or "USA"
            street_address = street_address_start
            for k in range(j + 1, len(lines)):  # Start extracting from the next line after the name
                next_line = lines[k].strip()

                # Stop at "UNITED STATES" or "USA" if found
                if re.search(r"\b(UNITED STATES|UNITED ST ATES|USA)\b", next_line, re.IGNORECASE):
                    # Extract up to "UNITED STATES" or "USA" and handle cases where the next part is concatenated
                    match = re.search(r"(UNITED STATES|UNITED ST ATES|USA)(\s*\d+)?", next_line, re.IGNORECASE)
                    if match:
                        street_address += " " + match.group(1).strip()
                        if match.group(2):  # If there's a concatenated number, it’s the start of mailing address
                            mailing_start_index = k
                            next_line = match.group(2).strip()  # Treat concatenated number as first mailing line
                        else:
                            mailing_start_index = k + 1  # Continue with next line for mailing address
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
            pdf_data = extract_data_from_text(pdf_text, path_excel_file, row_index)

            # Only append to the output if pdf_data is not None (i.e., Business Status wasn't ACTIVE)
            if pdf_data:
                pdf_data = clean_extracted_data(pdf_data)
                append_to_excel(output_excel_path, pdf_data)
                print(f"Data from {pdf_path} processed and added to {output_excel_path}")




paths_excel_file=input("Enter the full path to your file containing pdfs path: ").strip()
output_excel_path=input("Enter the name of the output file:").strip()
# paths_excel_file = 'initial_paths.xlsx'  # The Excel file with paths generated in the previous code
# output_excel_path = 'initial_output.xlsx'  # The Excel file where extracted data will be saved

# Process all PDFs listed in the paths Excel file
process_pdfs_from_excel(paths_excel_file, output_excel_path)
