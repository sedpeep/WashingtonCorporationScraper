# UBI Number Processing and PDF Report Automation

## Overview
This script automates the process of searching for UBI numbers on a website, downloading relevant PDF reports, and extracting key data. It processes both **initial** and **final reports**, parses the downloaded PDFs, and saves extracted data into an Excel file.

---

## Features
1. **Automated UBI Search**:
   - Reads UBI numbers from an Excel file and searches for them on a designated website.
   - Handles dynamic elements like loaders and unexpected dialogs.

2. **PDF Download and Parsing**:
   - Downloads PDF reports for each UBI and moves them to a project folder.
   - Extracts data such as business status, addresses, contact details, and more.

3. **PDF Report Types**:
   - Identifies and processes **Annual Reports**.
   - Skips irrelevant reports such as "Initial Reports."

4. **Excel Integration**:
   - Reads UBI numbers from an input Excel file.
   - Saves processed file paths and extracted data into a new Excel file.

5. **Error Handling**:
   - Skips missing or invalid entries gracefully.
   - Handles exceptions for unavailable elements or unresponsive pages.

---

## Prerequisites
1. **Python Libraries**:
   - `selenium`: For browser automation.
   - `PyPDF2`: For PDF parsing.
   - `pandas`: For Excel file manipulation.

   Install the required libraries using:
   ```bash
   pip install selenium PyPDF2 pandas openpyxl
   ```

2. **Browser Driver**:
   - Google Chrome with the corresponding ChromeDriver installed.
   - Ensure ChromeDriver is added to your system's PATH.

3. **Input Files**:
   - An Excel file containing UBI numbers with a column labeled "UBI Number."

4. **Folder Setup**:
   - Specify the default download directory and project folder for organizing downloaded files.

---

## Setup and Usage
1. **Prepare Input File**:
   - Ensure the Excel file contains a column labeled "UBI Number" with valid UBI numbers.

2. **Run the Script**:
   - Execute the script in your terminal or IDE:
     ```bash
     python ubi_automation.py
     ```
   - Provide the required paths when prompted:
     - Path to the input Excel file.
     - Download directory.
     - Project folder.
     - Output Excel file path.

3. **Output**:
   - Downloaded PDFs are moved to the specified project folder.
   - Extracted data is saved to an Excel file, including:
     - UBI number, business name, file path, and extracted fields.

---

## Key Functionalities
1. **Dynamic Web Interaction**:
   - Handles interactive elements like dialogs, loaders, and buttons.
   - Implements retries for element retrieval and page reloads.

2. **PDF Parsing**:
   - Extracts structured data from PDFs using text recognition.
   - Skips irrelevant reports based on content (e.g., "Initial Reports").

3. **Excel Integration**:
   - Reads UBI numbers from input files and writes extracted data to output files.
   - Supports appending new data to existing Excel files.

---

## Customization
- **Fields to Extract**:
   - Update the `extract_data_from_text` function to extract additional data from PDFs.
- **File Organization**:
   - Modify folder paths to suit your directory structure.
- **Error Logging**:
   - Enhance error handling for detailed logs.

---

## Notes
- Ensure the website structure matches the XPath selectors used in the script. If the website layout changes, the script may need updates.
- Maintain stable internet connectivity during execution.

---
