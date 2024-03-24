# Email-Validator-and-Domain-Analysis
Python application to validate and analyze email addresses from an Excel file, leveraging the Mailboxlayer API and Google Apps Script for advanced interaction with Google Sheets.


This Python application is designed to validate and analyze email addresses from an Excel file. It leverages the Mailboxlayer API for email validation and performs keyword matching to extract email domains.

# Features
Excel File Parsing: The script reads an Excel file containing email addresses using the Pandas library or similar tools for efficient data handling.

# 1.Email Validation: 
Utilizes the Mailboxlayer API to validate email addresses. Handles API responses and errors gracefully.

Domain Extraction: Extracts email domains from validated email addresses.

Keyword Matching: Performs keyword matching to categorize email addresses based on predefined keywords.

# Prerequisites
Before running the script, ensure you have the following:

Python 3.x installed on your system.
An API key from Mailboxlayer. Sign up here to obtain your API key..
Obtain an API key from Mailboxlayer and replace YOUR_API_KEY in the script with your actual API key.
# Usage:
Run the script Emails_Validation.ipynb in a Jupyter Notebook or any compatible environment.

Follow the prompts to input the path to your Excel file.

The script will validate email addresses, extract domains, and perform keyword matching. Results will be displayed or saved to a file as specified.

# Input
An Excel file 100 Emails.xlsx is provided in the repository for testing purposes.

# 2.Keyword Matching
This Python script extracts email domains from email addresses and performs keyword matching based on predefined keywords. It reads an Excel file containing email addresses and keywords, then processes the data to extract relevant information.

#Features
Domain Extraction: Extracts domain names from email addresses.

#Keyword Matching: Matches email addresses and domains against a list of predefined keywords.


Update the file_path variable in the script with the path to your Excel file.

Run the script keyword_matching.ipynb in a Jupyter Notebook or any compatible environment.

The script will extract domains from email addresses and perform keyword matching. Results will be displayed in a table format.

# 3. Google Apps Script Integration
This Google Apps Script provides advanced functionalities for parsing and updating Google Sheets based on column names. It includes custom menu options and a dialog window for selecting an Excel file path.

# Features
User Input Handling: Allows users to input the Excel file path either through a CLI prompt or a dialog window.

# File Path Validation: Validates the entered file path to ensure it meets specific criteria.

# Custom Menu: Creates a custom menu in Google Sheets with options to write or select the Excel file path.

# Prerequisites
Before using the Google Apps Script, ensure you have the following:

Access to Google Sheets and permission to run scripts.
# Usage
Open your Google Sheets document.

Click on "Extensions" > "Apps Script" to open the Apps Script editor.

Copy and paste the provided JavaScript Code0.gs and selectFilePage.html codes into the editor.

Save the script and close the editor.

Reload your Google Sheets document. You should now see a "Custom Menu" with options to write or select the Excel file path.

Use the menu options to input the file path or select it using the dialog window.

# Customization:
You can customize the validation logic in the validateFilePath function to match your specific requirements.
Modify the getUserInputCLI and selectFilePathDialog functions to change the input method or dialog window appearance.
You can customize the validation logic in the validateFilePath function to match your specific requirements.
Modify the getUserInputCLI and selectFilePathDialog functions to change the input method or dialog window appearance.
