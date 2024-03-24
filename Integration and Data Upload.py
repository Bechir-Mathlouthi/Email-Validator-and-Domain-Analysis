import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# Authorize Google Sheets API
scope = ['httpsspreadsheets.google.comfeeds', 'httpswww.googleapis.comauthdrive']
creds = ServiceAccountCredentials.from_json_keyfile_name('your_credentials.json', scope)
client = gspread.authorize(creds)

# Open the Google Sheets document
sheet = client.open('your_spreadsheet_name')  # Replace with your actual document name

# Access the first sheet (ALL EMAILS)
worksheet_all_emails = sheet.get_worksheet(0)

# Access the second sheet (Keywords for matching)
worksheet_keywords = sheet.get_worksheet(1)

# Read data from the first sheet into a DataFrame
data_all_emails = worksheet_all_emails.get_all_values()
df_all_emails = pd.DataFrame(data_all_emails[1], columns=data_all_emails[0])

# Read data from the second sheet into a list
data_keywords = worksheet_keywords.col_values(1)[1]

# Process the data (you can perform any processing here)
# For example, print the data
print(ALL EMAILS)
print(df_all_emails)
print(nKeywords for matching)
print(data_keywords)

# Update the data in the Google Sheets document (if needed)

worksheet.update('A1', 'New Data')

#append new data to the worksheet
worksheet.append_row(['New Data'])

