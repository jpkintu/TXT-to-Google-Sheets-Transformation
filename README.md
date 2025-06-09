# TXT-to-Google-Sheets-Transformation

This Google Apps Script automates importing data from TXT files stored in a specific Google Drive folder into a Google Sheet. It organizes valid data into a main sheet and flags malformed rows in a separate sheet for review.

## Features

- Imports data from all TXT files in a target Drive folder
- Populates the main sheet with parsed data
- Flags rows with unexpected column counts in a "Malformed Data" sheet
- Easy-to-use custom menu in Google Sheets to refresh data

## How to Use

1. Set your folder and spreadsheet IDs in the script configuration.
2. Open your Google Sheet and use the Data Import menu to refresh data.
3. Review imported data in the main sheet and check the "Malformed Data" sheet for any issues.
