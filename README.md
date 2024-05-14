# Email Extractor for Google Sheets
This Google Sheets `Email Extractor script` allows you to extract email addresses from URLs or cell text within your Google Sheets. The script provides an easy-to-use menu for running the extraction process and displays the results directly in your spreadsheet.
## Intro 
This script adds a custom menu to your Google Sheets called `Email Extractor` with options to:
1. Extract Emails from URLs: Extracts email addresses from the web pages of URLs listed in Column A and outputs the results in Column B.
2. Extract Emails from Cell Text: Extracts email addresses from the text of a specified cell and displays them in an alert dialog.
3. The script is capable of identifying both standard email formats and obfuscated formats (e.g., example [at] domain [dot] com).
# How to Use

## Installation
1. Open your Google Sheets document.
2. Go to `Extensions` > `Apps Script`.
3. Delete any code in the script editor and replace it with the provided script.
4. Save the script with a name (e.g., `Email Extractor`).
5. Refresh your Google Sheets document.

## Usage
After refreshing, you will see a new menu item called `Email Extractor` in your Google Sheets toolbar.
1. Click on Email Extractor to access the following options:
3. Extract Emails from URLs: This option extracts emails from URLs listed in Column A of your active sheet and displays the results in Column B.
2. Extract Emails from Cell Text: This option extracts emails from the text of a specified cell and displays the results in an alert dialog.

## Extract Emails from URLs
1. Enter the URLs in Column A of your Google Sheets document.
2. Click on `Email Extractor` > `Extract Emails from URLs`.
3. The script will process each URL and output the found email addresses in Column B.

## Extract Emails from Cell Text
1. Select `Email Extractor` > `Extract Emails from Cell Text`.
2. Enter the cell reference (e.g., `A1`) when prompted.
3. The script will extract emails from the text in the specified cell and display them in an alert dialog.

# Why to Use
This script is useful for:

1. Gathering contact information from multiple web pages efficiently.
2. Extracting emails from any text content within your spreadsheet without manually searching through each cell.
3. Automating the process of email extraction, saves you time and effort. 

# Requirements
To use this script, you need:

1. A Google account.
2. Access to Google Sheets.
3. Basic knowledge of using Google Sheets and navigating the Apps Script editor.

## Support
For any issues or questions, please file a GitHub issue on this repository.
