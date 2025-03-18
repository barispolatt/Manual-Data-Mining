# Web Scraping Contact Information with Selenium and Undetected Chromedriver

## Overview
This Python script automates the process of extracting contact information (emails and phone numbers) from company websites. It performs Google searches for company names provided in an Excel file, navigates to the most relevant company website, and scrapes contact details. The extracted data is then saved in an output Excel file.

## Features
- Uses **undetected-chromedriver** to bypass bot detection.
- Employs **randomized user-agents** to minimize detection.
- Handles **CAPTCHA detection manually** when necessary.
- Extracts **emails and phone numbers** using regular expressions.
- Avoids social media sites (LinkedIn, Facebook, Instagram, etc.).
- Supports multiple languages for "Contact" page detection.
- Saves results in an Excel file.

## Requirements
- Python 3.7+
- Google Chrome
- Chromedriver (compatible with your Chrome version)
- Required Python libraries:
  ```bash
  pip install selenium pandas undetected-chromedriver openpyxl
  ```

## Installation
1. Clone or download this repository.
2. Install the required dependencies using the command above.
3. Ensure Google Chrome and the matching Chromedriver are installed.

## Usage
1. Prepare an Excel file (`621.xlsx`) with a list of company names in the first column.
2. Run the script:
   ```bash
   python basiccontactminer.py
   ```
3. The script will:
   - Search each company on Google.
   - Visit the most relevant website (excluding social media links).
   - Attempt to find a contact page.
   - Extract email addresses and phone numbers.
   - Save the results in `output_contacts.xlsx`.

## Output
The script generates an Excel file (`output_contacts.xlsx`) with the following columns:
- **Company Name**: The name of the company from the input file.
- **Company Website**: The detected website URL.
- **Company E-mail**: Extracted email addresses (if found).
- **Company Telephone**: Extracted phone numbers (if found).

## Notes
- Some websites may require manual CAPTCHA solving.
- If no relevant website is found, the output will mark it as "Not Found."
- Random delays are added to mimic human behavior and reduce detection risk.

## Disclaimer
This script is intended for **educational and ethical** web scraping purposes only. Ensure compliance with a website’s terms of service before scraping.

## Author
Barış Polat

