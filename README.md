
# Google Sheets Grading System

This repository contains a Python script for updating Google Sheets based on grading criteria. The script retrieves data from a specified Google Sheets spreadsheet, applies grading criteria to the data, and updates the spreadsheet accordingly.

## Prerequisites

Before running the script, ensure you have the following:

- Python 3 installed on your machine.
- Google API credentials (`credentials.json`) obtained from the [Google Developers Console](https://console.developers.google.com/).
- The required Python packages installed (`google-auth`, `google-auth-oauthlib`, `google-auth-httplib2`, `google-api-python-client`).
```bash
	pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
   ```

## Usage

1. **Clone this repository to your local machine:**

   ```bash
   git clone https://github.com/hetmanek/ScriptGoogleSpreadsheets
   ```

2. **Navigate to the project directory:**


3. **Place your `credentials.json` file in the project directory.**

4. **Run the script:**

   ```bash
   python main.py
   ```
