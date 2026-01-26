"""
Configuration file for Gmail to Sheets automation.
Contains settings for Gmail API, Sheets API, and application behavior.
"""

import os
from pathlib import Path

# Project root directory
PROJECT_ROOT = Path(__file__).parent

# Credentials and tokens paths
CREDENTIALS_DIR = PROJECT_ROOT / "credentials"
CREDENTIALS_FILE = CREDENTIALS_DIR / "credentials.json"
TOKEN_FILE = CREDENTIALS_DIR / "token.json"

# State persistence file (tracks processed email IDs)
STATE_FILE = PROJECT_ROOT / "state.json"

# Google API scopes
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/spreadsheets'
]

# Gmail API settings
GMAIL_QUERY = 'is:unread in:inbox'  # Unread emails in inbox
GMAIL_BATCH_SIZE = 50  # Number of emails to fetch per batch

# Google Sheets settings
# These will be set via environment variables or user input
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID', '1-1yT6B0ZpZb5glxK_iJYIb2z9RdwbEwHM_Jj5XGnnjg')
SHEET_NAME = os.getenv('SHEET_NAME', 'Emails')  # Default sheet name

# Email filtering settings
SUBJECT_FILTER = os.getenv('SUBJECT_FILTER', '')  # Optional: filter by subject keyword
EXCLUDE_NO_REPLY = os.getenv('EXCLUDE_NO_REPLY', 'false').lower() == 'true'
LAST_24_HOURS_ONLY = os.getenv('LAST_24_HOURS_ONLY', 'false').lower() == 'true'  # Process only emails from last 24 hours

# Retry settings
MAX_RETRIES = 3
RETRY_DELAY = 2  # seconds

# Logging settings
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
LOG_FILE = PROJECT_ROOT / "gmail_to_sheets.log"
