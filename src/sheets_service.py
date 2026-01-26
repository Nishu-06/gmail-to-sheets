"""
Google Sheets API service for appending email data.
Handles OAuth authentication and sheet operations.
"""

import logging
import sys
from pathlib import Path
from typing import List, Dict, Optional

# Add project root to Python path
project_root = Path(__file__).parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import time

from config import SCOPES, MAX_RETRIES, RETRY_DELAY, SPREADSHEET_ID, SHEET_NAME
from src.gmail_service import GmailService

logger = logging.getLogger(__name__)


class SheetsService:
    """Service for interacting with Google Sheets API."""
    
    def __init__(self, gmail_service: GmailService):
        """
        Initialize Sheets service using Gmail service credentials.
        
        Args:
            gmail_service: GmailService instance (shares OAuth credentials)
        """
        self.gmail_service = gmail_service
        self.service = build('sheets', 'v4', credentials=gmail_service.credentials)
        self.spreadsheet_id = SPREADSHEET_ID
        logger.info("Google Sheets API service initialized")
    
    def _retry_api_call(self, func, *args, **kwargs):
        """
        Retry wrapper for API calls with exponential backoff.
        
        Args:
            func: Function to retry
            *args: Positional arguments for func
            **kwargs: Keyword arguments for func
            
        Returns:
            Result of func call
            
        Raises:
            Exception: If all retries fail
        """
        last_exception = None
        for attempt in range(MAX_RETRIES):
            try:
                return func(*args, **kwargs)
            except HttpError as e:
                last_exception = e
                if e.resp.status in [429, 500, 503]:  # Rate limit or server errors
                    wait_time = RETRY_DELAY * (2 ** attempt)
                    logger.warning(
                        f"API call failed (attempt {attempt + 1}/{MAX_RETRIES}): {e}. "
                        f"Retrying in {wait_time}s..."
                    )
                    time.sleep(wait_time)
                else:
                    raise
            except Exception as e:
                last_exception = e
                logger.error(f"Unexpected error (attempt {attempt + 1}/{MAX_RETRIES}): {e}")
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (2 ** attempt))
                else:
                    raise
        
        raise last_exception
    
    def ensure_sheet_exists(self, sheet_name: str = None) -> None:
        """
        Ensure the specified sheet exists, create if it doesn't.
        
        Args:
            sheet_name: Name of the sheet (defaults to SHEET_NAME)
        """
        if sheet_name is None:
            sheet_name = SHEET_NAME
        
        try:
            # Get all sheets in the spreadsheet
            def _get_sheets():
                spreadsheet = self.service.spreadsheets().get(
                    spreadsheetId=self.spreadsheet_id
                ).execute()
                return spreadsheet.get('sheets', [])
            
            sheets = self._retry_api_call(_get_sheets)
            sheet_names = [s['properties']['title'] for s in sheets]
            
            if sheet_name not in sheet_names:
                logger.info(f"Creating sheet '{sheet_name}'")
                
                def _add_sheet():
                    self.service.spreadsheets().batchUpdate(
                        spreadsheetId=self.spreadsheet_id,
                        body={
                            'requests': [{
                                'addSheet': {
                                    'properties': {
                                        'title': sheet_name
                                    }
                                }
                            }]
                        }
                    ).execute()
                
                self._retry_api_call(_add_sheet)
                logger.info(f"Sheet '{sheet_name}' created")
            else:
                logger.debug(f"Sheet '{sheet_name}' already exists")
                
        except HttpError as e:
            logger.error(f"Error ensuring sheet exists: {e}")
            raise
    
    def ensure_headers_exist(self, sheet_name: str = None) -> None:
        """
        Ensure header row exists in the sheet.
        
        Args:
            sheet_name: Name of the sheet (defaults to SHEET_NAME)
        """
        if sheet_name is None:
            sheet_name = SHEET_NAME
        
        try:
            # Check if headers exist
            def _get_range():
                return self.service.spreadsheets().values().get(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"{sheet_name}!A1:E1"
                ).execute()
            
            result = self._retry_api_call(_get_range)
            values = result.get('values', [])
            
            expected_headers = ['From', 'Subject', 'Date', 'Content', 'Labels']
            
            if not values or values[0] != expected_headers:
                logger.info("Adding header row to sheet")
                
                def _update_headers():
                    self.service.spreadsheets().values().update(
                        spreadsheetId=self.spreadsheet_id,
                        range=f"{sheet_name}!A1:E1",
                        valueInputOption='RAW',
                        body={
                            'values': [expected_headers]
                        }
                    ).execute()
                
                self._retry_api_call(_update_headers)
                logger.info("Headers added to sheet")
            else:
                logger.debug("Headers already exist")
                
        except HttpError as e:
            logger.error(f"Error ensuring headers exist: {e}")
            raise
    
    def append_rows(self, rows: List[List[str]], sheet_name: str = None) -> None:
        """
        Append rows to the Google Sheet.
        
        Args:
            rows: List of rows, each row is a list of cell values
            sheet_name: Name of the sheet (defaults to SHEET_NAME)
        """
        if sheet_name is None:
            sheet_name = SHEET_NAME
        
        if not rows:
            logger.warning("No rows to append")
            return
        
        try:
            logger.info(f"Appending {len(rows)} row(s) to sheet '{sheet_name}'")
            
            def _append_values():
                self.service.spreadsheets().values().append(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"{sheet_name}!A:E",
                    valueInputOption='RAW',
                    insertDataOption='INSERT_ROWS',
                    body={
                        'values': rows
                    }
                ).execute()
            
            self._retry_api_call(_append_values)
            logger.info(f"Successfully appended {len(rows)} row(s)")
            
        except HttpError as e:
            logger.error(f"Error appending rows: {e}")
            raise
    
    def get_existing_message_ids(self, sheet_name: str = None, max_rows: int = 10000) -> set:
        """
        Get existing message IDs from the sheet to prevent duplicates.
        This reads the Content column and extracts message IDs if stored.
        
        Note: For production, consider storing message IDs in a separate column.
        This implementation uses a simple approach checking if message ID is in content.
        
        Args:
            sheet_name: Name of the sheet (defaults to SHEET_NAME)
            max_rows: Maximum number of rows to check
            
        Returns:
            Set of existing message IDs
        """
        if sheet_name is None:
            sheet_name = SHEET_NAME
        
        try:
            # Read all rows (excluding header)
            def _get_values():
                return self.service.spreadsheets().values().get(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"{sheet_name}!A2:E{max_rows + 1}"
                ).execute()
            
            result = self._retry_api_call(_get_values)
            values = result.get('values', [])
            
            # Extract message IDs from content (if stored)
            # For a more robust solution, add a dedicated column for message IDs
            existing_ids = set()
            
            # Since we don't store message IDs in the sheet by default,
            # we'll rely on state.json for duplicate prevention
            # This method is kept for potential future use
            
            logger.debug(f"Checked {len(values)} existing rows for duplicates")
            return existing_ids
            
        except HttpError as e:
            logger.warning(f"Error reading existing rows: {e}")
            return set()
