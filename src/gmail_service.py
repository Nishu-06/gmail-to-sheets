"""
Gmail API service for fetching and managing emails.
Handles OAuth 2.0 authentication and email retrieval.
"""

import os
import pickle
import logging
import sys
from pathlib import Path
from typing import List, Dict, Optional

# Add project root to Python path
project_root = Path(__file__).parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import time

from config import SCOPES, CREDENTIALS_FILE, TOKEN_FILE, GMAIL_QUERY, MAX_RETRIES, RETRY_DELAY, LAST_24_HOURS_ONLY

logger = logging.getLogger(__name__)


class GmailService:
    """Service for interacting with Gmail API."""
    
    def __init__(self):
        """Initialize Gmail service with OAuth authentication."""
        self.service = None
        self.credentials = None
        self._authenticate()
    
    def _authenticate(self) -> None:
        """
        Authenticate with Gmail API using OAuth 2.0.
        Handles token refresh and storage.
        """
        creds = None
        
        # Load existing token if available
        if TOKEN_FILE.exists():
            try:
                creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
                logger.info("Loaded existing OAuth token")
            except Exception as e:
                logger.warning(f"Failed to load token: {e}")
        
        # If no valid credentials, request authorization
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    logger.info("Refreshed expired OAuth token")
                except Exception as e:
                    logger.error(f"Failed to refresh token: {e}")
                    creds = None
            
            if not creds:
                if not CREDENTIALS_FILE.exists():
                    raise FileNotFoundError(
                        f"Credentials file not found at {CREDENTIALS_FILE}. "
                        "Please download credentials.json from Google Cloud Console."
                    )
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(CREDENTIALS_FILE), SCOPES
                )
                creds = flow.run_local_server(port=0)
                logger.info("Completed OAuth flow")
            
            # Save credentials for next run
            TOKEN_FILE.parent.mkdir(exist_ok=True)
            with open(TOKEN_FILE, 'w') as token:
                token.write(creds.to_json())
            logger.info(f"Saved OAuth token to {TOKEN_FILE}")
        
        self.credentials = creds
        self.service = build('gmail', 'v1', credentials=creds)
        logger.info("Gmail API service initialized")
    
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
    
    def get_unread_emails(self, query: Optional[str] = None) -> List[Dict]:
        """
        Fetch unread emails from inbox.
        
        Args:
            query: Optional Gmail search query (defaults to GMAIL_QUERY)
            
        Returns:
            List of email message dictionaries with 'id' and 'threadId'
        """
        if query is None:
            query = GMAIL_QUERY
        
        # Add last 24 hours filter if enabled
        if LAST_24_HOURS_ONLY:
            from datetime import datetime, timedelta
            cutoff_date = (datetime.now() - timedelta(hours=24)).strftime('%Y/%m/%d')
            query = f"{query} after:{cutoff_date}"
            logger.info(f"Filtering emails from last 24 hours (after {cutoff_date})")
        
        try:
            logger.info(f"Fetching unread emails with query: {query}")
            
            def _fetch_messages():
                result = self.service.users().messages().list(
                    userId='me',
                    q=query,
                    maxResults=500
                ).execute()
                return result.get('messages', [])
            
            messages = self._retry_api_call(_fetch_messages)
            logger.info(f"Found {len(messages)} unread emails")
            return messages
            
        except HttpError as e:
            logger.error(f"Error fetching emails: {e}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error fetching emails: {e}")
            raise
    
    def get_email_details(self, message_id: str) -> Dict:
        """
        Get full details of a specific email message.
        
        Args:
            message_id: Gmail message ID
            
        Returns:
            Full message object from Gmail API
        """
        try:
            def _get_message():
                return self.service.users().messages().get(
                    userId='me',
                    id=message_id,
                    format='full'
                ).execute()
            
            message = self._retry_api_call(_get_message)
            return message
            
        except HttpError as e:
            logger.error(f"Error fetching email {message_id}: {e}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error fetching email {message_id}: {e}")
            raise
    
    def mark_as_read(self, message_id: str) -> None:
        """
        Mark an email as read by removing the UNREAD label.
        
        Args:
            message_id: Gmail message ID to mark as read
        """
        try:
            def _modify_message():
                self.service.users().messages().modify(
                    userId='me',
                    id=message_id,
                    body={'removeLabelIds': ['UNREAD']}
                ).execute()
            
            self._retry_api_call(_modify_message)
            logger.debug(f"Marked email {message_id} as read")
            
        except HttpError as e:
            logger.error(f"Error marking email {message_id} as read: {e}")
            raise
        except Exception as e:
            logger.error(f"Unexpected error marking email {message_id} as read: {e}")
            raise
    
    def mark_multiple_as_read(self, message_ids: List[str]) -> None:
        """
        Mark multiple emails as read.
        
        Args:
            message_ids: List of Gmail message IDs
        """
        for msg_id in message_ids:
            try:
                self.mark_as_read(msg_id)
            except Exception as e:
                logger.warning(f"Failed to mark {msg_id} as read: {e}")
                # Continue with other emails even if one fails
