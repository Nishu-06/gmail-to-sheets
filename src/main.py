"""
Main entry point for Gmail to Sheets automation.
Orchestrates email fetching, parsing, and sheet updates with state persistence.
"""

import json
import logging
import sys
from pathlib import Path
from typing import Set, List, Dict
from datetime import datetime

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config import (
    STATE_FILE, SPREADSHEET_ID, SHEET_NAME, LOG_LEVEL, LOG_FILE
)
from src.gmail_service import GmailService
from src.email_parser import parse_email_message
from src.sheets_service import SheetsService

# Configure logging with UTF-8 encoding support
class SafeStreamHandler(logging.StreamHandler):
    """Stream handler that safely handles Unicode characters for Windows console."""
    def emit(self, record):
        try:
            msg = self.format(record)
            # Replace problematic Unicode characters for console output
            try:
                self.stream.write(msg)
                self.stream.write(self.terminator)
            except UnicodeEncodeError:
                # Fallback: encode with error handling
                safe_msg = msg.encode('ascii', errors='replace').decode('ascii')
                self.stream.write(safe_msg)
                self.stream.write(self.terminator)
            self.flush()
        except Exception:
            self.handleError(record)

# Configure logging
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL.upper()),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        SafeStreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)


class StateManager:
    """Manages state persistence for processed email IDs."""
    
    def __init__(self, state_file: Path):
        """
        Initialize state manager.
        
        Args:
            state_file: Path to state JSON file
        """
        self.state_file = state_file
        self.processed_ids: Set[str] = set()
        self._load_state()
    
    def _load_state(self) -> None:
        """Load processed email IDs from state file."""
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r') as f:
                    data = json.load(f)
                    self.processed_ids = set(data.get('processed_message_ids', []))
                logger.info(f"Loaded {len(self.processed_ids)} processed email IDs from state")
            except Exception as e:
                logger.warning(f"Failed to load state: {e}. Starting fresh.")
                self.processed_ids = set()
        else:
            logger.info("No existing state file found. Starting fresh.")
    
    def _save_state(self) -> None:
        """Save processed email IDs to state file."""
        try:
            self.state_file.parent.mkdir(exist_ok=True)
            data = {
                'processed_message_ids': list(self.processed_ids),
                'last_updated': datetime.now().isoformat()
            }
            with open(self.state_file, 'w') as f:
                json.dump(data, f, indent=2)
            logger.debug(f"Saved state with {len(self.processed_ids)} processed IDs")
        except Exception as e:
            logger.error(f"Failed to save state: {e}")
    
    def is_processed(self, message_id: str) -> bool:
        """
        Check if an email has been processed.
        
        Args:
            message_id: Gmail message ID
            
        Returns:
            True if already processed, False otherwise
        """
        return message_id in self.processed_ids
    
    def mark_processed(self, message_id: str) -> None:
        """
        Mark an email as processed and save state.
        
        Args:
            message_id: Gmail message ID
        """
        self.processed_ids.add(message_id)
        self._save_state()
    
    def mark_multiple_processed(self, message_ids: List[str]) -> None:
        """
        Mark multiple emails as processed.
        
        Args:
            message_ids: List of Gmail message IDs
        """
        for msg_id in message_ids:
            self.processed_ids.add(msg_id)
        self._save_state()


def main():
    """Main execution function."""
    logger.info("=" * 60)
    logger.info("Gmail to Sheets Automation - Starting")
    logger.info("=" * 60)
    
    # Validate configuration
    if not SPREADSHEET_ID:
        logger.error("SPREADSHEET_ID not set. Please set it in config.py or as environment variable.")
        sys.exit(1)
    
    try:
        # Initialize services
        logger.info("Initializing Gmail service...")
        gmail_service = GmailService()
        
        logger.info("Initializing Sheets service...")
        sheets_service = SheetsService(gmail_service)
        
        # Initialize state manager
        logger.info("Loading state...")
        state_manager = StateManager(STATE_FILE)
        
        # Ensure sheet exists with headers
        logger.info(f"Ensuring sheet '{SHEET_NAME}' exists...")
        sheets_service.ensure_sheet_exists(SHEET_NAME)
        sheets_service.ensure_headers_exist(SHEET_NAME)
        
        # Fetch unread emails
        logger.info("Fetching unread emails from Gmail...")
        email_messages = gmail_service.get_unread_emails()
        
        if not email_messages:
            logger.info("No unread emails found. Exiting.")
            return
        
        logger.info(f"Found {len(email_messages)} unread email(s)")
        
        # Process emails
        new_emails: List[Dict] = []
        processed_ids: List[str] = []
        
        for msg_ref in email_messages:
            message_id = msg_ref['id']
            
            # Skip if already processed
            if state_manager.is_processed(message_id):
                logger.debug(f"Skipping already processed email: {message_id}")
                continue
            
            try:
                # Get full email details
                message = gmail_service.get_email_details(message_id)
                
                # Parse email
                parsed = parse_email_message(message)
                
                if parsed is None:
                    logger.debug(f"Email {message_id} filtered out or failed to parse")
                    # Still mark as processed to avoid reprocessing
                    state_manager.mark_processed(message_id)
                    continue
                
                new_emails.append(parsed)
                processed_ids.append(message_id)
                # Safely log subject (handle Unicode characters)
                subject_preview = parsed['subject'][:50].encode('ascii', errors='replace').decode('ascii')
                logger.info(f"Parsed email: {subject_preview}...")
                
            except Exception as e:
                logger.error(f"Error processing email {message_id}: {e}")
                # Don't mark as processed if there was an error
                continue
        
        if not new_emails:
            logger.info("No new emails to add to sheet")
            return
        
        # Prepare rows for sheet
        # Google Sheets has a 50,000 character limit per cell
        # Email body is already limited to 10,000 chars in parser
        # Other fields should be much shorter, but we'll enforce limits for safety
        MAX_CELL_LENGTH = 49990  # Safety margin below 50,000 limit
        rows = []
        row_errors = []
        
        def safe_truncate_field(field, max_len, suffix="..."):
            """
            Safely truncate a field to max_len characters without breaking Unicode.
            Returns truncated string with suffix if needed.
            """
            if not field:
                return field
            
            # Convert to list of characters to handle multi-byte Unicode properly
            field_chars = list(field)
            suffix_len = len(suffix)
            max_content_len = max_len - suffix_len
            
            if len(field_chars) > max_len:
                # Truncate and add suffix
                truncated = ''.join(field_chars[:max_content_len]) + suffix
                return truncated[:max_len]  # Final safety check
            return field
        
        for idx, email in enumerate(new_emails):
            try:
                # Email content is already limited to 10,000 chars in parser
                # But add final safety check
                content = safe_truncate_field(email['content'], MAX_CELL_LENGTH, "...[TRUNCATED]")
                from_addr = safe_truncate_field(email['from'], MAX_CELL_LENGTH)
                subject = safe_truncate_field(email['subject'], MAX_CELL_LENGTH)
                date_str = safe_truncate_field(email['date'], MAX_CELL_LENGTH)
                
                # Final safety check: ensure no field exceeds 50,000 characters
                # This handles edge cases where multi-byte characters might cause issues
                row = [
                    from_addr[:50000] if len(from_addr) <= 50000 else from_addr[:49987] + "...",
                    subject[:50000] if len(subject) <= 50000 else subject[:49987] + "...",
                    date_str[:50000] if len(date_str) <= 50000 else date_str[:49987] + "...",
                    content[:50000] if len(content) <= 50000 else content[:49987] + "...[TRUNCATED]"
                ]
                
                # Verify all fields are within limit
                for i, field in enumerate(row):
                    if len(field) > 50000:
                        logger.error(f"Row {idx}, Field {i} still exceeds 50000 chars: {len(field)}")
                        # Force truncate
                        field_chars = list(field)
                        row[i] = ''.join(field_chars[:49987]) + "..."
                
                rows.append(row)
                
            except Exception as e:
                # Log error but continue processing other emails
                error_msg = f"Error preparing row for email {email.get('message_id', 'unknown')}: {e}"
                logger.error(error_msg)
                row_errors.append(error_msg)
                # Skip this email but continue with others
                continue
        
        if row_errors:
            logger.warning(f"Encountered {len(row_errors)} error(s) while preparing rows, but continuing...")
        
        if not rows:
            logger.warning("No valid rows to append after processing")
            return
        
        # Append to sheet with error handling
        logger.info(f"Appending {len(rows)} new email(s) to Google Sheet...")
        try:
            sheets_service.append_rows(rows, SHEET_NAME)
        except Exception as e:
            logger.error(f"Error appending rows to sheet: {e}")
            # Try appending in smaller batches to isolate problematic rows
            logger.info("Attempting to append rows in smaller batches...")
            batch_size = 50
            successful_batches = 0
            failed_batches = 0
            
            for i in range(0, len(rows), batch_size):
                batch = rows[i:i + batch_size]
                try:
                    sheets_service.append_rows(batch, SHEET_NAME)
                    successful_batches += 1
                    logger.debug(f"Successfully appended batch {i // batch_size + 1}")
                except Exception as batch_error:
                    failed_batches += 1
                    logger.warning(f"Failed to append batch {i // batch_size + 1}: {batch_error}")
                    # Try individual rows if batch fails
                    for j, row in enumerate(batch):
                        try:
                            sheets_service.append_rows([row], SHEET_NAME)
                            logger.debug(f"Successfully appended individual row {i + j + 1}")
                        except Exception as row_error:
                            logger.error(f"Failed to append row {i + j + 1}: {row_error}")
                            # Skip this row and continue
                            continue
            
            if successful_batches > 0 or failed_batches == 0:
                logger.info(f"Successfully appended {successful_batches} batch(es) to sheet")
            else:
                logger.error("Failed to append any rows to sheet")
                raise
        
        # Mark emails as read in Gmail
        logger.info(f"Marking {len(processed_ids)} email(s) as read...")
        gmail_service.mark_multiple_as_read(processed_ids)
        
        # Update state
        logger.info("Updating state...")
        state_manager.mark_multiple_processed(processed_ids)
        
        logger.info("=" * 60)
        logger.info(f"Successfully processed {len(new_emails)} email(s)")
        logger.info("=" * 60)
        
    except FileNotFoundError as e:
        logger.error(f"Configuration error: {e}")
        logger.error("Please ensure credentials.json is in the credentials/ directory")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()
