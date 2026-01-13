"""
Email parsing utilities.
Handles extraction of email data including HTML to plain text conversion.
"""

import base64
import re
import logging
import sys
from pathlib import Path
from typing import Dict, Optional
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser
import quopri

# Add project root to Python path
project_root = Path(__file__).parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from config import SUBJECT_FILTER, EXCLUDE_NO_REPLY

logger = logging.getLogger(__name__)


class HTMLToTextParser(HTMLParser):
    """HTML parser that extracts plain text from HTML content."""
    
    def __init__(self):
        super().__init__()
        self.text = []
        self.skip_tags = {'script', 'style', 'head', 'meta'}
        self.current_tag = None
    
    def handle_starttag(self, tag, attrs):
        self.current_tag = tag.lower()
        if tag.lower() == 'br':
            self.text.append('\n')
    
    def handle_endtag(self, tag):
        if tag.lower() in ('p', 'div', 'li'):
            self.text.append('\n')
        self.current_tag = None
    
    def handle_data(self, data):
        if self.current_tag not in self.skip_tags:
            self.text.append(data)
    
    def get_text(self) -> str:
        """Get extracted plain text."""
        text = ''.join(self.text)
        # Clean up excessive whitespace
        text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)
        text = re.sub(r'[ \t]+', ' ', text)
        return text.strip()


def decode_base64(data: str) -> str:
    """
    Decode base64 encoded email content.
    
    Args:
        data: Base64 encoded string
        
    Returns:
        Decoded string
    """
    try:
        decoded = base64.urlsafe_b64decode(data)
        return decoded.decode('utf-8', errors='ignore')
    except Exception as e:
        logger.warning(f"Failed to decode base64: {e}")
        return ""


def extract_email_body(parts: list, body_text: str = "") -> str:
    """
    Recursively extract email body from multipart message.
    Handles both plain text and HTML content.
    
    Args:
        parts: List of message parts
        body_text: Accumulated body text
        
    Returns:
        Plain text email body
    """
    if not parts:
        return body_text
    
    for part in parts:
        mime_type = part.get('mimeType', '')
        body = part.get('body', {})
        data = body.get('data', '')
        
        if mime_type == 'text/plain':
            if data:
                decoded = decode_base64(data)
                # Handle quoted-printable encoding
                try:
                    decoded = quopri.decodestring(decoded.encode()).decode('utf-8', errors='ignore')
                except:
                    pass
                return decoded
                
        elif mime_type == 'text/html':
            if data:
                html_content = decode_base64(data)
                # Convert HTML to plain text
                parser = HTMLToTextParser()
                parser.feed(html_content)
                html_text = parser.get_text()
                # Prefer plain text, but use HTML if no plain text available
                if not body_text:
                    return html_text
                
        # Recursively check nested parts
        if 'parts' in part:
            nested_text = extract_email_body(part['parts'], body_text)
            if nested_text:
                return nested_text
    
    return body_text


def parse_email_message(message: Dict) -> Optional[Dict]:
    """
    Parse a Gmail message object into structured data.
    
    Args:
        message: Full Gmail message object
        
    Returns:
        Dictionary with 'from', 'subject', 'date', 'content', 'message_id'
        Returns None if email doesn't meet filter criteria
    """
    try:
        payload = message.get('payload', {})
        headers = payload.get('headers', [])
        
        # Extract headers
        header_dict = {h['name'].lower(): h['value'] for h in headers}
        
        from_email = header_dict.get('from', 'Unknown')
        subject = header_dict.get('subject', '(No Subject)')
        date_str = header_dict.get('date', '')
        
        # Extract message ID for duplicate tracking
        message_id = message.get('id', '')
        
        # Apply filters
        if EXCLUDE_NO_REPLY:
            if 'no-reply' in from_email.lower() or 'noreply' in from_email.lower():
                logger.debug(f"Skipping no-reply email: {from_email}")
                return None
        
        if SUBJECT_FILTER:
            if SUBJECT_FILTER.lower() not in subject.lower():
                logger.debug(f"Skipping email (subject filter): {subject}")
                return None
        
        # Parse date
        try:
            if date_str:
                dt = parsedate_to_datetime(date_str)
                date_iso = dt.isoformat()
            else:
                # Fallback to internal date
                internal_date = message.get('internalDate')
                if internal_date:
                    from datetime import datetime
                    dt = datetime.fromtimestamp(int(internal_date) / 1000)
                    date_iso = dt.isoformat()
                else:
                    date_iso = ""
        except Exception as e:
            logger.warning(f"Failed to parse date '{date_str}': {e}")
            date_iso = date_str
        
        # Extract body
        body_text = ""
        if 'parts' in payload:
            body_text = extract_email_body(payload['parts'])
        else:
            # Simple message without multipart
            body = payload.get('body', {})
            data = body.get('data', '')
            if data:
                body_text = decode_base64(data)
        
        # Clean up body text
        if body_text:
            # Remove excessive whitespace
            body_text = re.sub(r'\s+', ' ', body_text)
            body_text = body_text.strip()
        
        # Enforce 10,000 character limit on email body with safe Unicode truncation
        MAX_EMAIL_BODY_LENGTH = 10000
        TRUNCATE_SUFFIX = "...[TRUNCATED]"
        SUFFIX_LENGTH = len(TRUNCATE_SUFFIX)
        MAX_CONTENT_LENGTH = MAX_EMAIL_BODY_LENGTH - SUFFIX_LENGTH
        
        content = body_text or '(No content)'
        was_truncated = False
        
        if len(content) > MAX_EMAIL_BODY_LENGTH:
            # Truncate safely without breaking Unicode characters
            # Convert to list of characters to handle multi-byte Unicode properly
            content_chars = list(content)
            if len(content_chars) > MAX_CONTENT_LENGTH:
                content = ''.join(content_chars[:MAX_CONTENT_LENGTH]) + TRUNCATE_SUFFIX
                was_truncated = True
                logger.warning(
                    f"Truncated email body from {len(body_text)} to {MAX_EMAIL_BODY_LENGTH} chars "
                    f"(Subject: {subject[:50]})"
                )
        
        return {
            'from': from_email,
            'subject': subject,
            'date': date_iso,
            'content': content,
            'message_id': message_id,
            'was_truncated': was_truncated
        }
        
    except Exception as e:
        logger.error(f"Error parsing email message: {e}")
        return None
