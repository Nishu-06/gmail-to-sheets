# Gmail to Sheets Automation

# **Author:** Nishu Kumari

A production-ready Python 3 automation system that connects to Gmail API and Google Sheets API to automatically log qualifying emails into a Google Sheet. The system uses OAuth 2.0 authentication, implements robust duplicate prevention, and includes state persistence to ensure emails are only processed once.

---

## Table of Contents

1. [High-Level Architecture](#high-level-architecture)
2. [Features](#features)
3. [Setup Instructions](#setup-instructions)
4. [OAuth Flow Explanation](#oauth-flow-explanation)
5. [Duplicate Prevention Logic](#duplicate-prevention-logic)
6. [State Persistence Method](#state-persistence-method)
7. [Challenges and Solutions](#challenges-and-solutions)
8. [Limitations](#limitations)
9. [Proof of Execution](#proof-of-execution)
10. [Future Enhancements](#future-enhancements)

---

## High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        User Execution                        │
│                      (python main.py)                        │
└───────────────────────────┬─────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────┐
│                         main.py                             │
│  • Orchestrates the workflow                                │
│  • Manages state persistence                                │
│  • Coordinates services                                      │
└───────────────┬───────────────────────────┬─────────────────┘
                │                           │
                ▼                           ▼
┌───────────────────────────┐   ┌───────────────────────────┐
│     gmail_service.py      │   │    sheets_service.py      │
│  • OAuth 2.0 auth         │   │  • OAuth 2.0 auth         │
│  • Fetch unread emails    │   │  • Create/ensure sheet    │
│  • Get email details      │   │  • Append rows            │
│  • Mark emails as read    │   │  • Retry logic            │
│  • Retry logic            │   │                           │
└───────────────┬───────────┘   └───────────────────────────┘
                │
                ▼
┌───────────────────────────┐
│     email_parser.py       │
│  • Parse email headers    │
│  • Extract body text      │
│  • HTML → plain text      │
│  • Apply filters          │
└───────────────────────────┘
                │
                ▼
┌───────────────────────────┐
│      State Manager        │
│  • Track processed IDs    │
│  • Persist to state.json  │
│  • Prevent duplicates     │
└───────────────────────────┘
```

**Data Flow:**
1. **Authentication**: OAuth 2.0 flow authenticates user with Google APIs
2. **Email Fetching**: Gmail API retrieves unread emails from inbox
3. **Parsing**: Email parser extracts sender, subject, date, and body (with HTML conversion)
4. **Filtering**: Subject-based and no-reply filters are applied
5. **Duplicate Check**: State manager checks if email was already processed
6. **Sheet Update**: New emails are appended to Google Sheet
7. **State Update**: Processed email IDs are saved to state.json
8. **Mark as Read**: Emails are marked as read in Gmail

---

## Features

### Core Features
- ✅ Gmail API integration (OAuth 2.0)
- ✅ Google Sheets API integration (OAuth 2.0)
- ✅ Fetches unread emails from inbox only
- ✅ Marks emails as read after processing
- ✅ Robust duplicate prevention
- ✅ State persistence between runs

### Bonus Features Implemented
1. **Subject-based Filtering**: Filter emails by keyword in subject line
2. **HTML to Plain Text Conversion**: Converts HTML emails to readable plain text
3. **Comprehensive Logging**: Timestamped logs to both file and console
4. **Retry Logic**: Exponential backoff retry for API failures (rate limits, server errors)

### Future-Ready Architecture
- Easy to add: Last 24 hours filter
- Easy to add: Email labels column
- Easy to add: Exclude no-reply emails (already implemented as config option)

---

## Setup Instructions

### Prerequisites
- Python 3.7 or higher (or Docker)
- Google account with Gmail
- Google Cloud Project with APIs enabled

### Step 1: Google Cloud Console Setup

1. **Create a Google Cloud Project**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one

2. **Enable Required APIs**
   - Navigate to "APIs & Services" > "Library"
   - Enable the following APIs:
     - Gmail API
     - Google Sheets API

3. **Create OAuth 2.0 Credentials**
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - If prompted, configure OAuth consent screen:
     - User Type: External (or Internal if using Google Workspace)
     - App name: "Gmail to Sheets Automation"
     - Scopes: Add the following:
       - `https://www.googleapis.com/auth/gmail.readonly`
       - `https://www.googleapis.com/auth/gmail.modify`
       - `https://www.googleapis.com/auth/spreadsheets`
     - Add your email as a test user
   - Application type: "Desktop app"
   - Name: "Gmail to Sheets Client"
   - Click "Create"
   - Download the JSON file and save it as `credentials/credentials.json`

### Step 2: Create Google Sheet

1. Create a new Google Sheet
2. Note the Spreadsheet ID from the URL:
   ```
   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
   ```
   Copy the long string between `/d/` and `/edit`
3. Share the sheet with the Google account you'll authenticate with (or use the same account)

### Step 3: Install Dependencies

**Option A: Local Python Installation**
```bash
pip install -r requirements.txt
```

**Option B: Docker (Recommended for Production)**
```bash
docker build -t gmail-to-sheets .
```

### Step 4: Configure the Application

**Set Spreadsheet ID** (choose one method):

**Option A: Environment Variable (Recommended)**
```bash
# Windows (CMD)
set SPREADSHEET_ID=your_spreadsheet_id_here

# Windows (PowerShell)
$env:SPREADSHEET_ID="your_spreadsheet_id_here"

# Linux/Mac
export SPREADSHEET_ID=your_spreadsheet_id_here
```

**Option B: Edit config.py**
```python
SPREADSHEET_ID = 'your_spreadsheet_id_here'
```

**Optional Configuration** (via environment variables or `config.py`):
- `SHEET_NAME`: Name of the sheet tab (default: "Emails")
- `SUBJECT_FILTER`: Filter emails by subject keyword (empty = no filter)
- `EXCLUDE_NO_REPLY`: Set to "true" to exclude no-reply emails
- `LOG_LEVEL`: Logging level (DEBUG, INFO, WARNING, ERROR)

### Step 5: Run the Application

**Option A: Local Python**
```bash
python src/main.py
```

**Option B: Docker Compose**
```bash
# Create .env file
echo SPREADSHEET_ID=your_spreadsheet_id_here > .env

# Run
docker-compose up
```

**Option C: Docker Direct**
```bash
docker run --rm \
  -e SPREADSHEET_ID=your_spreadsheet_id_here \
  -v $(pwd)/credentials:/app/credentials:ro \
  -v $(pwd)/state.json:/app/state.json \
  -v $(pwd)/logs:/app/logs \
  gmail-to-sheets
```

**First Run:**
- A browser window will open for OAuth authentication
- Sign in with your Google account
- Grant permissions for Gmail and Sheets access
- A `token.json` file will be created in `credentials/` directory

**Subsequent Runs:**
- Uses saved `token.json` (automatically refreshes if expired)
- Processes only new unread emails
- Appends to Google Sheet

---

## OAuth Flow Explanation

The application uses **OAuth 2.0 Authorization Code Flow** with the following steps:

1. **Initial Authentication (First Run)**:
   - Application reads `credentials.json` (OAuth client ID and secret)
   - Opens a local web server on a random port
   - Redirects user to Google OAuth consent screen
   - User grants permissions for Gmail and Sheets
   - Google redirects back with an authorization code
   - Application exchanges code for access token and refresh token
   - Tokens are saved to `credentials/token.json`

2. **Subsequent Runs**:
   - Application loads `token.json`
   - If token is expired, uses refresh token to get a new access token
   - If refresh token is invalid, triggers full OAuth flow again

3. **Token Management**:
   - Access tokens are short-lived (typically 1 hour)
   - Refresh tokens are long-lived and stored securely
   - Token refresh happens automatically when needed

**Why OAuth 2.0 (not Service Account)?**
- Service accounts cannot access user Gmail inbox directly
- OAuth 2.0 allows the application to act on behalf of the user
- Provides better security with user consent
- Supports token refresh for long-running automation

---

## Duplicate Prevention Logic

The system uses a **two-layer duplicate prevention strategy**:

### Layer 1: State Persistence (Primary)
- **File**: `state.json` stores all processed Gmail message IDs
- **Process**:
  1. Before processing an email, check if its message ID exists in `state.json`
  2. If exists → skip (already processed)
  3. If not exists → process and add to state
- **Advantages**:
  - Fast lookup (in-memory set)
  - Persistent across script runs
  - Reliable (Gmail message IDs are unique and permanent)

### Layer 2: Gmail Query (Secondary)
- Only fetches unread emails (`is:unread in:inbox`)
- After processing, emails are marked as read
- Prevents the same email from being fetched again

### Why This Approach?

**Gmail Message IDs are ideal for duplicate tracking because:**
- Unique per email (even if forwarded/replied)
- Permanent (don't change)
- Available immediately (no need to parse email first)

**Alternative approaches considered:**
- **Content hash**: Could fail if email is modified
- **Date + Subject + From**: Could have false positives
- **Sheet lookup**: Slow for large sheets, requires API calls

**State file structure:**
```json
{
  "processed_message_ids": [
    "18c1234567890abcdef",
    "18c0987654321fedcba"
  ],
  "last_updated": "2024-01-15T10:30:00"
}
```

---

## State Persistence Method

### Implementation Details

**File**: `state.json` (in project root)

**Storage Format**: JSON
```json
{
  "processed_message_ids": ["id1", "id2", ...],
  "last_updated": "2024-01-15T10:30:00"
}
```

**Lifecycle**:
1. **Load on startup**: `StateManager` reads `state.json` into memory (set of message IDs)
2. **Check during processing**: Each email's message ID is checked against the set
3. **Update after processing**: New message IDs are added to the set and file is saved
4. **Persist across runs**: State survives script restarts

**Why JSON?**
- Human-readable (for debugging)
- Easy to parse and modify
- No external dependencies
- Suitable for small to medium datasets (thousands of IDs)

**Scalability Considerations**:
- For millions of emails, consider:
  - Database (SQLite, PostgreSQL)
  - Bloom filter for memory efficiency
  - Periodic cleanup of old IDs (if not needed)

**Security**:
- `state.json` is excluded from git (via `.gitignore`)
- Contains only message IDs (no sensitive data)
- Can be safely deleted to reprocess all emails

---

## Challenges and Solutions

### Challenge: Google Sheets 50,000 Character Cell Limit

**Problem**: Some emails have extremely long content (newsletters, automated reports, etc.) that exceed Google Sheets' 50,000 character limit per cell. When appending such emails, the entire batch would fail with `HttpError 400`.

**Solution**: Implemented a multi-layered approach:

1. **Hard Limit in Parser**: Enforced a 10,000 character limit on email body during parsing (well below the 50,000 limit)
   - Uses Unicode-safe character list truncation to avoid breaking multi-byte characters
   - Adds "...[TRUNCATED]" suffix when content is cut
   - Logs warnings for truncated emails

2. **Safe Truncation Function**: Created a `safe_truncate_field()` function that:
   - Converts strings to character lists for proper Unicode handling
   - Truncates all fields (From, Subject, Date, Content) to ensure compliance
   - Adds appropriate suffixes when truncation occurs

3. **Batch Resilience**: Modified append logic to:
   - Continue processing even if individual emails cause errors
   - Retry in smaller batches (50 rows) if full batch fails
   - Fall back to individual row appends if batch fails
   - Log errors but never fail the entire operation

**Result**: The script now successfully processes all emails, truncating only the extremely long ones, and always completes successfully.

**Code Locations**: 
- `src/email_parser.py` - Email body truncation (lines 190-210)
- `src/main.py` - Row preparation and batch error handling (lines 210-280)

---

## Limitations

1. **State File Growth**: 
   - `state.json` grows indefinitely with processed email IDs
   - For very long-running systems, consider periodic cleanup
   - **Mitigation**: Can add date-based cleanup or move to database

2. **Gmail API Quotas**:
   - Gmail API has daily quota limits (1 billion quota units per day)
   - Each email fetch consumes quota units
   - **Mitigation**: Retry logic handles rate limits; batch processing is efficient

3. **Sheet Size Limits**:
   - Google Sheets has a limit of 10 million cells
   - With 4 columns, that's ~2.5 million emails
   - **Mitigation**: Can implement sheet rotation or archiving

4. **OAuth Token Expiry**:
   - Refresh tokens can expire if not used for 6 months
   - **Mitigation**: Regular script runs keep tokens fresh

5. **No Real-time Processing**:
   - Script must be run manually or via cron/scheduler
   - **Mitigation**: Can be scheduled with cron (Linux/Mac) or Task Scheduler (Windows)

6. **Single User Only**:
   - Currently processes emails for one authenticated user
   - **Mitigation**: Can be extended to support multiple users with separate state files

7. **No Attachment Handling**:
   - Attachments are not processed or logged
   - **Mitigation**: Can be added as future enhancement

---

## Proof of Execution

### How to Verify the System Works

1. **Check Logs**:
   ```bash
   # View the log file
   cat gmail_to_sheets.log
   # or
   type gmail_to_sheets.log  # Windows
   ```
   Look for:
   - "Gmail to Sheets Automation - Starting"
   - "Found X unread email(s)"
   - "Successfully processed X email(s)"

2. **Check Google Sheet**:
   - Open your Google Sheet
   - Verify new rows were added with columns: From, Subject, Date, Content
   - Check that emails are marked as read in Gmail

3. **Check State File**:
   ```bash
   cat state.json
   ```
   Should contain a list of processed message IDs

4. **Run Again (Duplicate Test)**:
   - Run the script again immediately
   - Should see: "No unread emails found" or "No new emails to add to sheet"
   - No duplicate rows should be added

5. **Test with New Email**:
   - Send yourself a test email
   - Run the script
   - Verify the email appears in the sheet
   - Verify the email is marked as read in Gmail

### Expected Output

```
2024-01-15 10:30:00 - __main__ - INFO - ============================================================
2024-01-15 10:30:00 - __main__ - INFO - Gmail to Sheets Automation - Starting
2024-01-15 10:30:00 - __main__ - INFO - ============================================================
2024-01-15 10:30:01 - __main__ - INFO - Initializing Gmail service...
2024-01-15 10:30:02 - gmail_service - INFO - Gmail API service initialized
2024-01-15 10:30:02 - __main__ - INFO - Initializing Sheets service...
2024-01-15 10:30:02 - sheets_service - INFO - Google Sheets API service initialized
2024-01-15 10:30:02 - __main__ - INFO - Loading state...
2024-01-15 10:30:02 - __main__ - INFO - Loaded 5 processed email IDs from state
2024-01-15 10:30:02 - __main__ - INFO - Fetching unread emails from Gmail...
2024-01-15 10:30:03 - gmail_service - INFO - Found 3 unread emails
2024-01-15 10:30:03 - __main__ - INFO - Found 3 unread email(s)
2024-01-15 10:30:04 - __main__ - INFO - Parsed email: Test Email Subject...
2024-01-15 10:30:05 - __main__ - INFO - Appending 3 new email(s) to Google Sheet...
2024-01-15 10:30:06 - sheets_service - INFO - Successfully appended 3 row(s)
2024-01-15 10:30:06 - __main__ - INFO - Marking 3 email(s) as read...
2024-01-15 10:30:07 - __main__ - INFO - ============================================================
2024-01-15 10:30:07 - __main__ - INFO - Successfully processed 3 email(s)
2024-01-15 10:30:07 - __main__ - INFO - ============================================================
```

---

## Future Enhancements

The codebase is structured to easily add the following features:

### 1. Last 24 Hours Filter
**Location**: `config.py` and `src/gmail_service.py`
```python
# In config.py
LAST_24_HOURS_ONLY = True

# In gmail_service.py
if LAST_24_HOURS_ONLY:
    from datetime import datetime, timedelta
    cutoff = (datetime.now() - timedelta(hours=24)).strftime('%Y/%m/%d')
    query += f' after:{cutoff}'
```

### 2. Email Labels Column
**Location**: `src/email_parser.py` and `src/sheets_service.py`
```python
# In email_parser.py
labels = message.get('labelIds', [])
parsed['labels'] = ', '.join(labels)

# In sheets_service.py
expected_headers = ['From', 'Subject', 'Date', 'Content', 'Labels']
```

### 3. Exclude No-Reply Emails
**Already implemented!** Just set in `config.py`:
```python
EXCLUDE_NO_REPLY = True
```

### Additional Ideas
- Email attachment detection and logging
- Multiple sheet support (one sheet per label)
- Webhook integration for real-time processing
- Email threading support
- Custom column mappings
- Scheduled runs with cron/Task Scheduler integration

---

## Project Structure

```
gmail-to-sheets/
│
├── src/
│   ├── __init__.py
│   ├── gmail_service.py      # Gmail API integration
│   ├── sheets_service.py     # Google Sheets API integration
│   ├── email_parser.py       # Email parsing and HTML conversion
│   └── main.py               # Main orchestration logic
│
├── credentials/
│   ├── credentials.json      # OAuth credentials (NOT committed)
│   ├── token.json            # OAuth tokens (NOT committed)
│   └── credentials.json.example  # Template file
│
├── config.py                 # Configuration settings
├── requirements.txt          # Python dependencies
├── Dockerfile                # Docker image definition
├── docker-compose.yml        # Docker Compose configuration
├── .dockerignore             # Docker ignore rules
├── README.md                 # This file
├── .gitignore               # Git ignore rules
├── state.json               # Processed email IDs (NOT committed)
└── gmail_to_sheets.log      # Application logs
```

---

## Troubleshooting

### "Credentials file not found"
- Ensure `credentials.json` is in `credentials/` directory
- Download from Google Cloud Console > Credentials

### "SPREADSHEET_ID not set"
- Set environment variable or edit `config.py`
- Get ID from Google Sheet URL

### "Permission denied" errors
- Ensure OAuth consent screen is configured
- Add your email as a test user
- Re-authenticate by deleting `token.json`

### "No unread emails found"
- Check Gmail inbox for unread emails
- Verify query in `config.py` matches your needs

### Duplicate rows appearing
- Check that `state.json` is being created and updated
- Ensure script has write permissions in project directory

---

## Docker Deployment

### Quick Start with Docker

1. **Build the image:**
   ```bash
   docker build -t gmail-to-sheets .
   ```

2. **Run with Docker Compose (Recommended):**
   ```bash
   # Create .env file
   echo SPREADSHEET_ID=your_spreadsheet_id_here > .env
   
   # Run
   docker-compose up
   ```

3. **Run directly with Docker:**
   ```bash
   docker run --rm \
     -e SPREADSHEET_ID=your_spreadsheet_id_here \
     -e SHEET_NAME=Emails \
     -v $(pwd)/credentials:/app/credentials:ro \
     -v $(pwd)/state.json:/app/state.json \
     -v $(pwd)/logs:/app/logs \
     gmail-to-sheets
   ```

### Docker Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `SPREADSHEET_ID` | Google Sheets ID (required) | - |
| `SHEET_NAME` | Name of the sheet tab | `Emails` |
| `LOG_LEVEL` | Logging level | `INFO` |
| `SUBJECT_FILTER` | Filter emails by subject keyword | - |
| `EXCLUDE_NO_REPLY` | Exclude no-reply emails | `false` |

### Docker Volume Mounts

- **`credentials/`**: Read-only mount for OAuth credentials and tokens
- **`state.json`**: Persistent state file for tracking processed emails
- **`logs/`**: Log files directory

**Note**: Credentials are mounted as volumes (not baked into image) for security.

---

## License

This project is provided as-is for educational and automation purposes.

---

**Version:** 1.0.0
