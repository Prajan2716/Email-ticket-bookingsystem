# Gmail Ticket System - Modular Version

## File Structure

```
├── gmail_handler.py      # Gmail API operations
├── sheets_handler.py     # Google Sheets operations  
├── main.py              # Core sync logic
├── app.py               # Flask web application
├── credentials.json     # Google API credentials (you provide)
├── token.pickle         # Auto-generated after first auth
├── last_sync.txt        # Auto-generated sync state
└── thread_state.txt     # Auto-generated thread state
```

## Files Explained

### 1. `gmail_handler.py`
**Purpose:** All Gmail operations
- Authenticate with Gmail API
- Fetch email threads
- Get message details
- Manage Gmail labels
- Extract email addresses

### 2. `sheets_handler.py`
**Purpose:** All Google Sheets operations
- Connect to Google Sheets
- Read/write ticket data
- Generate ticket IDs
- Create/update ticket rows

### 3. `main.py`
**Purpose:** Core business logic
- Combines Gmail + Sheets handlers
- Main sync function: `sync_mail_to_sheet()`
- Process threads and update tickets
- Manage status (Awaiting admin/customer reply)
- Can run standalone or be imported

### 4. `app.py`
**Purpose:** Flask web interface
- Import and call `sync_mail_to_sheet()` from main.py
- Provides REST API endpoints
- Web interface for controlling sync
- Background thread for auto-sync

## How to Use

### Option 1: Run Standalone (No Flask)
```bash
python main.py
```
This will run continuous sync every 5 seconds in the terminal.

### Option 2: Run with Flask
```bash
python app.py
```
Then access in browser:
- `http://localhost:5000/` - Home page
- `http://localhost:5000/start` - Start auto-sync
- `http://localhost:5000/stop` - Stop auto-sync  
- `http://localhost:5000/sync` - Manual sync
- `http://localhost:5000/status` - Check status

## Installation

1. Install dependencies:
```bash
pip install gspread google-auth-oauthlib google-api-python-client flask
```

2. Add your `credentials.json` file

3. Run for first time to authenticate:
```bash
python main.py
```

## Key Features

✅ **No functionality changes** - Same logic as before
✅ **Automated email removed** - No auto-replies sent
✅ **Modular code** - Easy to maintain
✅ **Flask ready** - Web interface available
✅ **Standalone capable** - Can run without Flask

## What Was Removed

❌ `send_auto_reply()` function
❌ `get_attachments_from_message()` function  
❌ All auto-reply email sending logic
❌ Attachment processing for replies

## What Remains

✅ Thread monitoring
✅ Ticket creation
✅ Status updates (Awaiting admin/customer reply)
✅ Gmail label management
✅ Google Sheets sync
✅ All core functionality

## Configuration

Edit these in `main.py`:
```python
SPREADSHEET_ID = "your-spreadsheet-id"
CHECK_INTERVAL = 5  # seconds between syncs
```

## Troubleshooting

**Authentication Error:**
- Delete `token.pickle` and re-authenticate

**Import Error:**
- Make sure all files are in the same directory
- Install required packages

**Sync Not Working:**
- Check `credentials.json` exists
- Verify spreadsheet ID is correct
- Check API permissions in Google Cloud Console
