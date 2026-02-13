"""
Sheets Handler - Manages all Google Sheets operations
"""
import gspread
from datetime import datetime


def get_sheets_client(credentials):
    """Create and return gspread client"""
    return gspread.authorize(credentials)


def open_spreadsheet(gc, spreadsheet_id):
    """Open spreadsheet by ID"""
    return gc.open_by_key(spreadsheet_id)


def get_worksheet(sheet, worksheet_name):
    """Get a specific worksheet"""
    return sheet.worksheet(worksheet_name)


def get_all_tickets(worksheet):
    """
    Get all existing tickets from the sheet
    Returns: dict mapping thread_id to row_number
    """
    rows = worksheet.get_all_values()
    thread_map = {}
    
    for i, row in enumerate(rows[1:], start=2):  # Skip header, start from row 2
        if len(row) > 1 and row[1]:  # Check if thread_id exists
            thread_map[row[1]] = i
    
    return thread_map


def get_next_ticket_id(sheet):
    """
    Generate next ticket ID
    Reads from Ticket_Config sheet and increments counter
    """
    config_sheet = sheet.worksheet("Ticket_Config")
    last_ticket = int(config_sheet.acell("B1").value or 0)
    next_ticket = last_ticket + 1
    
    # Update counter
    config_sheet.update(range_name="B1", values=[[next_ticket]])
    
    return f"TCK-{next_ticket:06d}"


def create_ticket_row(ticket_id, thread_id, from_email, subject, status, new_sender=False):
    """
    Create a ticket row for the spreadsheet
    Returns: list of values for the row
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    link = f'=HYPERLINK("https://mail.google.com/mail/u/0/#inbox/{thread_id}","Open Mail")'
    
    return [
        ticket_id,
        thread_id,
        timestamp,
        from_email,
        subject,
        status,
        "Yes" if new_sender else "No",  # Dynamic new sender detection
        link,
    ]


def add_new_ticket(worksheet, row_data):
    """Add a new ticket row to the sheet"""
    worksheet.append_row(row_data, value_input_option="USER_ENTERED")


def update_existing_ticket(worksheet, row_number, row_data):
    """Update an existing ticket row"""
    range_name = f"A{row_number}:H{row_number}"
    worksheet.update(range_name, [row_data], value_input_option="USER_ENTERED")


def get_ticket_data(worksheet, row_number):
    """Get ticket data from a specific row"""
    return worksheet.row_values(row_number)


# ================= SYNC STATE MANAGEMENT =================
def initialize_state_sheets(sheet):
    """
    Initialize Sync_State and Thread_State sheets if they don't exist
    """
    # Initialize Sync_State
    try:
        sheet.worksheet("Sync_State")
        print("📊 Sync_State sheet already exists")
    except:
        sync_sheet = sheet.add_worksheet(title="Sync_State", rows=10, cols=2)
        sync_sheet.update("A1", [["Last Sync"]])
        print("✅ Created Sync_State sheet")
    
    # Initialize Thread_State
    try:
        sheet.worksheet("Thread_State")
        print("📊 Thread_State sheet already exists")
    except:
        thread_sheet = sheet.add_worksheet(title="Thread_State", rows=1000, cols=2)
        thread_sheet.update("A1", [["Thread ID", "Last Processed Timestamp"]])
        print("✅ Created Thread_State sheet")


def get_last_sync_timestamp(sheet):
    """
    Get last sync timestamp from Sync_State sheet
    Returns: timestamp (int) or None
    """
    try:
        sync_sheet = sheet.worksheet("Sync_State")
        value = sync_sheet.acell("B1").value
        return int(value) if value else None
    except:
        # Sheet doesn't exist or no value
        return None


def save_last_sync_timestamp(sheet, timestamp):
    """
    Save last sync timestamp to Sync_State sheet
    """
    try:
        sync_sheet = sheet.worksheet("Sync_State")
    except:
        # Create sheet if it doesn't exist
        sync_sheet = sheet.add_worksheet(title="Sync_State", rows=10, cols=2)
        # Set header
        sync_sheet.update("A1", [["Last Sync"]])
    
    # Update timestamp
    sync_sheet.update("B1", [[timestamp]])


# ================= THREAD STATE MANAGEMENT =================
def load_thread_state_from_sheet(sheet):
    """
    Load thread processing state from Thread_State sheet
    Returns: dict mapping thread_id to timestamp
    """
    state = {}
    try:
        thread_sheet = sheet.worksheet("Thread_State")
        rows = thread_sheet.get_all_values()
        
        # Skip header row
        for row in rows[1:]:
            if len(row) >= 2 and row[0]:
                state[row[0]] = int(row[1])
    except:
        # Sheet doesn't exist yet
        pass
    
    return state


def save_thread_state_to_sheet(sheet, state):
    """
    Save thread processing state to Thread_State sheet
    """
    try:
        thread_sheet = sheet.worksheet("Thread_State")
    except:
        # Create sheet if it doesn't exist
        thread_sheet = sheet.add_worksheet(title="Thread_State", rows=1000, cols=2)
    
    # Convert state dict to list of lists
    data = [["Thread ID", "Last Processed Timestamp"]]
    for tid, ts in state.items():
        data.append([tid, ts])
    
    # Clear and update
    thread_sheet.clear()
    thread_sheet.update("A1", data)