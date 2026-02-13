"""
Main Sync Logic - Combines Gmail and Sheets operations
"""
import os
import time
from gmail_handler import (
    get_gmail_service,
    get_gmail_credentials,
    get_my_email,
    extract_email,
    get_or_create_label,
    fetch_all_threads,
    get_thread_details,
    get_last_message,
    update_thread_labels
)
from sheets_handler import (
    get_sheets_client,
    open_spreadsheet,
    get_worksheet,
    get_all_tickets,
    get_next_ticket_id,
    create_ticket_row,
    add_new_ticket,
    update_existing_ticket,
    initialize_state_sheets,
    get_last_sync_timestamp,
    save_last_sync_timestamp,
    load_thread_state_from_sheet,
    save_thread_state_to_sheet
)

# ================= CONFIG =================
SPREADSHEET_ID = "1F1STl7ubwviajSvu1LGfw6FGiDUQJbcZ-fI586pqwiE"
THREAD_STATE_FILE = "thread_state.txt"
SHEET_BACKUP_INTERVAL = 50  # Backup to sheet every 50 syncs
TICKET_MAP_REFRESH_INTERVAL = 20  # Refresh ticket map every 20 syncs

# Auto-close settings
AUTO_CLOSE_ENABLED = True  # Set to False to disable auto-close
AUTO_CLOSE_HOURS = 6  # Close tickets after N hours of no customer response
AUTO_CLOSE_ACTION = "close"  # Options: "close" (mark as closed) or "delete" (remove from sheet and trash email)

# Admin email addresses - Add ALL your email addresses here
ADMIN_EMAILS = [
    "support-ticketana@he5.in"
    # Add your email addresses here (lowercase)
    # Example:
    # "admin@example.com",
    # "support@example.com",
    # "sales@example.com",
]

# Global counters and cache
sync_counter = 0
cached_thread_map = None
last_ticket_map_refresh = 0
sheets_initialized = False

# New sender tracking
known_senders = set()
known_senders_loaded = False


# ================= FILE-BASED STATE MANAGEMENT =================
def load_thread_state_from_file():
    """Load thread processing state from file"""
    state = {}
    if os.path.exists(THREAD_STATE_FILE):
        with open(THREAD_STATE_FILE, "r") as f:
            for line in f:
                if "|" in line:
                    tid, ts = line.strip().split("|")
                    state[tid] = int(ts)
    return state


def save_thread_state_to_file(state):
    """Save thread processing state to file"""
    with open(THREAD_STATE_FILE, "w") as f:
        for tid, ts in state.items():
            f.write(f"{tid}|{ts}\n")


def is_admin_email(email):
    """Check if email belongs to admin"""
    email_lower = email.lower()
    return email_lower in ADMIN_EMAILS


def load_admin_emails_from_sheet(sheet):
    """
    Load admin emails from Admin emails sheet
    Returns: list of admin email addresses
    """
    try:
        admin_sheet = sheet.worksheet("Admin emails")
        rows = admin_sheet.get_all_values()
        
        admin_emails = []
        # Skip header row (row 1)
        for row in rows[1:]:
            if len(row) > 0 and row[0]:  # If column A has a value
                email = row[0].strip().lower()  # Clean and lowercase
                if email:  # If not empty after stripping
                    admin_emails.append(email)
        
        print(f"📧 Loaded {len(admin_emails)} admin emails from sheet: {', '.join(admin_emails)}")
        return admin_emails
    except Exception as e:
        print(f"⚠️ Could not load Admin emails sheet: {e}")
        print(f"⚠️ Using hardcoded ADMIN_EMAILS instead")
        return []


def load_known_senders(main_worksheet):
    """
    Load all unique sender emails from the sheet into a set
    Returns: set of email addresses
    """
    all_rows = main_worksheet.get_all_values()
    senders = set()
    
    # Column D (index 3) contains email addresses
    for row in all_rows[1:]:  # Skip header
        if len(row) > 3 and row[3]:
            senders.add(row[3].lower())
    
    print(f"📧 Loaded {len(senders)} known senders from sheet")
    return senders


def is_new_sender_cached(from_email):
    """
    Check if sender is new using cached set
    Returns: True if new, False if seen before
    """
    global known_senders
    return from_email.lower() not in known_senders


def check_and_close_stale_tickets(gmail, sheet, main_worksheet):
    """
    Check for tickets awaiting customer reply for more than AUTO_CLOSE_HOURS
    and close/delete them
    """
    if not AUTO_CLOSE_ENABLED:
        return
    
    print(f"\n🔍 Checking for stale tickets (>{AUTO_CLOSE_HOURS}h no customer response)...")
    
    # Get all tickets
    all_rows = main_worksheet.get_all_values()
    current_time = int(time.time())
    closed_count = 0
    deleted_count = 0
    
    for i, row in enumerate(all_rows[1:], start=2):  # Skip header
        if len(row) < 6:
            continue
            
        ticket_id = row[0]
        thread_id = row[1]
        timestamp_str = row[2]
        status = row[5]
        
        # Only check tickets awaiting customer reply
        if status != "Awaiting customer reply":
            continue
        
        # Parse timestamp
        try:
            from datetime import datetime
            ticket_time = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
            ticket_timestamp = int(ticket_time.timestamp())
        except:
            continue
        
        # Check if ticket is older than AUTO_CLOSE_HOURS
        hours_passed = (current_time - ticket_timestamp) / 3600
        
        if hours_passed >= AUTO_CLOSE_HOURS:
            if AUTO_CLOSE_ACTION == "delete":
                # Delete from sheet
                main_worksheet.delete_rows(i)
                print(f"   🗑️ Deleted ticket {ticket_id} (no response for {hours_passed:.1f}h)")
                
                # Trash the email thread
                try:
                    gmail.users().threads().trash(userId="me", id=thread_id).execute()
                    print(f"   📧 Trashed email thread {thread_id}")
                except Exception as e:
                    print(f"   ⚠️ Could not trash thread: {e}")
                
                deleted_count += 1
            else:  # "close"
                # Update status to Closed
                row[5] = "Closed - No customer response"
                main_worksheet.update(f"A{i}:H{i}", [row], value_input_option="USER_ENTERED")
                print(f"   ✅ Closed ticket {ticket_id} (no response for {hours_passed:.1f}h)")
                closed_count += 1
    
    if closed_count > 0 or deleted_count > 0:
        print(f"📊 Auto-close summary: {closed_count} closed, {deleted_count} deleted")


# ================= MAIN SYNC FUNCTION =================
def sync_mail_to_sheet():
    """
    Main synchronization function
    Syncs Gmail threads with Google Sheets ticket system
    """
    print("\n" + "="*50)
    print("Starting sync...")
    print("="*50)
    
    # Initialize Gmail
    gmail = get_gmail_service()
    my_email = get_my_email(gmail)
    
    # Auto-add authenticated email to admin list if not already there
    global ADMIN_EMAILS
    if my_email not in ADMIN_EMAILS:
        ADMIN_EMAILS.append(my_email)
    
    print(f"📧 Authenticated as: {my_email}")
    print(f"👥 Admin emails: {', '.join(ADMIN_EMAILS)}")

    # Initialize Sheets
    creds = get_gmail_credentials()
    gc = get_sheets_client(creds)
    sheet = open_spreadsheet(gc, SPREADSHEET_ID)
    main_worksheet = get_worksheet(sheet, "Email log")
    print(f"📊 Connected to spreadsheet")

    # Initialize state sheets ONCE (only on first run)
    global sheets_initialized
    if not sheets_initialized:
        initialize_state_sheets(sheet)
        sheets_initialized = True

    # Get or create labels
    admin_label = get_or_create_label(gmail, "Awaiting_Admin_Reply")
    cust_label = get_or_create_label(gmail, "Awaiting_Customer_Reply")
    print(f"🏷️ Labels configured")

    # Get existing tickets (cached - refresh periodically)
    global cached_thread_map, last_ticket_map_refresh, sync_counter
    sync_counter += 1
    
    if cached_thread_map is None or (sync_counter - last_ticket_map_refresh) >= TICKET_MAP_REFRESH_INTERVAL:
        cached_thread_map = get_all_tickets(main_worksheet)
        last_ticket_map_refresh = sync_counter
        print(f"📋 Refreshed ticket map: {len(cached_thread_map)} existing tickets")
    else:
        print(f"📋 Using cached ticket map: {len(cached_thread_map)} existing tickets")
    
    # Load known senders (once per run)
    global known_senders, known_senders_loaded
    if not known_senders_loaded:
        known_senders = load_known_senders(main_worksheet)
        known_senders_loaded = True
    
    # Load thread state from FILE (fast, no API calls)
    thread_state = load_thread_state_from_file()
    
    # Get last sync from sheet ONLY on first run, then use file
    if sync_counter == 1:
        last_sync = get_last_sync_timestamp(sheet)
        print(f"📊 Loaded initial sync timestamp from sheet")
    else:
        # After first run, we know it's current time - 5 seconds
        last_sync = int(time.time()) - 10  # Look back 10 seconds to be safe
    
    print(f"📊 Loaded state: {len(thread_state)} threads tracked (sync #{sync_counter})")
    
    # Build query
    query = f"after:{last_sync}" if last_sync else "newer_than:7d"
    
    # Fetch threads
    threads = fetch_all_threads(gmail, query)
    
    # CRITICAL: Deduplicate threads (Gmail sometimes returns duplicates)
    seen_thread_ids = set()
    unique_threads = []
    for thread in threads:
        tid = thread["id"]
        if tid not in seen_thread_ids:
            seen_thread_ids.add(tid)
            unique_threads.append(thread)
    
    if len(threads) != len(unique_threads):
        print(f"⚠️ Removed {len(threads) - len(unique_threads)} duplicate thread(s)")
    
    threads = unique_threads
    print(f"📬 Found {len(threads)} threads to process\n")

    # Process each thread
    for thread_info in threads:
        tid = thread_info["id"]
        
        print(f"\n{'='*60}")
        print(f"🔍 DEBUG: Examining thread {tid}")
        print(f"   In cached_thread_map: {tid in cached_thread_map}")
        print(f"   In thread_state: {tid in thread_state}")
        if tid in thread_state:
            print(f"   Thread_state timestamp: {thread_state[tid]}")
        print(f"{'='*60}")
        
        # Get full thread details
        thread = get_thread_details(gmail, tid)
        msg, headers = get_last_message(thread)
        
        if not msg:
            print(f"⏭️ Skipping thread {tid} - no messages")
            continue

        ts = int(msg["internalDate"]) // 1000
        
        # Skip if already processed
        if ts <= thread_state.get(tid, 0):
            print(f"⏭️ Skipping thread {tid} - already processed")
            continue

        from_email = extract_email(headers.get("From", ""))
        subject = headers.get("Subject", "No Subject")

        print(f"\n📨 Processing thread {tid}")
        print(f"   From: {from_email}")
        print(f"   Subject: {subject}")

        # Determine if new or existing ticket
        is_new_ticket = tid not in cached_thread_map
        
        print(f"   🎯 DEBUG: is_new_ticket = {is_new_ticket}")
        print(f"   🎯 DEBUG: cached_thread_map size = {len(cached_thread_map)}")
        
        # CRITICAL: If we've seen this thread in THIS sync already, skip it
        # This prevents duplicates when Gmail returns same thread multiple times
        if tid in thread_state and thread_state[tid] >= ts:
            print(f"   ⏭️ Skipping thread {tid} - already processed in this sync")
            continue
        
        # Skip NEW threads initiated by admin
        if is_new_ticket and is_admin_email(from_email):
            print(f"   ⏭️ Skipping - admin-initiated thread")
            thread_state[tid] = ts
            continue

        if not is_new_ticket:
            # Existing ticket - get ticket ID
            row_num = cached_thread_map[tid]
            ticket_data = main_worksheet.row_values(row_num)
            ticket_id = ticket_data[0]
        else:
            # New ticket - generate ticket ID
            
            # FINAL SAFETY CHECK: Re-check cache one more time
            # (in case another process created it)
            final_check_map = get_all_tickets(main_worksheet)
            if tid in final_check_map:
                print(f"   ⚠️ WARNING: Thread {tid} was just created by another process!")
                print(f"   ⏭️ Skipping to avoid duplicate")
                cached_thread_map = final_check_map  # Update cache
                continue
            
            ticket_id = get_next_ticket_id(sheet)
            print(f"   🎫 New ticket: {ticket_id}")
            print(f"   🆔 DEBUG: Full thread ID = {tid}")
            print(f"   🆔 DEBUG: Thread ID length = {len(tid)}")
            
            # CRITICAL: Mark as processed BEFORE creating ticket
            # This prevents duplicate creation if thread appears again in same batch
            thread_state[tid] = ts
            print(f"   ✅ DEBUG: Marked {tid} as processed with timestamp {ts}")

        # Determine status based on last sender
        if is_admin_email(from_email):
            status = "Awaiting customer reply"
        else:
            status = "Awaiting admin reply"

        # Update Gmail labels
        labels_to_add = [admin_label] if status == "Awaiting admin reply" else [cust_label]
        labels_to_remove = [cust_label] if status == "Awaiting admin reply" else [admin_label]
        
        update_thread_labels(gmail, tid, labels_to_add, labels_to_remove)

        # Check if new sender (only for new tickets)
        new_sender = False
        if is_new_ticket:
            new_sender = is_new_sender_cached(from_email)
            if new_sender:
                print(f"   🆕 NEW SENDER: {from_email}")
                known_senders.add(from_email.lower())  # Add to cache
            else:
                print(f"   👤 Known sender: {from_email}")

        # Create row data
        row_data = create_ticket_row(ticket_id, tid, from_email, subject, status, new_sender)

        if not is_new_ticket:
            # Update existing ticket
            update_existing_ticket(main_worksheet, row_num, row_data)
            print(f"   ✅ Updated ticket {ticket_id}")
        else:
            # Create new ticket
            add_new_ticket(main_worksheet, row_data)
            print(f"   ✅ Created ticket {ticket_id}")
            
            # IMPORTANT: Refresh the cached thread map immediately after creating a new ticket
            # This prevents duplicate ticket creation if the same thread is processed again
            cached_thread_map = get_all_tickets(main_worksheet)
            print(f"   🔄 Refreshed cache to include new ticket")

        # Mark as processed (update timestamp)
        thread_state[tid] = ts

    # Save thread state to FILE (always - fast, no API quota)
    if threads:
        save_thread_state_to_file(thread_state)
        print(f"💾 Saved thread state to file")
    
    # Backup to sheet every N syncs (reduce API calls)
    if sync_counter % SHEET_BACKUP_INTERVAL == 0:
        save_thread_state_to_sheet(sheet, thread_state)
        save_last_sync_timestamp(sheet, int(time.time()))
        print(f"📊 Backed up thread state AND sync timestamp to sheet (sync #{sync_counter})")
    
    # Check for stale tickets every 20 syncs (every ~100 seconds)
    if sync_counter % 20 == 0:
        check_and_close_stale_tickets(gmail, sheet, main_worksheet)
    
    print("\n" + "="*50)
    print("✅ Sync complete!")
    print("="*50 + "\n")