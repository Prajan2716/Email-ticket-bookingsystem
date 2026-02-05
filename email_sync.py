import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import pickle
import time
from datetime import datetime
import re

# ============================================
# CONFIGURATION - CHANGE THESE VALUES!
# ============================================
SPREADSHEET_NAME = "Email Tracker"  # Your Google Sheet name
SPREADSHEET_ID = "1E8iz5VOA8hnIpEFM-ZiBF9wh7tpKKcDHBx1hfBwc2VE"  # ← GET THIS FROM YOUR SHEET URL
CHECK_INTERVAL = 5  # seconds between checks

# Google API Scopes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify'
]

def get_credentials():
    """Get user credentials with OAuth"""
    
    # Print current directory for debugging
    current_dir = os.getcwd()
    print(f"📁 Current directory: {current_dir}")
    
    # Check if credentials.json exists
    if not os.path.exists('credentials.json'):
        print("\n" + "="*60)
        print("❌ ERROR: credentials.json NOT FOUND!")
        print("="*60)
        print(f"📂 Files in current directory:")
        for file in os.listdir('.'):
            print(f"   - {file}")

        input("\nPress Enter to exit...")
        exit()
    
    print("✅ credentials.json found!")
    
    creds = None
    
    # Check if we have saved credentials
    if os.path.exists('token.pickle'):
        print("📌 Loading saved credentials...")
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If no valid credentials, authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Refreshing expired credentials...")
            creds.refresh(Request())
        else:
            print("🔐 Opening browser for Google login...")
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for next time
        print("💾 Saving credentials...")
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def extract_email(from_field):
    """Extract email from 'Name <email@domain.com>' format"""
    match = re.search(r'<(.+?)>', from_field)
    return match.group(1) if match else from_field

def get_or_create_label(service, label_name):
    """Get existing Gmail label or create new one"""
    try:
        labels = service.users().labels().list(userId='me').execute()
        label_dict = {label['name']: label['id'] for label in labels.get('labels', [])}
        
        if label_name in label_dict:
            return label_dict[label_name]
        else:
            # Create new label
            new_label = service.users().labels().create(
                userId='me',
                body={
                    'name': label_name,
                    'labelListVisibility': 'labelShow',
                    'messageListVisibility': 'show'
                }
            ).execute()
            print(f"✅ Created Gmail label: {label_name}")
            return new_label['id']
    except Exception as e:
        print(f"⚠️ Warning - Could not manage label '{label_name}': {e}")
        return None

def sync_mail_to_sheet():
    """Main function to sync Gmail to Google Sheets"""
    
    try:
        # Get credentials
        creds = get_credentials()
        
        # Connect to Google Sheets
        print("📊 Connecting to Google Sheets...")
        gc = gspread.authorize(creds)
        
        # Open spreadsheet by ID (more reliable than name)
        try:
            spreadsheet = gc.open_by_key(SPREADSHEET_ID)
            print(f"✅ Spreadsheet opened successfully!")
        except Exception as e:
            print("\n" + "="*60)
            print(f"❌ ERROR: Cannot open spreadsheet!")
            print("="*60)
            print(f"Error details: {e}")
            print("\n💡 SOLUTION:")
            print("1. Open your Google Sheet in browser")
            print("2. Copy the ID from the URL:")
            print("   https://docs.google.com/spreadsheets/d/YOUR_ID_HERE/edit")
            print("3. Paste it in the SPREADSHEET_ID variable in the code")
            print("="*60)
            return
        
        # Get worksheets
        try:
            main_sheet = spreadsheet.worksheet("Email log")
            admin_sheet = spreadsheet.worksheet("Admin emails")  # lowercase 'e'
        except Exception as e:
            print(f"❌ ERROR: Worksheet not found - {e}")
            print("Make sure your spreadsheet has these sheets:")
            print("  - Email log")
            print("  - Admin emails")
            return
        
        # Load admin emails from sheet
        print("📧 Loading admin emails...")
        admin_emails = set()
        admin_data = admin_sheet.get_all_values()[1:]  # Skip header row
        for row in admin_data:
            if row and row[0]:  # If row exists and has an email
                admin_emails.add(row[0].lower().strip())
        
        print(f"✅ Loaded {len(admin_emails)} admin email(s)")
        
        # Connect to Gmail API
        print("📬 Connecting to Gmail...")
        gmail_service = build('gmail', 'v1', credentials=creds)
        
        # Get or create Gmail labels
        admin_label_id = get_or_create_label(gmail_service, "Awaiting_Admin_Reply")
        customer_label_id = get_or_create_label(gmail_service, "Awaiting_Customer_Reply")
        
        # Map existing threads from sheet
        sheet_data = main_sheet.get_all_values()
        thread_map = {}
        for i, row in enumerate(sheet_data[1:], start=2):  # Start from row 2 (skip header)
            if row and row[0]:  # If row exists and has a thread ID
                thread_map[row[0]] = i
        
        print(f"📋 Found {len(thread_map)} existing threads in sheet")
        
        # Search for recent Gmail threads
        print("🔍 Searching Gmail (last 7 days)...")
        results = gmail_service.users().messages().list(
            userId='me',
            q='newer_than:7d',
            maxResults=30
        ).execute()
        
        messages = results.get('messages', [])
        
        if not messages:
            print("📭 No messages found in the last 7 days")
            return
        
        print(f"📬 Processing {len(messages)} message(s)...")
        
        # Track processed threads to avoid duplicates
        processed_threads = set()
        updates_count = 0
        new_count = 0
        
        # Process each message
        for msg in messages:
            # Get full message details
            full_msg = gmail_service.users().messages().get(
                userId='me',
                id=msg['id'],
                format='full'
            ).execute()
            
            thread_id = full_msg['threadId']
            
            # Skip if already processed in this run
            if thread_id in processed_threads:
                continue
            processed_threads.add(thread_id)
            
            # Get full thread to find latest message
            thread = gmail_service.users().threads().get(
                userId='me',
                id=thread_id
            ).execute()
            
            # Find the latest message in thread
            thread_messages = thread['messages']
            latest_message = max(thread_messages, key=lambda x: int(x['internalDate']))
            
            # Extract email headers
            headers = {h['name']: h['value'] for h in latest_message['payload']['headers']}
            
            from_email = extract_email(headers.get('From', '')).lower()
            subject = headers.get('Subject', 'No Subject')
            date_str = headers.get('Date', '')
            
            # Parse date
            try:
                from email.utils import parsedate_to_datetime
                date = parsedate_to_datetime(date_str)
            except:
                date = datetime.now()
            
            # Determine status based on sender
            if from_email in admin_emails:
                status = "Awaiting customer reply"
            else:
                status = "Awaiting admin reply"
            
            # Create Gmail hyperlink formula (with = sign)
            gmail_link = f'=HYPERLINK("https://mail.google.com/mail/u/0/#inbox/{thread_id}", "Open Mail")'
            
            # Apply Gmail labels
            if admin_label_id and customer_label_id:
                try:
                    if status == "Awaiting admin reply":
                        gmail_service.users().threads().modify(
                            userId='me',
                            id=thread_id,
                            body={
                                'addLabelIds': [admin_label_id],
                                'removeLabelIds': [customer_label_id]
                            }
                        ).execute()
                    else:
                        gmail_service.users().threads().modify(
                            userId='me',
                            id=thread_id,
                            body={
                                'addLabelIds': [customer_label_id],
                                'removeLabelIds': [admin_label_id]
                            }
                        ).execute()
                except Exception as e:
                    print(f"⚠️ Could not update labels for thread: {e}")
            
            # Prepare row data
            row_data = [
                thread_id,
                date.strftime("%Y-%m-%d %H:%M:%S"),
                from_email,
                subject,
                status,
                gmail_link
            ]
            
            # Update or insert in sheet with USER_ENTERED option (THIS FIXES THE HYPERLINK!)
            if thread_id in thread_map:
                # Update existing row
                row_num = thread_map[thread_id]
                main_sheet.update(
                    f'A{row_num}:F{row_num}', 
                    [row_data],
                    value_input_option='USER_ENTERED'  # ← THIS MAKES HYPERLINKS WORK!
                )
                updates_count += 1
            else:
                # Append new row
                main_sheet.append_row(
                    row_data,
                    value_input_option='USER_ENTERED'  # ← THIS MAKES HYPERLINKS WORK!
                )
                thread_map[thread_id] = len(thread_map) + 2
                new_count += 1
                print(f"  ➕ New: {subject[:50]}...")
        
        print(f"✅ Sync complete! New: {new_count}, Updated: {updates_count}")
        
    except Exception as e:
        print(f"\n❌ ERROR during sync: {e}")
        import traceback
        traceback.print_exc()

def main():
    """Main loop - runs sync every 30 seconds"""
    
    print("\n" + "="*60)
    print("🚀 EMAIL TO SHEET SYNC - STARTED")
    print("="*60)
    print(f"📊 Spreadsheet ID: {SPREADSHEET_ID}")
    print(f"⏱️  Check interval: {CHECK_INTERVAL} seconds")
    print(f"⌨️  Press Ctrl+C to stop")
    print("="*60)
    print()
    
    # First-time setup message
    if not os.path.exists('token.pickle'):
        print("🔐 FIRST TIME SETUP:")
        print("A browser window will open for Google login...")
        print("Please authorize the app to access Gmail and Sheets")
        print()
    
    run_count = 0
    
    while True:
        try:
            run_count += 1
            current_time = datetime.now().strftime('%H:%M:%S')
            
            print(f"🔄 Run #{run_count} at {current_time}")
            print("-" * 60)
            
            sync_mail_to_sheet()
            
            print("-" * 60)
            print(f"⏳ Waiting {CHECK_INTERVAL} seconds until next sync...")
            print()
            
            time.sleep(CHECK_INTERVAL)
            
        except KeyboardInterrupt:
            print("\n" + "="*60)
            print("🛑 SYNC STOPPED BY USER")
            print("="*60)
            print(f"Total runs completed: {run_count}")
            print("Goodbye! 👋")
            break
            
        except Exception as e:
            print(f"\n❌ Unexpected error: {e}")
            print(f"⏳ Waiting {CHECK_INTERVAL} seconds before retry...")
            print()
            time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    main()
