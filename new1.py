import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os, pickle, time, re, base64
from datetime import datetime
from email.message import EmailMessage

# ================= CONFIG =================
SPREADSHEET_ID = "1E8iz5VOA8hnIpEFM-ZiBF9wh7tpKKcDHBx1hfBwc2VE"
CHECK_INTERVAL = 5
LAST_SYNC_FILE = "last_sync.txt"
THREAD_STATE_FILE = "thread_state.txt"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
]

# ================= AUTH =================
def get_credentials():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open("token.pickle", "wb") as f:
            pickle.dump(creds, f)
    return creds

# ================= STATE FILE =================
def load_thread_state():
    state = {}
    if os.path.exists(THREAD_STATE_FILE):
        with open(THREAD_STATE_FILE, "r") as f:
            for line in f:
                if "|" in line:
                    tid, ts = line.strip().split("|")
                    state[tid] = int(ts)
    return state

def save_thread_state(state):
    with open(THREAD_STATE_FILE, "w") as f:
        for tid, ts in state.items():
            f.write(f"{tid}|{ts}\n")

# ================= LAST SYNC =================
def load_last_sync():
    if os.path.exists(LAST_SYNC_FILE):
        with open(LAST_SYNC_FILE, "r") as f:
            return int(f.read().strip())
    return None

def save_last_sync(ts):
    with open(LAST_SYNC_FILE, "w") as f:
        f.write(str(ts))

# ================= HELPERS =================
def extract_email(v):
    m = re.search(r"<(.+?)>", v)
    return m.group(1).lower() if m else v.lower()

def get_my_email(gmail):
    return gmail.users().getProfile(userId="me").execute()["emailAddress"].lower()

def get_or_create_label(gmail, name):
    for l in gmail.users().labels().list(userId="me").execute()["labels"]:
        if l["name"] == name:
            return l["id"]
    return gmail.users().labels().create(
        userId="me",
        body={"name": name, "labelListVisibility": "labelShow", "messageListVisibility": "show"},
    ).execute()["id"]

def fetch_all_threads(gmail, query):
    """Fetch threads instead of messages to avoid duplicates"""
    threads, token = [], None
    while True:
        res = gmail.users().threads().list(
            userId="me", q=query, maxResults=100, pageToken=token
        ).execute()
        threads.extend(res.get("threads", []))
        token = res.get("nextPageToken")
        if not token:
            break
    return threads

def get_last_message(thread):
    """Get the LAST message in the thread from ANYONE"""
    if not thread.get("messages"):
        return None, None
    
    last_msg = max(thread["messages"], key=lambda x: int(x["internalDate"]))
    headers = {x["name"]: x["value"] for x in last_msg["payload"]["headers"]}
    
    return last_msg, headers

def get_first_customer_message(thread, my_email):
    """Get the FIRST message from a customer for attachments"""
    if not thread.get("messages"):
        return None
    
    for msg in sorted(thread["messages"], key=lambda x: int(x["internalDate"])):
        headers = {x["name"]: x["value"] for x in msg["payload"]["headers"]}
        sender = extract_email(headers.get("From", ""))
        if sender != my_email:
            return msg
    
    return None

def get_next_ticket_id(sheet):
    cfg = sheet.worksheet("Ticket_Config")
    last = int(cfg.acell("B1").value or 0)
    nxt = last + 1
    cfg.update(range_name="B1", values=[[nxt]])
    return f"TCK-{nxt:06d}"

def get_attachments_from_message(gmail, message_id):
    """Extract all attachments from a Gmail message"""
    attachments = []
    try:
        msg = gmail.users().messages().get(userId="me", id=message_id).execute()
        
        if "parts" in msg["payload"]:
            for part in msg["payload"]["parts"]:
                if part.get("filename") and part.get("body", {}).get("attachmentId"):
                    attachment = gmail.users().messages().attachments().get(
                        userId="me",
                        messageId=message_id,
                        id=part["body"]["attachmentId"]
                    ).execute()
                    
                    attachments.append({
                        "filename": part["filename"],
                        "mimeType": part["mimeType"],
                        "data": attachment["data"]
                    })
    except Exception as e:
        print(f"⚠️ Error fetching attachments: {e}")
    
    return attachments

def send_auto_reply(gmail, to_email, ticket_id, subject, thread_id, msg_id, original_message_id):
    """Send auto-reply with all attachments from the original message"""
    
    print(f"📤 Sending auto-reply to: {to_email}, Ticket: {ticket_id}")
    
    # Get attachments
    attachments = []
    if original_message_id:
        attachments = get_attachments_from_message(gmail, original_message_id)
        print(f"   Found {len(attachments)} attachment(s)")
    
    msg = EmailMessage()
    msg["To"] = to_email
    msg["From"] = "me"
    msg["Subject"] = f"Re: {subject}" if not subject.startswith("Re:") else subject
    
    # IMPORTANT: Set threading headers properly
    if msg_id:
        msg["In-Reply-To"] = msg_id
        msg["References"] = msg_id
    
    msg.set_content(
        f"""Hi,

Your support ticket has been created.
Ticket ID: {ticket_id}

Please reply to this email to continue the conversation.

Regards,
Support Team"""
    )
    
    # Add attachments
    if attachments:
        for att in attachments:
            try:
                file_data = base64.urlsafe_b64decode(att["data"])
                mime_parts = att["mimeType"].split("/")
                maintype = mime_parts[0] if len(mime_parts) > 0 else "application"
                subtype = mime_parts[1] if len(mime_parts) > 1 else "octet-stream"
                
                msg.add_attachment(
                    file_data,
                    maintype=maintype,
                    subtype=subtype,
                    filename=att["filename"]
                )
                print(f"   ✅ Attached: {att['filename']}")
            except Exception as e:
                print(f"   ⚠️ Failed to attach {att.get('filename')}: {e}")
    
    try:
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        result = gmail.users().messages().send(
            userId="me", 
            body={"raw": raw, "threadId": thread_id}
        ).execute()
        print(f"   ✅ Email sent! ID: {result.get('id')}")
        return True
    except Exception as e:
        print(f"   ❌ Failed to send: {e}")
        import traceback
        traceback.print_exc()
        return False

# ================= MAIN =================
def sync_mail_to_sheet():
    creds = get_credentials()
    gmail = build("gmail", "v1", credentials=creds)
    my_email = get_my_email(gmail)

    gc = gspread.authorize(creds)
    sheet = gc.open_by_key(SPREADSHEET_ID)
    main = sheet.worksheet("Email log")

    admin_label = get_or_create_label(gmail, "Awaiting_Admin_Reply")
    cust_label = get_or_create_label(gmail, "Awaiting_Customer_Reply")

    # Get existing tickets
    rows = main.get_all_values()
    thread_map = {r[1]: i + 2 for i, r in enumerate(rows[1:]) if len(r) > 1}
    print(f"📋 Found {len(thread_map)} existing tickets")

    thread_state = load_thread_state()
    last_sync = load_last_sync()
    query = f"after:{last_sync}" if last_sync else "newer_than:7d"

    # Fetch THREADS not messages - this prevents duplicates
    threads = fetch_all_threads(gmail, query)
    print(f"🔍 Found {len(threads)} threads to process")

    for thread_info in threads:
        tid = thread_info["id"]
        
        # Get full thread details
        thread = gmail.users().threads().get(userId="me", id=tid).execute()
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
        msg_id = headers.get("Message-ID")

        print(f"\n📨 Processing thread {tid}")
        print(f"   From: {from_email}")
        print(f"   Subject: {subject}")

        # Determine if new or existing ticket
        is_new_ticket = tid not in thread_map
        
        # Skip NEW threads initiated by admin
        if is_new_ticket and from_email == my_email:
            print(f"   ⏭️ Skipping - admin-initiated thread")
            thread_state[tid] = ts
            continue

        if not is_new_ticket:
            # Update existing ticket
            row_num = thread_map[tid]
            prev = main.row_values(row_num)
            ticket_id = prev[0]
        else:
            # Create new ticket
            ticket_id = get_next_ticket_id(sheet)
            print(f"   🎫 New ticket: {ticket_id}")

        # Determine status
        if from_email == my_email:
            status = "Awaiting customer reply"
        else:
            status = "Awaiting admin reply"

        # Update labels
        gmail.users().threads().modify(
            userId="me",
            id=tid,
            body={
                "addLabelIds": [admin_label if status == "Awaiting admin reply" else cust_label],
                "removeLabelIds": [cust_label if status == "Awaiting admin reply" else admin_label],
            },
        ).execute()

        link = f'=HYPERLINK("https://mail.google.com/mail/u/0/#inbox/{tid}","Open Mail")'
        
        row = [
            ticket_id,
            tid,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            from_email,
            subject,
            status,
            "No",
            link,
        ]

        if not is_new_ticket:
            # Update existing row
            main.update(f"A{row_num}:H{row_num}", [row], value_input_option="USER_ENTERED")
            print(f"   ✅ Updated ticket {ticket_id}")
        else:
            # Create new row and send auto-reply
            main.append_row(row, value_input_option="USER_ENTERED")
            print(f"   ✅ Created ticket {ticket_id}")
            
            # Send auto-reply with attachments
            first_customer_msg = get_first_customer_message(thread, my_email)
            if first_customer_msg:
                send_auto_reply(
                    gmail, 
                    from_email, 
                    ticket_id, 
                    subject, 
                    tid, 
                    msg_id, 
                    first_customer_msg.get("id")
                )
            else:
                print(f"   ⚠️ Could not find customer message for attachments")

        # Mark as processed
        thread_state[tid] = ts
        save_thread_state(thread_state)

    save_last_sync(int(time.time()))
    print(f"\n✅ Sync complete\n")

# ================= LOOP =================
def main():
    print("🚀 Ticket system running")
    while True:
        try:
            sync_mail_to_sheet()
            time.sleep(CHECK_INTERVAL)
        except KeyboardInterrupt:
            print("🛑 Stopped")
            break
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
            time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    main()
