import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import pickle
import time
from datetime import datetime
import re

# ============================================
# CONFIG
# ============================================
SPREADSHEET_ID = "1E8iz5VOA8hnIpEFM-ZiBF9wh7tpKKcDHBx1hfBwc2VE"
CHECK_INTERVAL = 5
LAST_SYNC_FILE = "last_sync.txt"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
]

# ============================================
# AUTH
# ============================================
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

# ============================================
# LAST SYNC
# ============================================
def load_last_sync():
    if os.path.exists(LAST_SYNC_FILE):
        with open(LAST_SYNC_FILE, "r") as f:
            return int(f.read().strip())
    return None

def save_last_sync(ts):
    with open(LAST_SYNC_FILE, "w") as f:
        f.write(str(ts))

# ============================================
# HELPERS
# ============================================
def extract_email(from_field):
    match = re.search(r"<(.+?)>", from_field)
    return match.group(1) if match else from_field

def get_or_create_label(service, name):
    labels = service.users().labels().list(userId="me").execute()
    for label in labels.get("labels", []):
        if label["name"] == name:
            return label["id"]

    label = service.users().labels().create(
        userId="me",
        body={
            "name": name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        },
    ).execute()
    return label["id"]

def fetch_all_messages(service, query):
    messages = []
    page_token = None

    while True:
        response = service.users().messages().list(
            userId="me",
            q=query,
            maxResults=500,
            pageToken=page_token,
        ).execute()

        messages.extend(response.get("messages", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    return messages

# ============================================
# MAIN SYNC
# ============================================
def sync_mail_to_sheet():
    creds = get_credentials()
    gmail = build("gmail", "v1", credentials=creds)
    gc = gspread.authorize(creds)
    sheet = gc.open_by_key(SPREADSHEET_ID)

    main_sheet = sheet.worksheet("Email log")
    admin_sheet = sheet.worksheet("Admin emails")

    try:
        known_sheet = sheet.worksheet("Known Senders")
    except:
        known_sheet = sheet.add_worksheet("Known Senders", 1000, 2)
        known_sheet.update([["Email", "First Seen"]], "A1:B1")

    admin_emails = {
        r[0].lower().strip()
        for r in admin_sheet.get_all_values()[1:]
        if r and r[0]
    }

    known_senders = {
        r[0].lower().strip()
        for r in known_sheet.get_all_values()[1:]
        if r and r[0]
    }

    admin_label = get_or_create_label(gmail, "Awaiting_Admin_Reply")
    customer_label = get_or_create_label(gmail, "Awaiting_Customer_Reply")

    rows = main_sheet.get_all_values()
    thread_map = {r[0]: i + 2 for i, r in enumerate(rows[1:]) if r}

    last_sync = load_last_sync()
    if last_sync:
        query = f"after:{last_sync}"
        print(f"🔍 Fetching emails after {last_sync}")
    else:
        query = "newer_than:7d"
        print("🔍 First run – last 7 days")

    messages = fetch_all_messages(gmail, query)
    print(f"📬 Total messages fetched: {len(messages)}")

    processed_threads = set()

    for msg in messages:
        msg_data = gmail.users().messages().get(
            userId="me", id=msg["id"], format="full"
        ).execute()

        thread_id = msg_data["threadId"]
        if thread_id in processed_threads:
            continue
        processed_threads.add(thread_id)

        thread = gmail.users().threads().get(
            userId="me", id=thread_id
        ).execute()

        latest = max(
            thread["messages"],
            key=lambda x: int(x["internalDate"]),
        )

        headers = {h["name"]: h["value"] for h in latest["payload"]["headers"]}
        from_email = extract_email(headers.get("From", "")).lower()
        subject = headers.get("Subject", "No Subject")

        from email.utils import parsedate_to_datetime
        try:
            date = parsedate_to_datetime(headers.get("Date"))
        except:
            date = datetime.now()

        is_new = from_email not in known_senders and from_email not in admin_emails
        if is_new:
            known_sheet.append_row(
                [from_email, datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                value_input_option="USER_ENTERED",
            )
            known_senders.add(from_email)

        status = (
            "Awaiting customer reply"
            if from_email in admin_emails
            else "Awaiting admin reply"
        )

        gmail_link = (
            f'=HYPERLINK("https://mail.google.com/mail/u/0/#inbox/{thread_id}","Open Mail")'
        )

        gmail.users().threads().modify(
            userId="me",
            id=thread_id,
            body={
                "addLabelIds": [admin_label if status == "Awaiting admin reply" else customer_label],
                "removeLabelIds": [customer_label if status == "Awaiting admin reply" else admin_label],
            },
        ).execute()

        row = [
            thread_id,
            date.strftime("%Y-%m-%d %H:%M:%S"),
            from_email,
            subject,
            status,
            "Yes" if is_new else "No",
            gmail_link,
        ]

        if thread_id in thread_map:
            r = thread_map[thread_id]
            main_sheet.update(
                [row],
                f"A{r}:G{r}",
                value_input_option="USER_ENTERED",
            )
        else:
            main_sheet.append_row(row, value_input_option="USER_ENTERED")

    save_last_sync(int(time.time()))
    print("✅ Sync completed safely")

# ============================================
# LOOP
# ============================================
def main():
    print("🚀 Gmail → Sheets Sync Running")
    while True:
        try:
            sync_mail_to_sheet()
            time.sleep(CHECK_INTERVAL)
        except KeyboardInterrupt:
            print("🛑 Stopped")
            break
        except Exception as e:
            print(f"❌ Error: {e}")
            time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    main()
