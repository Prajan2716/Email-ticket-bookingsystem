"""
Gmail Handler - Manages all Gmail API operations
"""
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os
import pickle
import re

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
]


def get_gmail_credentials():
    """Authenticate and return Gmail credentials"""
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


def get_gmail_service():
    """Build and return Gmail service"""
    creds = get_gmail_credentials()
    return build("gmail", "v1", credentials=creds)


def get_my_email(gmail):
    """Get the authenticated user's email address"""
    return gmail.users().getProfile(userId="me").execute()["emailAddress"].lower()


def extract_email(value):
    """Extract email address from string like 'Name <email@example.com>'"""
    m = re.search(r"<(.+?)>", value)
    return m.group(1).lower() if m else value.lower()


def is_noreply_email(email):
    """
    Check if email is a no-reply address
    Returns: True if it's a no-reply email
    """
    email_lower = email.lower()
    
    # Common no-reply patterns
    noreply_patterns = [
        "noreply@",
        "no-reply@",
        "no_reply@",
        "donotreply@",
        "do-not-reply@",
        "do_not_reply@",
        "notifications@",
        "notification@",
        "automated@",
        "automation@",
        "mailer@",
        "daemon@",
        "bounce@",
        "bounces@"
    ]
    
    # Check if email starts with any no-reply pattern
    for pattern in noreply_patterns:
        if email_lower.startswith(pattern):
            return True
    
    return False


def get_or_create_label(gmail, name):
    """Get existing label or create new one"""
    labels = gmail.users().labels().list(userId="me").execute()["labels"]
    for label in labels:
        if label["name"] == name:
            return label["id"]
    
    # Create label if it doesn't exist
    label_object = {
        "name": name,
        "labelListVisibility": "labelShow",
        "messageListVisibility": "show"
    }
    created = gmail.users().labels().create(userId="me", body=label_object).execute()
    return created["id"]


def fetch_all_threads(gmail, query):
    """Fetch all threads matching the query"""
    threads = []
    token = None
    
    while True:
        response = gmail.users().threads().list(
            userId="me",
            q=query,
            maxResults=100,
            pageToken=token
        ).execute()
        
        threads.extend(response.get("threads", []))
        token = response.get("nextPageToken")
        
        if not token:
            break
    
    return threads


def get_thread_details(gmail, thread_id):
    """Get full thread details including all messages"""
    return gmail.users().threads().get(userId="me", id=thread_id).execute()


def get_last_message(thread):
    """Get the most recent message in the thread"""
    if not thread.get("messages"):
        return None, None
    
    last_msg = max(thread["messages"], key=lambda x: int(x["internalDate"]))
    headers = {x["name"]: x["value"] for x in last_msg["payload"]["headers"]}
    
    return last_msg, headers


def update_thread_labels(gmail, thread_id, add_labels=None, remove_labels=None):
    """Update labels on a thread"""
    body = {}
    if add_labels:
        body["addLabelIds"] = add_labels
    if remove_labels:
        body["removeLabelIds"] = remove_labels
    
    gmail.users().threads().modify(
        userId="me",
        id=thread_id,
        body=body
    ).execute()
