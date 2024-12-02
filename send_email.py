import os
import sys
import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import json
import os.path
from google.auth.transport.requests import Request
import pathlib
from email.message import EmailMessage
from googleapiclient.errors import HttpError
from email.utils import formataddr


# Set up the path for the credentials and refresh token files
BASE_SCOPE_URL = "https://www.googleapis.com/auth/"
PROJECT_DIR = pathlib.Path(os.path.dirname(os.path.abspath(__file__)))

CLIENT_SECRET_FILE = PROJECT_DIR / "credentials.json"
REFRESH_TOKEN_FILE = PROJECT_DIR / "refresh_token.json"

# Define the required scope for Gmail API
SCOPES = [f"{BASE_SCOPE_URL}gmail.send", f"{BASE_SCOPE_URL}gmail.readonly"]

def authenticate_google_account():
    """Authenticate and return a valid Google API service."""
    creds = None

    # Check if refresh token file exists and load credentials
    if os.path.exists(REFRESH_TOKEN_FILE):
        with open(REFRESH_TOKEN_FILE, "r") as token:
            creds_data = json.load(token)
        creds = Credentials.from_authorized_user_info(creds_data, SCOPES)

    # If there are no (valid) credentials available, request new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())  # Refresh expired token
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)  # Request new credentials if needed

        # Save the credentials for the next run
        with open(REFRESH_TOKEN_FILE, "w") as token:
            token.write(creds.to_json())

    # Build the Gmail API service
    service = build("gmail", "v1", credentials=creds)
    return service


def send_email_with_attachment(recipient, subject, message_text, attachment_paths=None):
    """Send an email with attachments using Gmail API."""
    try:

        service = authenticate_google_account()

        # Fetch the email address
        user_profile = service.users().getProfile(userId="me").execute()
        email_address = user_profile["emailAddress"]
        formatted_from = formataddr(("Elio Gonzalez", email_address))

        # Create the email message
        message = EmailMessage()
        message["To"] = formataddr(recipient) 
        message["From"] = formatted_from
        message["Subject"] = subject
        message.set_content(message_text, subtype="html")  # Specify HTML content if needed

        # Attach files, if any
        if attachment_paths:
            for file_path in attachment_paths:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as f:
                        file_data = f.read()
                        file_name = os.path.basename(file_path)
                        message.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
                else:
                    print(f"File not found: {file_path}")

        # Encode the message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
        create_message = {"raw": encoded_message}

        # Send the email
        sent_message = service.users().messages().send(userId="me", body=create_message).execute()
        print(f"Email sent! Message ID: {sent_message['id']}")
        return sent_message

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None



def generate_email_template(email_type, details):
    if email_type == 'missing_info':
        return f"""<html>
    <body>
    <p>Hi Pablo,</p>
    
    <p>Below are the details about the client:</p>
    
    <p>{details}</p>  <!-- Make sure details are enclosed in proper HTML tags -->
    
    <p>Let me know if any additional steps are needed or if clarification is required to move this case forward.</p>
    
    <p>Best regards,</p>
    <p><b>Elio Gonzalez</b></p>
    </body>
    </html>"""

    elif email_type == 'completed':
        return f"""<html>
    <body>
    <p>Hi Pablo,</p>
    
    <p>Here is a detailed explanation, and attached is the completed demand letter I wrote for this client:</p>
    
    <p>{details}</p>

    <p>Any feedback on how to improve would be greatly appreciated.</p>
    
    <p>Best regards,</p>
    <p><b>Elio Gonzalez</b></p>
    </body>
    </html>"""

    else:
        return "Invalid email type"
    

def generate_details_html(client_details):
    """
    Generates an HTML-formatted string with proper escaping for client details.

    :param client_details: String containing details in a certain format
    :return: HTML formatted string
    """
    # Split the input details into separate client entries
    entries = client_details.split("\n\n")

    html_content = ""

    html_content = f"<b>{entries[0].replace("**", "")}</b><br><ul>"
    
    for entry in entries[1:]:
        lines = entry.split("\n")
        for line in lines:
            if line.replace("- ", "")[0:4] != "    ":
                html_content += f"<li>{line.replace("- ", "").strip()}</li>"
            else: # Sub category of the cat
                html_content += f"<ul><li>{line.replace("- ", "")}</li></ul>"
        
    return html_content + "</ul>"


def send_email(to, email_type, client_name, case_details, attachment=None):

    email_body = generate_email_template(email_type, generate_details_html(case_details))

    if email_type == "missing_info":
        subject = f"Missing Information for: {client_name}"
    elif email_type == "completed":
        subject = f"Completed Demand Letter for Review: {client_name}"
    else:
        print("Make sure to add the field email_type to send the email!")
        sys.exit(1)

    # Send the email
    send_email_with_attachment(
        recipient=to,
        subject=subject,
        message_text=email_body,
        attachment_paths=attachment
    )


