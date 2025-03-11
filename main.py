import os
import base64
import json
import time
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
import pickle
import tempfile
import anthropic
import re

# Google API scopes
SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Configuration - replace with your values
EMAIL_LABEL = "ai-receipts"  # The Gmail label/group to search for
SENDER_EMAIL = "paul@paulhauner.com"  # Your email address
SPREADSHEET_ID = "1oM5APBsN7JqADj71VHLp14BVwaOd5Azx9g3m_Cm5B8M"  # Google Sheet ID
SHEET_NAME = "Transactions"  # Sheet name in your Google Spreadsheet
TRACKING_SHEET_NAME = "ProcessedEmails"  # Sheet to track processed emails
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")  # Set as environment variable


def get_credentials():
    """Get and refresh Google API credentials"""
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local(port=0)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return creds


def get_email_content(service, message_id):
    """Get email content and attachments"""
    message = service.users().messages().get(userId="me", id=message_id).execute()

    # Extract email body
    email_body = ""
    if "payload" in message:
        if "body" in message["payload"] and message["payload"]["body"].get("data"):
            body_data = message["payload"]["body"]["data"]
            email_body = base64.urlsafe_b64decode(body_data).decode("utf-8")
        elif "parts" in message["payload"]:
            for part in message["payload"]["parts"]:
                if part.get("mimeType") == "text/plain" and part.get("body", {}).get(
                    "data"
                ):
                    body_data = part["body"]["data"]
                    email_body = base64.urlsafe_b64decode(body_data).decode("utf-8")
                    break

    # Extract subject
    headers = message["payload"]["headers"]
    subject = next(
        (h["value"] for h in headers if h["name"] == "Subject"), "No Subject"
    )

    # Extract attachments
    attachments = []

    if "payload" in message and "parts" in message["payload"]:
        parts = message["payload"].get("parts", [])

        for part in parts:
            if (
                part.get("filename")
                and part.get("body")
                and part.get("body").get("attachmentId")
            ):
                attachment = (
                    service.users()
                    .messages()
                    .attachments()
                    .get(
                        userId="me",
                        messageId=message_id,
                        id=part["body"]["attachmentId"],
                    )
                    .execute()
                )

                file_data = base64.urlsafe_b64decode(attachment["data"])
                temp_file = tempfile.NamedTemporaryFile(
                    delete=False, suffix=part["filename"]
                )
                temp_file.write(file_data)
                temp_file.close()

                attachments.append(
                    {
                        "filename": part["filename"],
                        "filepath": temp_file.name,
                        "mimetype": part.get("mimeType", ""),
                    }
                )

    return {"subject": subject, "body": email_body, "attachments": attachments}


def process_with_anthropic(email_content, attachments):
    """Send email content and files to Anthropic API and request JSON response"""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Prepare the files for the API
    media = [
        {
            "type": "text",
            "text": f"Subject: {email_content['subject']}\n\nEmail Body:\n{email_content['body']}\n\nPlease analyze this email and any attached documents to extract financial information.",
        }
    ]

    # Add attachments if available
    for attachment in attachments:
        with open(attachment["filepath"], "rb") as f:
            file_content = f.read()
            mime_type = attachment["mimetype"]
            if not mime_type or mime_type == "":
                # Try to guess MIME type from extension
                if attachment["filename"].lower().endswith(".pdf"):
                    mime_type = "application/pdf"
                elif attachment["filename"].lower().endswith((".jpg", ".jpeg")):
                    mime_type = "image/jpeg"
                elif attachment["filename"].lower().endswith(".png"):
                    mime_type = "image/png"
                else:
                    mime_type = "application/octet-stream"

            media.append(
                {
                    "type": "file",
                    "source": {
                        "type": "base64",
                        "media_type": mime_type,
                        "data": base64.b64encode(file_content).decode("utf-8"),
                    },
                }
            )

    try:
        system_prompt = """
You analyze financial documents and extract structured data. Extract the following information:
1. Date: The invoice/statement date (not today's date)
2. Description: Brief description of the transaction
3. Amount: Numeric value (positive for income, negative for expenses)
4. Category: The type of expense or income (e.g., Rent Income, Utilities, Maintenance, Groceries, Salary)
5. Property: If related to a property, include the property address, otherwise leave blank

Return a valid JSON object with these fields. The JSON should be valid and parseable.
Use negative amounts for expenses and positive amounts for income.
Use the email body context to help categorize the transaction if needed.
"""

        response = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=4000,
            system=system_prompt,
            messages=[{"role": "user", "content": media}],
        )

        # The response should be just JSON, but let's try to extract it if there's any wrapper text
        content = response.content[0].text

        # Try to find JSON within the text if needed
        try:
            # First try direct parse
            return json.loads(content)
        except json.JSONDecodeError:
            # Try to extract JSON if wrapped in markdown code blocks
            json_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", content)
            if json_match:
                return json.loads(json_match.group(1))
            else:
                raise Exception("Could not parse JSON from Anthropic response")

    except Exception as e:
        raise Exception(f"Error processing with Anthropic: {str(e)}")


def add_to_spreadsheet(creds, json_data):
    """Add data from JSON to Google Sheets"""
    service = build("sheets", "v4", credentials=creds)

    # Ensure date is in the right format (YYYY-MM-DD)
    try:
        # Try to parse and standardize the date format
        date_obj = None
        date_formats = [
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%m/%d/%Y",
            "%d-%m-%Y",
            "%m-%d-%Y",
            "%d.%m.%Y",
            "%m.%d.%Y",
            "%d %b %Y",
            "%d %B %Y",
        ]

        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(json_data["date"], fmt)
                break
            except (ValueError, TypeError):
                continue

        if date_obj:
            formatted_date = date_obj.strftime("%Y-%m-%d")
        else:
            formatted_date = json_data.get("date", "")
    except Exception:
        formatted_date = json_data.get("date", "")

    # Create a row with the data
    row = [
        formatted_date,  # Date
        json_data.get("description", ""),  # Description
        json_data.get("amount", 0),  # Amount
        json_data.get("category", ""),  # Category
        json_data.get("property", ""),  # Property
    ]

    # Append the row to the spreadsheet
    result = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A:E",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        )
        .execute()
    )

    return result.get("updates").get("updatedRange")


def check_if_processed(creds, message_id):
    """Check if the email has already been processed"""
    try:
        service = build("sheets", "v4", credentials=creds)

        # Check if tracking sheet exists, if not create it
        try:
            service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID, range=f"{TRACKING_SHEET_NAME}!A1"
            ).execute()
        except Exception:
            # Create the tracking sheet with headers
            body = {
                "requests": [
                    {"addSheet": {"properties": {"title": TRACKING_SHEET_NAME}}}
                ]
            }
            service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID, body=body
            ).execute()

            # Add headers
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{TRACKING_SHEET_NAME}!A1:D1",
                valueInputOption="USER_ENTERED",
                body={
                    "values": [
                        ["MessageID", "ProcessedTimestamp", "EmailSubject", "Status"]
                    ]
                },
            ).execute()

        # Get all processed message IDs
        result = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=SPREADSHEET_ID, range=f"{TRACKING_SHEET_NAME}!A:A")
            .execute()
        )

        values = result.get("values", [])

        # Skip header row and check if message_id exists
        processed_ids = [row[0] for row in values[1:]] if len(values) > 1 else []

        return message_id in processed_ids

    except Exception as e:
        print(f"Error checking processed emails: {e}")
        # If there's an error checking, assume not processed so we don't miss anything
        return False


def mark_as_processed(creds, message_id, subject, status="Success"):
    """Mark email as processed in the tracking sheet"""
    try:
        service = build("sheets", "v4", credentials=creds)

        # Add to tracking sheet
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{TRACKING_SHEET_NAME}!A:D",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={
                "values": [
                    [
                        message_id,
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        subject,
                        status,
                    ]
                ]
            },
        ).execute()

        return True

    except Exception as e:
        print(f"Error marking as processed: {e}")
        return False


def send_email_summary(service, to_email, subject, body):
    """Send an email summary"""
    message = MIMEText(body)
    message["to"] = to_email
    message["subject"] = subject

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")

    try:
        service.users().messages().send(
            userId="me", body={"raw": raw_message}
        ).execute()
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False


def main():
    """Main function to process invoices and financial emails"""
    try:
        # Get credentials
        creds = get_credentials()

        # Build Gmail service
        gmail_service = build("gmail", "v1", credentials=creds)

        # Get label ID for the specified label
        results = gmail_service.users().labels().list(userId="me").execute()
        labels = results.get("labels", [])

        label_id = None
        for label in labels:
            if label["name"] == EMAIL_LABEL:
                label_id = label["id"]
                break

        if not label_id:
            raise Exception(f"Label '{EMAIL_LABEL}' not found in Gmail")

        # Check if Finance sheet exists, if not create it
        sheets_service = build("sheets", "v4", credentials=creds)
        try:
            sheets_service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID, range=f"{SHEET_NAME}!A1"
            ).execute()
        except Exception:
            # Create the finance sheet with headers
            body = {"requests": [{"addSheet": {"properties": {"title": SHEET_NAME}}}]}
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID, body=body
            ).execute()

            # Add headers
            sheets_service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{SHEET_NAME}!A1:E1",
                valueInputOption="USER_ENTERED",
                body={
                    "values": [
                        ["Date", "Description", "Amount", "Category", "Property"]
                    ]
                },
            ).execute()

        # Get emails with the specified label
        query = f"label:{EMAIL_LABEL}"
        results = gmail_service.users().messages().list(userId="me", q=query).execute()
        messages = results.get("messages", [])

        if not messages:
            print("No emails found with the specified label")
            return

        for message in messages:
            message_id = message["id"]

            # Check if already processed
            if check_if_processed(creds, message_id):
                print(f"Skipping already processed email: {message_id}")
                continue

            try:
                # Get email content and attachments
                email_content = get_email_content(gmail_service, message_id)

                # Process with Anthropic
                json_data = process_with_anthropic(
                    email_content, email_content["attachments"]
                )

                # Add to spreadsheet
                updated_range = add_to_spreadsheet(creds, json_data)

                # Create summary for email
                # Format the amount with proper sign and 2 decimal places
                amount = float(json_data.get("amount", 0))
                formatted_amount = f"${abs(amount):.2f}"
                if amount < 0:
                    formatted_amount = f"-{formatted_amount}"
                else:
                    formatted_amount = f"+{formatted_amount}"

                summary = f"""
Transaction Processing Summary:
-----------------------------
Date: {json_data.get('date', 'N/A')}
Description: {json_data.get('description', 'N/A')}
Amount: {formatted_amount}
Category: {json_data.get('category', 'N/A')}
Property: {json_data.get('property', 'N/A') if json_data.get('property') else 'Not property-related'}

Successfully added to Google Sheets at {updated_range}
                """

                send_email_summary(
                    gmail_service,
                    SENDER_EMAIL,
                    f"Transaction Processed: {email_content['subject']}",
                    summary,
                )

                # Mark as processed in our tracking sheet
                mark_as_processed(
                    creds, message_id, email_content["subject"], "Success"
                )

                # Mark as read in Gmail
                gmail_service.users().messages().modify(
                    userId="me", id=message_id, body={"removeLabelIds": ["UNREAD"]}
                ).execute()

                # Clean up temp files
                for attachment in email_content["attachments"]:
                    if os.path.exists(attachment["filepath"]):
                        os.unlink(attachment["filepath"])

            except Exception as e:
                error_message = f"""
Transaction Processing Error:
---------------------------
Subject: {email_content['subject'] if 'email_content' in locals() else 'Unknown'}
Error: {str(e)}

Please check the email and attachments manually.
                """

                send_email_summary(
                    gmail_service,
                    SENDER_EMAIL,
                    f"Transaction Processing Error: {email_content['subject'] if 'email_content' in locals() else 'Unknown'}",
                    error_message,
                )

                # Mark as processed but with error status
                if "email_content" in locals():
                    mark_as_processed(
                        creds, message_id, email_content["subject"], "Error"
                    )

                # Clean up temp files if they exist
                if "email_content" in locals() and "attachments" in email_content:
                    for attachment in email_content["attachments"]:
                        if os.path.exists(attachment["filepath"]):
                            os.unlink(attachment["filepath"])

    except Exception as e:
        print(f"Error in main function: {e}")


if __name__ == "__main__":
    main()
