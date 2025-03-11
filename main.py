import os
import base64
import email
import re
import imaplib
from email.header import decode_header
import anthropic
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import gspread
import tempfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
import datetime
import json
import PyPDF2
import docx
import csv
import io

# Set up logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Configuration
SERVICE_ACCOUNT_FILE = (
    "./credentials/service-account.json"  # Path to your service account credentials
)
GOOGLE_GROUP_EMAIL = "ai-receipts@paulhauner.com"  # Your Google Group email
SPREADSHEET_ID = "1oM5APBsN7JqADj71VHLp14BVwaOd5Azx9g3m_Cm5B8M"  # Google Sheet ID
WORKSHEET_NAME = "Transactions"  # Name of the worksheet
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")  # Anthropic API key

# Gmail IMAP configuration
GMAIL_EMAIL = "haunereceipts@gmail.com"  # Your dedicated Gmail account
GMAIL_PASSWORD = os.environ.get(
    "GMAIL_APP_PASSWORD"
)  # App password for Gmail (not your regular password)
GMAIL_IMAP_SERVER = "imap.gmail.com"
GMAIL_SMTP_SERVER = "smtp.gmail.com"
GMAIL_SMTP_PORT = 587
FORWARDING_EMAIL = "paul@paulhauner.com"  # Email to forward summaries to


class InvoiceProcessor:
    def __init__(self):
        # Set up Google Sheets access via service account
        self.credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        self.gc = gspread.authorize(self.credentials)
        self.spreadsheet = self.gc.open_by_key(SPREADSHEET_ID)
        self.worksheet = self.spreadsheet.worksheet(WORKSHEET_NAME)

        # Set up Anthropic client
        self.anthropic_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    def connect_to_gmail(self):
        """Connect to Gmail via IMAP."""
        try:
            mail = imaplib.IMAP4_SSL(GMAIL_IMAP_SERVER)
            mail.login(GMAIL_EMAIL, GMAIL_PASSWORD)
            return mail
        except Exception as e:
            logger.error(f"Error connecting to Gmail: {e}")
            return None

    def get_unread_emails(self):
        """Retrieve unread emails from Gmail."""
        mail = self.connect_to_gmail()
        if not mail:
            return []

        try:
            mail.select("inbox")
            status, messages = mail.search(None, "UNSEEN")

            if status != "OK":
                logger.error("Error searching for emails")
                return []

            message_ids = messages[0].split()
            logger.info(f"Found {len(message_ids)} unread emails.")

            return mail, message_ids
        except Exception as e:
            logger.error(f"Error retrieving emails: {e}")
            mail.logout()
            return []

    def decode_email_header(self, header):
        """Decode email header."""
        decoded_header = decode_header(header)
        header_parts = []
        for content, encoding in decoded_header:
            if isinstance(content, bytes):
                if encoding:
                    header_parts.append(content.decode(encoding))
                else:
                    header_parts.append(content.decode("utf-8", errors="replace"))
            else:
                header_parts.append(content)
        return "".join(header_parts)

    def extract_text_from_attachment(self, file_path, mime_type):
        """Extract text from various file types."""
        try:
            text = ""

            # Process based on mime type
            if mime_type == "application/pdf":
                # Extract text from PDF
                with open(file_path, "rb") as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    for page_num in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[page_num]
                        text += page.extract_text() + "\n\n"

            elif (
                mime_type
                == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ):
                # Extract text from DOCX
                doc = docx.Document(file_path)
                for para in doc.paragraphs:
                    text += para.text + "\n"

            elif mime_type == "text/plain":
                # Extract text from plain text file
                with open(file_path, "r", errors="replace") as txt_file:
                    text = txt_file.read()

            elif mime_type == "text/csv":
                # Extract text from CSV
                with open(file_path, "r", errors="replace") as csv_file:
                    csv_reader = csv.reader(csv_file)
                    for row in csv_reader:
                        text += ", ".join(row) + "\n"

            elif mime_type.startswith("image/"):
                # For images, we can't extract text directly
                text = "[This is an image attachment]"

            else:
                # For other file types
                text = f"[Attachment of type {mime_type}]"

            return text
        except Exception as e:
            logger.error(f"Error extracting text from attachment: {e}")
            return f"[Error extracting text: {str(e)}]"

    def get_email_content(self, mail, message_id):
        """Get the content of an email, including attachments."""
        try:
            status, message_data = mail.fetch(message_id, "(RFC822)")

            if status != "OK":
                logger.error(f"Error fetching email with ID {message_id}")
                return None

            raw_email = message_data[0][1]
            email_message = email.message_from_bytes(raw_email)

            # Get email headers
            subject = self.decode_email_header(email_message["Subject"] or "No Subject")
            sender = self.decode_email_header(email_message["From"] or "Unknown Sender")
            date = self.decode_email_header(email_message["Date"] or "")

            # Process email body and attachments
            body_content = ""
            attachments = []

            if email_message.is_multipart():
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    if (
                        content_type == "text/plain"
                        and "attachment" not in content_disposition
                    ):
                        # Get the email body
                        try:
                            body = part.get_payload(decode=True)
                            charset = part.get_content_charset() or "utf-8"
                            body_content += body.decode(charset, errors="replace")
                        except Exception as e:
                            logger.error(f"Error decoding email body: {e}")

                    elif "attachment" in content_disposition or part.get_filename():
                        # This is an attachment
                        filename = part.get_filename()
                        if filename:
                            # Decode filename if needed
                            filename = self.decode_email_header(filename)

                            # Get attachment content
                            attachment_data = part.get_payload(decode=True)

                            # Save attachment to a temporary file
                            with tempfile.NamedTemporaryFile(
                                delete=False, suffix=os.path.splitext(filename)[1]
                            ) as temp:
                                temp.write(attachment_data)
                                attachments.append(
                                    {
                                        "filename": filename,
                                        "path": temp.name,
                                        "mime_type": part.get_content_type(),
                                    }
                                )
            else:
                # Not multipart - plain text email
                try:
                    body = email_message.get_payload(decode=True)
                    charset = email_message.get_content_charset() or "utf-8"
                    body_content = body.decode(charset, errors="replace")
                except Exception as e:
                    logger.error(f"Error decoding email body: {e}")

            # Mark the message as read
            mail.store(message_id, "+FLAGS", "\\Seen")

            return {
                "id": message_id.decode(),
                "subject": subject,
                "sender": sender,
                "date": date,
                "body": body_content,
                "attachments": attachments,
            }
        except Exception as e:
            logger.error(f"Error processing email {message_id}: {e}")
            return None

    def analyze_with_anthropic(self, email_content):
        """Send email content and attachments to Anthropic API for analysis."""
        try:
            # Process and extract text from attachments
            attachment_texts = []
            for attachment in email_content["attachments"]:
                try:
                    # Extract text from the attachment
                    extracted_text = self.extract_text_from_attachment(
                        attachment["path"], attachment["mime_type"]
                    )

                    attachment_texts.append(
                        {"filename": attachment["filename"], "content": extracted_text}
                    )

                    logger.info(
                        f"Successfully extracted text from {attachment['filename']}"
                    )
                except Exception as e:
                    logger.error(
                        f"Error extracting text from {attachment['filename']}: {e}"
                    )
                    attachment_texts.append(
                        {
                            "filename": attachment["filename"],
                            "content": f"[Error extracting content: {str(e)}]",
                        }
                    )

            # Prepare the prompt for Anthropic with email content
            prompt = f"""
I need you to analyze this email and any attachments related to rental property invoices or statements.
Extract line items and categorize them appropriately for accounting purposes.

EMAIL SUBJECT: {email_content['subject']}
EMAIL DATE: {email_content['date']}
EMAIL BODY:
{email_content['body']}

"""

            # Add attachment content to the prompt
            for attachment in attachment_texts:
                prompt += f"\n\nATTACHMENT: {attachment['filename']}\n"
                prompt += f"CONTENT:\n{attachment['content']}\n"

            prompt += """
For each line item you identify, please provide:
1. Date (in YYYY-MM-DD format)
2. Description (what the charge or payment is for)
3. Amount (negative for expenses, positive for income)
4. Category (e.g., Utilities, Repairs, Rent)
5. Property (if a specific property address is mentioned)

Format your response as JSON objects in the following structure:
[
  {
    "date": "YYYY-MM-DD",
    "description": "Description of item",
    "amount": 123.45,
    "category": "Category",
    "property": "Property address or empty if not specified"
  }
]
"""

            # Log prompt size for debugging
            logger.info(f"Prompt size: {len(prompt)} characters")

            # Call Anthropic API
            response = self.anthropic_client.messages.create(
                model="claude-3-7-sonnet-20250219",
                max_tokens=4000,
                temperature=0,
                system="You are an expert accountant specialized in processing rental property invoices and statements. Extract line items accurately, following the format instructions exactly.",
                messages=[{"role": "user", "content": prompt}],
            )

            # Extract the JSON part from the response
            # Attempt to find a JSON array in the response
            json_match = re.search(
                r"\[\s*\{.*?\}\s*\]", response.content[0].text, re.DOTALL
            )
            if json_match:
                try:
                    line_items = json.loads(json_match.group(0))
                    return line_items
                except json.JSONDecodeError:
                    logger.error("Could not parse JSON from Anthropic response")
                    logger.error(f"Raw response text: {response.content[0].text}")
                    return []
            else:
                logger.error("No JSON data found in Anthropic response")
                logger.error(f"Raw response text: {response.content[0].text}")
                return []

        except Exception as e:
            logger.error(f"Error analyzing with Anthropic: {e}")
            return []
        finally:
            # Clean up temporary attachment files
            for attachment in email_content["attachments"]:
                try:
                    os.unlink(attachment["path"])
                except Exception as e:
                    logger.error(
                        f"Error deleting temporary file {attachment['path']}: {e}"
                    )

    def add_to_spreadsheet(self, line_items):
        """Add the analyzed line items to the Google Sheet."""
        added_rows = []
        errors = []

        try:
            for item in line_items:
                # Validate and format the data
                try:
                    # Convert string date to datetime object for validation
                    date_obj = datetime.datetime.strptime(item["date"], "%Y-%m-%d")
                    formatted_date = date_obj.strftime("%Y-%m-%d")

                    # Ensure amount is a float
                    amount = float(item["amount"])

                    # Prepare row data
                    row_data = [
                        formatted_date,
                        item["description"],
                        amount,
                        item["category"],
                        item.get("property", ""),
                    ]

                    # Add to spreadsheet
                    self.worksheet.append_row(row_data)
                    added_rows.append(row_data)

                except (ValueError, KeyError) as e:
                    errors.append(f"Error adding item {item}: {str(e)}")
                    logger.error(f"Error adding item to spreadsheet: {e}")

            return added_rows, errors
        except Exception as e:
            logger.error(f"Error updating spreadsheet: {e}")
            return added_rows, [f"General error updating spreadsheet: {str(e)}"]

    def send_summary_email(self, email_data, added_rows, errors):
        """Send a summary email with the results."""
        try:
            # Set up SMTP connection
            server = smtplib.SMTP(GMAIL_SMTP_SERVER, GMAIL_SMTP_PORT)
            server.starttls()
            server.login(GMAIL_EMAIL, GMAIL_PASSWORD)

            msg = MIMEMultipart()
            msg["To"] = FORWARDING_EMAIL
            msg["From"] = GMAIL_EMAIL
            msg["Subject"] = f"Invoice Processing Summary: {email_data['subject']}"

            email_body = f"""
            <html>
            <body>
            <h2>Invoice Processing Summary</h2>
            <p><strong>Original Email:</strong> {email_data['subject']}</p>
            <p><strong>From:</strong> {email_data['sender']}</p>
            <p><strong>Date:</strong> {email_data['date']}</p>
            <p><strong>Attachments:</strong> {', '.join([a['filename'] for a in email_data['attachments']])}</p>
            
            <h3>Processed Items:</h3>
            """

            if added_rows:
                email_body += "<table border='1' cellpadding='5'>"
                email_body += "<tr><th>Date</th><th>Description</th><th>Amount</th><th>Category</th><th>Property</th></tr>"

                for row in added_rows:
                    email_body += f"<tr><td>{row[0]}</td><td>{row[1]}</td><td>{row[2]}</td><td>{row[3]}</td><td>{row[4]}</td></tr>"

                email_body += "</table>"
            else:
                email_body += "<p>No items were processed.</p>"

            if errors:
                email_body += "<h3>Errors:</h3><ul>"
                for error in errors:
                    email_body += f"<li>{error}</li>"
                email_body += "</ul>"

            email_body += """
            </body>
            </html>
            """

            msg.attach(MIMEText(email_body, "html"))

            # Send the message
            server.sendmail(GMAIL_EMAIL, FORWARDING_EMAIL, msg.as_string())
            server.quit()

            logger.info(f"Summary email sent to {FORWARDING_EMAIL}")

        except Exception as e:
            logger.error(f"Error sending summary email: {e}")

    def process_emails(self):
        """Main function to process all unread emails."""
        result = self.get_unread_emails()
        if not result or len(result) != 2:
            logger.info("No unread emails to process or connection failed.")
            return

        mail, message_ids = result

        try:
            if not message_ids:
                logger.info("No unread emails to process.")
                return

            for message_id in message_ids:
                logger.info(f"Processing email ID: {message_id}")

                # Get email content
                email_content = self.get_email_content(mail, message_id)
                if not email_content:
                    continue

                # Analyze with Anthropic
                line_items = self.analyze_with_anthropic(email_content)

                # Add to spreadsheet
                added_rows, errors = self.add_to_spreadsheet(line_items)

                # Send summary email
                self.send_summary_email(email_content, added_rows, errors)

                logger.info(f"Completed processing email ID: {email_content['id']}")
        finally:
            # Close IMAP connection
            try:
                mail.close()
                mail.logout()
            except Exception as e:
                logger.error(f"Error closing IMAP connection: {e}")


if __name__ == "__main__":
    processor = InvoiceProcessor()
    processor.process_emails()
