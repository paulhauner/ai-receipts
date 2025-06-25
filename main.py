import os
import email
import re
import imaplib2
from email.header import decode_header
import anthropic
from google.oauth2 import service_account
import gspread
import tempfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import logging
import datetime
import json
import PyPDF2
import docx
import csv
import time
import threading
import yaml
import requests
from urllib.parse import urlparse

# Set up logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def load_config():
    """Load configuration from YAML file."""
    config_path = os.environ.get("CONFIG_PATH", "./config.yaml")
    try:
        with open(config_path, "r") as config_file:
            config = yaml.safe_load(config_file)
        logger.info(f"Configuration loaded from {config_path}")
        return config
    except Exception as e:
        logger.error(f"Error loading configuration: {e}")
        raise


class InvoiceProcessor:
    def __init__(self):
        # Load configuration
        self.config = load_config()

        # Set up Google Sheets access via service account
        self.credentials = service_account.Credentials.from_service_account_file(
            self.config["service_account_file"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        self.gc = gspread.authorize(self.credentials)
        self.spreadsheet = self.gc.open_by_key(self.config["spreadsheet_id"])
        self.worksheet = self.spreadsheet.worksheet(self.config["worksheet_name"])

        # Set up Anthropic client
        self.anthropic_client = anthropic.Anthropic(
            api_key=self.config["anthropic_api_key"]
        )

        # IMAP connection
        self.mail = None
        self.idle_event = threading.Event()

    def connect_to_gmail(self):
        """Connect to Gmail via IMAP."""
        try:
            mail = imaplib2.IMAP4_SSL(self.config["gmail_imap_server"])
            mail.login(self.config["gmail_email"], self.config["gmail_app_password"])
            return mail
        except Exception as e:
            logger.error(f"Error connecting to Gmail: {e}")
            return None

    def process_new_message(self, message_id):
        """Process a single new email message."""
        logger.info(f"Processing email ID: {message_id}")

        try:
            # Get email content
            email_content = self.get_email_content(self.mail, message_id)
            if not email_content:
                return

            # Analyze with Anthropic
            line_items = self.analyze_with_anthropic(email_content)

            # Add to spreadsheet
            added_rows, errors = self.add_to_spreadsheet(line_items)

            # Send summary email
            self.send_summary_email(email_content, added_rows, errors)

            logger.info(f"Completed processing email ID: {email_content['id']}")
        except Exception as e:
            logger.error(f"Error processing message {message_id}: {e}")
        finally:
            # Clean up temporary attachment files
            if 'email_content' in locals() and email_content:
                for attachment in email_content.get("attachments", []):
                    try:
                        os.unlink(attachment["path"])
                    except Exception as e:
                        logger.error(f"Error deleting temporary file {attachment['path']}: {e}")

    def get_unread_emails(self):
        """Retrieve unread emails from Gmail."""
        if not self.mail:
            return []

        try:
            self.mail.select("inbox")
            status, messages = self.mail.search(None, "UNSEEN")

            if status != "OK":
                logger.error("Error searching for emails")
                return []

            message_ids = messages[0].split()
            count = len(message_ids)
            if count > 0:
                logger.info(f"Found {count} unread emails.")

            return message_ids
        except Exception as e:
            logger.error(f"Error retrieving emails: {e}")
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

    def find_pdf_urls(self, text):
        """Find PDF URLs in email text, including HTML links."""
        urls = []
        
        # First, extract URLs from HTML href attributes
        href_pattern = r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>([^<]*(?:\.pdf|invoice|statement|receipt|document)[^<]*)</a>'
        html_matches = re.findall(href_pattern, text, re.IGNORECASE)
        for url, link_text in html_matches:
            logger.info(f"Found HTML link: '{link_text}' -> {url}")
            urls.append(url)
        
        # Also try to extract just href URLs without requiring specific link text
        simple_href_pattern = r'href=["\']([^"\']+)["\']'
        href_urls = re.findall(simple_href_pattern, text, re.IGNORECASE)
        for url in href_urls:
            # Only include if it looks like it could be a document
            if any(keyword in url.lower() for keyword in ['document', 'statement', 'invoice', 'receipt', 'download', 'file', 'attachment']):
                logger.info(f"Found potential document URL in href: {url}")
                urls.append(url)
        
        # Pattern to match URLs that end with .pdf or contain pdf in the path
        pdf_url_patterns = [
            r'https?://[^\s<>"]+\.pdf(?:\?[^\s<>"]*)?',  # URLs ending with .pdf
            r'https?://[^\s<>"]*[/\?&]pdf[/\?&][^\s<>"]*',  # URLs with pdf in path
            r'https?://[^\s<>"]*(?:document|statement|invoice|receipt|download|file|attachment)[^\s<>"]*',  # Document-related URLs
        ]
        
        for pattern in pdf_url_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            urls.extend(matches)
        
        # Remove duplicates while preserving order
        unique_urls = []
        for url in urls:
            if url not in unique_urls:
                unique_urls.append(url)
        
        logger.info(f"All detected URLs: {unique_urls}")
        return unique_urls

    def download_pdf_from_url(self, url):
        """Download PDF from URL and save to temporary file."""
        try:
            logger.info(f"Downloading PDF from URL: {url}")
            
            # Set up headers to mimic a browser request
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            # Download the file with timeout
            response = requests.get(url, headers=headers, timeout=30, stream=True)
            response.raise_for_status()
            
            # Check if the response might be a PDF (be more lenient)
            content_type = response.headers.get('content-type', '').lower()
            content_length = response.headers.get('content-length', '0')
            logger.info(f"Downloaded content: type={content_type}, length={content_length} bytes")
            
            # Be more lenient - many document servers don't set correct content-type
            if 'html' in content_type and int(content_length or 0) < 10000:
                logger.warning(f"URL {url} returned HTML content, likely not a direct document link")
                return None
            
            # Create temporary file with .pdf extension
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            
            # Download in chunks to handle large files
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    temp_file.write(chunk)
            
            temp_file.close()
            
            # Generate a filename from the URL
            parsed_url = urlparse(url)
            filename = os.path.basename(parsed_url.path)
            if not filename or not filename.endswith('.pdf'):
                filename = f"downloaded_pdf_{int(time.time())}.pdf"
            
            logger.info(f"Successfully downloaded PDF: {filename}")
            return {
                'filename': filename,
                'path': temp_file.name,
                'mime_type': 'application/pdf',
                'source_url': url
            }
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error downloading PDF from {url}: {e}")
            return None
        except Exception as e:
            logger.error(f"Unexpected error downloading PDF from {url}: {e}")
            return None

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
            downloaded_pdfs = []

            if email_message.is_multipart():
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    if (
                        content_type in ["text/plain", "text/html"]
                        and "attachment" not in content_disposition
                    ):
                        # Get the email body (both plain text and HTML)
                        try:
                            body = part.get_payload(decode=True)
                            charset = part.get_content_charset() or "utf-8"
                            decoded_body = body.decode(charset, errors="replace")
                            body_content += decoded_body + "\n\n"
                            logger.debug(f"Added {content_type} content, length: {len(decoded_body)}")
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

            # Look for PDF URLs in the email body and download them as attachments
            if body_content:
                logger.info(f"Scanning email body for PDF URLs. Body length: {len(body_content)} chars")
                logger.debug(f"Email body content preview: {body_content[:500]}...")
                pdf_urls = self.find_pdf_urls(body_content)
                if pdf_urls:
                    logger.info(f"Found {len(pdf_urls)} PDF URLs in email body: {pdf_urls}")
                    for url in pdf_urls:
                        downloaded_pdf = self.download_pdf_from_url(url)
                        if downloaded_pdf:
                            downloaded_pdfs.append(downloaded_pdf)
                            attachments.append(downloaded_pdf)
                            logger.info(f"Added downloaded PDF as attachment: {downloaded_pdf['filename']}")
                        else:
                            logger.warning(f"Failed to download PDF from URL: {url}")
                else:
                    logger.info("No PDF URLs found in email body")
            else:
                logger.warning("Email body is empty, cannot search for PDF URLs")

            # Mark the message as read
            mail.store(message_id, "+FLAGS", "\\Seen")

            email_data = {
                "id": (
                    message_id.decode() if isinstance(message_id, bytes) else message_id
                ),
                "subject": subject,
                "sender": sender,
                "date": date,
                "body": body_content,
                "attachments": attachments,
                "downloaded_pdfs": downloaded_pdfs,
            }

            if "References" in email_message:
                email_data["references"] = email_message["References"]

            return email_data
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
            additional_prompt = self.config["additional_prompt"]
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

            prompt += f"\n\n{additional_prompt}"

            # Log prompt size for debugging
            logger.info(f"Prompt size: {len(prompt)} characters")

            # Call Anthropic API
            response = self.anthropic_client.messages.create(
                model=self.config["anthropic_model"],
                max_tokens=self.config["max_tokens"],
                temperature=self.config["temperature"],
                system=self.config["system_prompt"],
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
            server = smtplib.SMTP(
                self.config["gmail_smtp_server"], self.config["gmail_smtp_port"]
            )
            server.starttls()
            server.login(self.config["gmail_email"], self.config["gmail_app_password"])

            msg = MIMEMultipart()
            msg["To"] = self.config["forwarding_email"]
            msg["From"] = self.config["gmail_email"]
            if not email_data["subject"].startswith("Re: "):
                msg["Subject"] = f"Re: {email_data['subject']}"
            else:
                msg["Subject"] = email_data["subject"]
            msg["In-Reply-To"] = email_data["id"]
            if "references" in email_data:
                msg["References"] = f"{email_data['references']} {email_data['id']}"
            else:
                msg["References"] = email_data["id"]

            email_body = f"""
            <html>
            <body>
            """

            if added_rows:
                email_body += f"""
                <h3>👍 Added to the <a href="https://docs.google.com/spreadsheets/d/{self.config["spreadsheet_id"]}">Google Sheet<a> ('{self.config["worksheet_name"]}' worksheet)</h3>
                """
                email_body += "<table border='1' cellpadding='5'>"
                email_body += "<tr><th>Date</th><th>Description</th><th>Amount</th>\
                    <th>Category</th><th>Property</th></tr>"

                for row in added_rows:
                    email_body += f"<tr><td>{row[0]}</td><td>{row[1]}</td>\
                        <td>{row[2]}</td><td>{row[3]}</td><td>{row[4]}</td></tr>"

                email_body += "</table>"
            else:
                email_body += "<p>🤔 No items were processed.</p>"

            if errors:
                email_body += "<h2>🚨 Errors:</h2><ul>"
                for error in errors:
                    email_body += f"<li>{error}</li>"
                email_body += "</ul>"

            email_body += """
            </body>
            </html>
            """

            msg.attach(MIMEText(email_body, "html"))

            # Attach downloaded PDFs to the email
            downloaded_pdfs = email_data.get("downloaded_pdfs", [])
            if downloaded_pdfs:
                logger.info(f"Attaching {len(downloaded_pdfs)} downloaded PDFs to reply email")
                for pdf in downloaded_pdfs:
                    try:
                        with open(pdf["path"], "rb") as f:
                            pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")
                            pdf_attachment.add_header(
                                "Content-Disposition", 
                                f"attachment; filename={pdf['filename']}"
                            )
                            msg.attach(pdf_attachment)
                            logger.info(f"Attached PDF: {pdf['filename']}")
                    except Exception as e:
                        logger.error(f"Error attaching PDF {pdf['filename']}: {e}")

            # Send the message
            server.sendmail(
                self.config["gmail_email"],
                self.config["forwarding_email"],
                msg.as_string(),
            )
            server.quit()

            logger.info(f"Summary email sent to {self.config['forwarding_email']}")

        except Exception as e:
            logger.error(f"Error sending summary email: {e}")

    def process_pending_emails(self):
        """Process any pending unread emails."""
        message_ids = self.get_unread_emails()

        if not message_ids:
            return

        for message_id in message_ids:
            self.process_new_message(message_id)

    def idle_callback(self, args):
        """Callback function for IMAP IDLE events."""
        response, data, error = args
        if response == "OK" and data[0].endswith(b"EXISTS"):
            logger.info("New email notification received via IDLE")
            return True
        return False

    def listen_for_emails(self):
        """
        Listen for new emails using IMAP IDLE with imaplib2.
        This method will run indefinitely, processing emails as they arrive.
        """
        reconnect_attempts = 0

        while reconnect_attempts < self.config["max_reconnect_attempts"]:
            try:
                # Connect to Gmail
                logger.info("Connecting to Gmail IMAP server...")
                self.mail = self.connect_to_gmail()

                if not self.mail:
                    logger.error("Failed to connect to Gmail. Retrying...")
                    reconnect_attempts += 1
                    time.sleep(self.config["reconnect_delay"])
                    continue

                # Reset reconnect counter on successful connection
                reconnect_attempts = 0

                # Select the inbox
                self.mail.select("inbox")

                # Process any existing unread emails first
                self.process_pending_emails()

                logger.info("Starting IMAP IDLE mode, waiting for new emails...")

                # Reset the event flag
                self.idle_event.clear()

                while not self.idle_event.is_set():
                    # Start IDLE mode with callback
                    self.mail.idle(
                        callback=self.idle_callback, timeout=self.config["idle_timeout"]
                    )

                    # Process new emails if the callback was triggered
                    self.process_pending_emails()

                    # Keep connection alive with NOOP
                    status, data = self.mail.noop()
                    if status != "OK":
                        logger.warning("NOOP failed, reconnecting...")
                        self.cleanup_connection()
                        break

            except Exception as e:
                logger.error(f"Unexpected error in IDLE loop: {e}")
                self.cleanup_connection()
                reconnect_attempts += 1
                time.sleep(self.config["reconnect_delay"])

        logger.critical(
            f"Failed to reconnect after {self.config['max_reconnect_attempts']} attempts. Exiting."
        )

    def cleanup_connection(self):
        """Clean up IMAP connection if it exists."""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
            except:
                pass  # Ignore errors during cleanup
            self.mail = None


if __name__ == "__main__":
    processor = InvoiceProcessor()

    try:
        logger.info("Starting Invoice Processor with IMAP IDLE monitoring")
        processor.listen_for_emails()
    except KeyboardInterrupt:
        logger.info("Received keyboard interrupt, shutting down...")
        processor.cleanup_connection()
        logger.info("Invoice Processor shut down gracefully")
    except Exception as e:
        logger.critical(f"Fatal error: {e}")
        processor.cleanup_connection()
