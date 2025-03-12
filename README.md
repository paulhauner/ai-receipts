# ai-receipts

## Running

# Setup and Running Instructions

## Project Structure
Ensure your project directory is set up like this:

```
project/
├── main.py                 # Your Python script
├── Dockerfile              # Docker configuration
├── docker-compose.yml      # Docker Compose configuration
├── requirements.txt        # Python dependencies
└── config/                 # Configuration directory
    ├── config.yaml         # Your configuration file
    └── service-account.json # Google service account credentials
```

## Configuration Files

1. Place your `config.yaml` and `service-account.json` files in the `config/` directory.
2. Use `config.example.yaml` as the basis for `config.yaml`.

## Building and Running

### First-time setup:

```bash
# Create the config directory if it doesn't exist
mkdir -p config

# Copy your configuration files into the config directory
cp /path/to/your/config.yaml config/
cp /path/to/your/service-account.json config/

# Build and start the container
docker-compose up -d
```

### Managing the application:

```bash
# View logs
docker-compose logs -f

# Stop the application
docker-compose down

# Restart the application
docker-compose restart
```

### Updating the application:

If you make changes to your code:

```bash
# Rebuild and restart
docker-compose up -d --build
```

## Troubleshooting

- If you encounter permissions issues, ensure your configuration files have appropriate read permissions:
  ```bash
  chmod 644 config/config.yaml config/service-account.json
  ```

- To check if the container is running properly:
  ```bash
  docker ps
  ```

- To execute commands inside the container:
  ```bash
  docker exec -it invoice-processor bash
  ```

## Instructions

I'll create a Python script that can accomplish this task. The script will:

Read emails from a Google Group
Process attachments and email content
Send these to the Anthropic API for analysis
Add extracted line items to a Google Sheet
Forward you a summary of the additions

This script is designed to automate your rental property accounting workflow by processing emails from a Google Group and extracting invoice data to your spreadsheet. Here's how it works:
Setup Requirements

A Google service account with necessary permissions
Your Anthropic API key
A Google Sheet with the requested columns
Access to the Google Group emails
The "Google Sheets API" enabled in the Google Cloud workspace

Key Features

Retrieves unread emails from your specified Google Group
Processes email content and attachments
Uses Anthropic's Claude 3.7 Sonnet model to extract:

Date
Description
Amount (negative for expenses, positive for income)
Category
Property (when specific addresses are mentioned)


Adds extracted line items to your Google Sheet
Sends you a summary email with all processed items or any errors

How to Use

Replace the placeholder values in the configuration section:

SERVICE_ACCOUNT_FILE: Path to your service account JSON file
GOOGLE_GROUP_EMAIL: Your Google Group email address
SPREADSHEET_ID: Your Google Sheet ID
WORKSHEET_NAME: Name of the worksheet
FORWARDING_EMAIL: Your email for receiving summaries
ANTHROPIC_API_KEY: Your Anthropic API key


Install the required dependencies:
Copypip install google-api-python-client google-auth-httplib2 google-auth-oauthlib gspread anthropic

Set up a scheduled task to run this script regularly (e.g., daily or hourly)

Would you like me to explain any specific part of the code in more detail?

## Updated for IMAP settings

Updated Gmail Invoice Processor
Key Changes

IMAP Connection: Now uses Python's imaplib library to connect directly to Gmail's IMAP server
Authentication: Uses Gmail username and app password instead of service account
Email Processing: Retrieves and processes unread emails via IMAP protocol
Email Sending: Uses SMTP protocol to send summary emails from the Gmail account
Simplified Permissions: Service account is only used for Google Sheets access, not for Gmail

Configuration Updates
You'll need to configure these new settings:

GMAIL_EMAIL: Your dedicated Gmail account address
GMAIL_PASSWORD: An app password generated specifically for this application (not your regular Gmail password)
GMAIL_IMAP_SERVER: Set to 'imap.gmail.com'
GMAIL_SMTP_SERVER: Set to 'smtp.gmail.com'
GMAIL_SMTP_PORT: Set to 587 for TLS

Important Notes:

App Password: You'll need to generate an "App Password" in your Google Account settings:

Go to your Google Account → Security → 2-Step Verification → App passwords
Create a new app password for this script


Requirements: You'll need these Python libraries:
Copypip install imaplib email google-api-python-client google-auth gspread anthropic

Service Account: You still need a service account for Google Sheets access, but with fewer permissions

The rest of the functionality remains the same - the script will still analyze emails with Anthropic, add entries to your Google Sheet, and send you summary emails.