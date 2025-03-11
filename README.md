# ai-receipts

## Instructions

### Install Required Packages

```bash
pip install -r requirements.txt
```

### Get Google API credentials

1. Go to Google Cloud Console
1. Create a new project
1. Enable both Gmail API and Google Sheets API
1. Create OAuth credentials (Desktop application)
1. Download the credentials and save as credentials.json in the same directory as the script

### Set your Anthropic API Key as an environment variable

```bash
export ANTHROPIC_API_KEY=your_api_key_here
```

### Update the configuration variables

- `EMAIL_LABEL`: The Gmail label for invoice emails
- `SENDER_EMAIL`: Your email address
- `SPREADSHEET_ID`: ID of your Google Sheet (found in the URL)
- `SHEET_NAME`: Name of the sheet tab