# ai-receipts

This app is designed to automate rental property accounting by processing
emails with an LLM, extracting invoice data and adding to a Google Sheet.

- Reads emails from an IMAP endpoint (presently Google)
- Processes attachments and email content
- Sends these to the Anthropic API for analysis as invoices/statements.
- Adds extracted line items to a Google Sheet
- Replies to the emails with a summary of events

Whilst this app presently assumes *rental* invoices, it should be trivial for a
developer to modify the script to handle other types of invoices/statements.

The spreadsheet should have the following columns:

- `Date`
- `Description`
- `Amount`
- `Category`
- `Property`

It should be trivial for any developer to modify these columns in the Python
script.

## Requirements

Apart from Python on your local machine, you'll need three external services to
run this app:

- **A Gmail account**: you'll forward emails to this account for processing
    - You'll need an *app password* from Google for IMAP. There's plenty of guides online for this.
    - Any other IMAP account will also work, it doesn't need to be Google.
- **A Google Cloud Service Account**: this provides programmatic access to the Google Sheet
    - See below for setup details.
- **An Anthropic API Key**: This provides access to the Anthropic/Claude API.
    - Using this API will cost credits. Setting up this entire app and
      processing 14 invoices has cost me US$0.17 so far.
    - Use the Anthropic Console to generate an API key.

### Setting Up a Google Service Account

Follow these steps to setup a Google service account with access to the Google
Sheets API (you'll need this to run the app):

1. **Create a Google Cloud Project**
   - Go to the [Google Cloud Console](https://console.cloud.google.com/)
   - Click on "New Project" and give it a name
   - Create the project and select it

2. **Enable the Google Sheets API**
   - Navigate to "APIs & Services" > "Library"
   - Search for "Google Sheets API"
   - Select it and click "Enable"

3. **Create a Service Account**
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "Service Account"
   - Fill in a name, ID, and description for your service account
   - Click "Create and Continue"
   - Assign the role "Editor" (or a more specific role depending on your needs)
   - Click "Continue" and then "Done"

4. **Generate a Key for the Service Account**
   - Find your service account in the list and click on it
   - Go to the "Keys" tab
   - Click "Add Key" > "Create new key"
   - Choose JSON format
   - Click "Create" to download the key file

5. **Share Your Google Sheets**
   - Open the Google Sheet you want to access
   - Click the "Share" button
   - Add the email address of your service account (found in the downloaded JSON file)
   - Assign appropriate permissions (Editor or Viewer)

6. **Use the Service Account in Your Code**
   - Store the JSON key file securely
   - Place it in `./config/service-account.json`

## Running the App

You can run the app locally with `python main.py`, or use the Docker setup
described below. Running the project locally requires the same `./config`
directory as the Docker setup.

### Project Structure

Ensure you create and populate a `config` directory, so your project looks like
this:

```
project/
├── main.py                 # Your Python script
├── Dockerfile              # Docker configuration
├── docker-compose.yml      # Docker Compose configuration
├── requirements.txt        # Python dependencies
└── config/                 # Configuration directory (you need to create this)
    ├── config.yaml         # Your configuration file
    └── service-account.json # Google service account credentials
```

### Configuration Files

1. Place your `config.yaml` and `service-account.json` files in the `config/` directory.
2. Use `config.example.yaml` as the basis for `config.yaml`.

### Building and Running

#### Startup:

```bash
# Build and start the container
docker-compose up -d
```

#### Managing the application:

```bash
# View logs
docker-compose logs -f

# Stop the application
docker-compose down

# Restart the application
docker-compose restart
```

#### Updating the application:

If you make changes to your code:

```bash
# Rebuild and restart
docker-compose up -d --build
```

### Troubleshooting

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