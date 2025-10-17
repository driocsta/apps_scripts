# Apps Script Copier

A Node.js application that copies Google Apps Script code from a source spreadsheet to multiple target spreadsheets listed in a column.

## Features

- Reads spreadsheet IDs from a column in your source spreadsheet
- Automatically extracts spreadsheet IDs from URLs or uses direct IDs
- Copies all Apps Script files from the source spreadsheet to target spreadsheets
- Provides detailed progress and error reporting

## Prerequisites

1. **Node.js** (version 14 or higher)
2. **Google Cloud Project** with the following APIs enabled:
   - Google Sheets API
   - Google Apps Script API
   - Google Drive API
3. **Service Account** or **OAuth 2.0 credentials**

## Setup Instructions

### Step 1: Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the required APIs:
   - Google Sheets API
   - Google Apps Script API
   - Google Drive API

### Step 2: Create Service Account Credentials

1. In Google Cloud Console, go to **IAM & Admin** > **Service Accounts**
2. Click **Create Service Account**
3. Give it a name (e.g., "Apps Script Copier")
4. Click **Create and Continue**
5. Click **Done**
6. Click on the created service account
7. Go to the **Keys** tab
8. Click **Add Key** > **Create New Key**
9. Choose **JSON** format
10. Download the JSON file and save it as `credentials.json` in this project directory

### Step 3: Grant Permissions

You need to share your spreadsheets with the service account:

1. Open the `credentials.json` file
2. Find the `client_email` field (looks like `xxx@xxx.iam.gserviceaccount.com`)
3. Share your source spreadsheet with this email address (Editor access)
4. Share all target spreadsheets with this email address (Editor access)

### Step 4: Configure Environment Variables

1. Copy the example environment file:
   ```
   cp env.example .env
   ```

2. Edit `.env` and set your values:
   ```
   SOURCE_SPREADSHEET_ID=your_spreadsheet_id_here
   SHEET_NAME=Sheet1
   SHEET_ID_COLUMN=A
   START_ROW=2
   ```

   - `SOURCE_SPREADSHEET_ID`: The ID of your source spreadsheet (from the URL)
   - `SHEET_NAME`: The name of the sheet tab containing the spreadsheet IDs
   - `SHEET_ID_COLUMN`: The column letter containing the spreadsheet IDs (e.g., A, B, C)
   - `START_ROW`: The row to start reading from (default: 2, assuming row 1 is header)

### Step 5: Install Dependencies

```bash
npm install
```

### Step 6: Prepare Your Source Spreadsheet

Make sure your source spreadsheet has:
1. A column (e.g., column A) with the header "Sheet Id" (or any name)
2. Rows containing either:
   - Full spreadsheet URLs (e.g., `https://docs.google.com/spreadsheets/d/ABC123/edit`)
   - Just the spreadsheet IDs (e.g., `ABC123`)
3. An Apps Script project attached (Tools > Script editor in Google Sheets)

## Usage

Run the application:

```bash
npm start
```

The application will:
1. Read all spreadsheet IDs from the specified column
2. Extract the Apps Script code from your source spreadsheet
3. Copy that code to all target spreadsheets
4. Display a summary of successes and failures

## Spreadsheet ID Format

The app supports multiple formats in the "Sheet Id" column:
- Full URL: `https://docs.google.com/spreadsheets/d/1ABC123DEF456/edit`
- Short URL: `/spreadsheets/d/1ABC123DEF456`
- Just the ID: `1ABC123DEF456`

## Troubleshooting

### "Could not find Apps Script project"
- Make sure the spreadsheet has an Apps Script project attached
- Open the spreadsheet in Google Sheets
- Go to Extensions > Apps Script to create a project if needed

### "Permission denied"
- Ensure the service account email has Editor access to all spreadsheets
- Check that all required APIs are enabled in Google Cloud Console

### "Authentication failed"
- Verify `credentials.json` is in the project directory
- Ensure the service account key is valid

### "No spreadsheet IDs found"
- Check that `SHEET_NAME` matches your sheet tab name exactly
- Verify `SHEET_ID_COLUMN` points to the correct column
- Ensure `START_ROW` is set correctly

## Example

If your spreadsheet looks like this:

| Sheet Id | Name |
|----------|------|
| https://docs.google.com/spreadsheets/d/ABC123/edit | Project 1 |
| DEF456 | Project 2 |
| https://docs.google.com/spreadsheets/d/GHI789/edit | Project 3 |

With these settings:
```
SHEET_NAME=Sheet1
SHEET_ID_COLUMN=A
START_ROW=2
```

The app will copy your Apps Script code to all three spreadsheets (ABC123, DEF456, and GHI789).

## Security Notes

- **Never commit `credentials.json` or `.env` to version control**
- Add them to `.gitignore`
- Keep your service account credentials secure
- Only grant necessary permissions

## License

ISC

