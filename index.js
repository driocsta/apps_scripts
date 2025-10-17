const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

// Configuration
const SOURCE_SPREADSHEET_ID = process.env.SOURCE_SPREADSHEET_ID;
const SOURCE_SCRIPT_ID = process.env.SOURCE_SCRIPT_ID; // Optional: manual script ID
const SHEET_NAME = process.env.SHEET_NAME || 'Sheet Id';
const SHEET_ID_COLUMN = process.env.SHEET_ID_COLUMN || 'D';
const SCRIPT_ID_COLUMN = process.env.SCRIPT_ID_COLUMN || 'E'; // Optional: column with script IDs
const START_ROW = parseInt(process.env.START_ROW) || 2;

// Initialize Google APIs
const sheets = google.sheets('v4');
const script = google.script('v1');

/**
 * Authorize using OAuth2 (user authentication)
 */
async function authorize() {
  const fs = require('fs').promises;
  const path = require('path');
  
  // Check if we have saved tokens
  const TOKEN_PATH = path.join(__dirname, 'token.json');
  const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
  
  let credentials;
  try {
    const content = await fs.readFile(CREDENTIALS_PATH);
    credentials = JSON.parse(content);
  } catch (err) {
    console.error('Error loading credentials.json');
    throw err;
  }

  const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token
  try {
    const token = await fs.readFile(TOKEN_PATH);
    oAuth2Client.setCredentials(JSON.parse(token));
    return oAuth2Client;
  } catch (err) {
    return getNewToken(oAuth2Client, TOKEN_PATH);
  }
}

/**
 * Get new OAuth2 token
 */
async function getNewToken(oAuth2Client, tokenPath) {
  const fs = require('fs').promises;
  const readline = require('readline');
  
  const scopes = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/drive.readonly',
  ];

  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: scopes,
  });

  console.log('\n' + '='.repeat(70));
  console.log('AUTHORIZATION REQUIRED');
  console.log('='.repeat(70));
  console.log('\nPlease visit this URL to authorize this application:');
  console.log('\n' + authUrl + '\n');
  
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve, reject) => {
    rl.question('Enter the code from that page here: ', async (code) => {
      rl.close();
      try {
        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        
        // Store the token for future use
        await fs.writeFile(tokenPath, JSON.stringify(tokens));
        console.log('Token stored to', tokenPath);
        console.log('='.repeat(70) + '\n');
        
        resolve(oAuth2Client);
      } catch (err) {
        console.error('Error retrieving access token', err);
        reject(err);
      }
    });
  });
}

/**
 * Extract spreadsheet ID from various URL formats or return as-is if already an ID
 */
function extractSpreadsheetId(input) {
  if (!input) return null;
  
  // If it's already just an ID (no slashes or http), return it
  if (!input.includes('/') && !input.includes('http')) {
    return input.trim();
  }
  
  // Match spreadsheet ID from URL patterns
  const patterns = [
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
    /spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
    /\/d\/([a-zA-Z0-9-_]+)/
  ];
  
  for (const pattern of patterns) {
    const match = input.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return input.trim();
}

/**
 * Read the "Sheet Id" and optional "Script Id" columns from the source spreadsheet
 */
async function readSheetIds(auth) {
  console.log(`Reading sheet IDs from ${SOURCE_SPREADSHEET_ID}...`);
  
  // Read both columns: Sheet ID and Script ID
  const range = `${SHEET_NAME}!${SHEET_ID_COLUMN}${START_ROW}:${SCRIPT_ID_COLUMN}`;
  
  const response = await sheets.spreadsheets.values.get({
    auth,
    spreadsheetId: SOURCE_SPREADSHEET_ID,
    range,
  });

  const rows = response.data.values || [];
  const spreadsheetData = rows
    .map(row => ({
      spreadsheetId: extractSpreadsheetId(row[0]),
      scriptId: row[1] ? row[1].trim() : null // Script ID from column B (or specified column)
    }))
    .filter(item => item.spreadsheetId && item.spreadsheetId.length > 0);

  console.log(`Found ${spreadsheetData.length} spreadsheet(s) to update`);
  return spreadsheetData;
}

/**
 * Get the script project ID associated with a spreadsheet
 */
async function getScriptProjectId(auth, spreadsheetId) {
  try {
    const drive = google.drive({ version: 'v3', auth });
    
    // Get the spreadsheet metadata
    const spreadsheet = await sheets.spreadsheets.get({
      auth,
      spreadsheetId,
    });
    
    // The script project ID is typically in the format: script_<spreadsheetId>
    // But we need to find it through the Drive API
    const response = await drive.files.list({
      q: `'${spreadsheetId}' in parents and mimeType='application/vnd.google-apps.script'`,
      fields: 'files(id, name)',
    });

    if (response.data.files && response.data.files.length > 0) {
      return response.data.files[0].id;
    }
    
    return null;
  } catch (error) {
    console.error(`Error getting script project for spreadsheet ${spreadsheetId}:`, error.message);
    return null;
  }
}

/**
 * Get all script files from a script project
 */
async function getScriptContent(auth, scriptProjectId) {
  try {
    const response = await script.projects.getContent({
      auth,
      scriptId: scriptProjectId,
    });

    return response.data;
  } catch (error) {
    console.error(`Error reading script content from ${scriptProjectId}:`, error.message);
    throw error;
  }
}

/**
 * Update script content in a target script project
 * This will overwrite the existing Code.gs file with the source content
 */
async function updateScriptContent(auth, scriptProjectId, sourceContent) {
  try {
    // First, get the existing content to preserve the manifest and file structure
    const existingContent = await script.projects.getContent({
      auth,
      scriptId: scriptProjectId,
    });

    // Find the Code.gs or Código.gs file from source
    const sourceCodeFile = sourceContent.files.find(f => 
      f.name === 'Code' || f.name === 'Código' || f.name.includes('Code')
    );

    if (!sourceCodeFile) {
      throw new Error('No Code.gs file found in source project');
    }

    // Create updated files array, replacing Code.gs content
    const updatedFiles = existingContent.data.files.map(file => {
      // If this is the Code.gs file (or similar), replace its content
      if (file.name === 'Code' || file.name === 'Código' || file.type === 'SERVER_JS') {
        return {
          name: file.name,
          type: file.type,
          source: sourceCodeFile.source
        };
      }
      // Keep manifest and other files as-is
      return file;
    });

    // Update the script project
    const response = await script.projects.updateContent({
      auth,
      scriptId: scriptProjectId,
      requestBody: {
        files: updatedFiles,
      },
    });

    return response.data;
  } catch (error) {
    console.error(`Error updating script content for ${scriptProjectId}:`, error.message);
    throw error;
  }
}

/**
 * Main function
 */
async function main() {
  try {
    console.log('Starting Apps Script copy process...\n');

    // Validate configuration
    if (!SOURCE_SPREADSHEET_ID) {
      console.error('ERROR: SOURCE_SPREADSHEET_ID is not set in .env file');
      process.exit(1);
    }

    // Authorize
    console.log('Authenticating with Google APIs...');
    const auth = await authorize();
    console.log('Authentication successful!\n');

    // Read spreadsheet IDs from the source sheet
    const targetSpreadsheets = await readSheetIds(auth);
    
    if (targetSpreadsheets.length === 0) {
      console.log('No spreadsheet IDs found. Exiting.');
      return;
    }

    console.log('\nTarget spreadsheets:');
    targetSpreadsheets.forEach((item, index) => {
      const scriptInfo = item.scriptId ? `(Script ID: ${item.scriptId})` : '(Script ID will be auto-discovered)';
      console.log(`  ${index + 1}. ${item.spreadsheetId} ${scriptInfo}`);
    });
    console.log('');

    // Get the source script project
    console.log('Getting source script project...');
    let sourceScriptProjectId = SOURCE_SCRIPT_ID;
    
    if (!sourceScriptProjectId) {
      console.log('No SOURCE_SCRIPT_ID provided, attempting to find it automatically...');
      sourceScriptProjectId = await getScriptProjectId(auth, SOURCE_SPREADSHEET_ID);
    } else {
      console.log('Using manually provided SOURCE_SCRIPT_ID');
    }
    
    if (!sourceScriptProjectId) {
      console.error('\nERROR: Could not find Apps Script project for source spreadsheet.');
      console.error('\nTo fix this, add the Script ID to your .env file:');
      console.error('1. Open your spreadsheet');
      console.error('2. Go to Extensions > Apps Script');
      console.error('3. Click the gear icon (Project Settings)');
      console.error('4. Copy the "Script ID"');
      console.error('5. Add it to .env file as: SOURCE_SCRIPT_ID=your_script_id_here\n');
      process.exit(1);
    }

    console.log(`Source script project ID: ${sourceScriptProjectId}\n`);

    // Read the source script content
    console.log('Reading source script content...');
    const sourceContent = await getScriptContent(auth, sourceScriptProjectId);
    console.log(`Found ${sourceContent.files.length} script file(s) in source project\n`);

    // Copy script to each target spreadsheet
    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < targetSpreadsheets.length; i++) {
      const target = targetSpreadsheets[i];
      console.log(`[${i + 1}/${targetSpreadsheets.length}] Processing spreadsheet: ${target.spreadsheetId}`);

      try {
        // Get target script project ID (use provided ID or auto-discover)
        let targetScriptProjectId = target.scriptId;
        
        if (!targetScriptProjectId) {
          console.log(`  No Script ID provided, attempting auto-discovery...`);
          targetScriptProjectId = await getScriptProjectId(auth, target.spreadsheetId);
        }
        
        if (!targetScriptProjectId) {
          console.log(`  ⚠ No Apps Script project found. Add the Script ID to column ${SCRIPT_ID_COLUMN} or create a script project. Skipping.`);
          failCount++;
          continue;
        }

        console.log(`  Script project ID: ${targetScriptProjectId}`);

        // Update the script content
        await updateScriptContent(auth, targetScriptProjectId, sourceContent);
        console.log(`  ✓ Successfully copied script content\n`);
        successCount++;

      } catch (error) {
        console.error(`  ✗ Failed: ${error.message}\n`);
        failCount++;
      }
    }

    // Summary
    console.log('\n' + '='.repeat(50));
    console.log('SUMMARY');
    console.log('='.repeat(50));
    console.log(`Total spreadsheets: ${targetSpreadsheets.length}`);
    console.log(`Successful: ${successCount}`);
    console.log(`Failed: ${failCount}`);
    console.log('='.repeat(50));

  } catch (error) {
    console.error('Fatal error:', error.message);
    process.exit(1);
  }
}

// Run the script
main();

