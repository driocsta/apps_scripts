const express = require('express');
const { google } = require('googleapis');
const fs = require('fs').promises;
const path = require('path');
const readline = require('readline');
require('dotenv').config();

// Initialize Express app
const app = express();
app.use(express.json());

// Initialize Google APIs
const sheets = google.sheets('v4');
const script = google.script('v1');

// Configuration from .env
const SHEET_NAME = process.env.SHEET_NAME || 'TVDE Users';
const SHEET_ID_COLUMN = process.env.SHEET_ID_COLUMN || 'D';
const SCRIPT_ID_COLUMN = process.env.SCRIPT_ID_COLUMN || 'E';
const BUTTON_IMAGE_ID_COLUMN = process.env.BUTTON_IMAGE_ID_COLUMN || 'F';
const BUTTON_COORDINATES_COLUMN = process.env.BUTTON_COORDINATES_COLUMN || 'G';
const START_ROW = parseInt(process.env.START_ROW) || 2;
const SOURCE_SCRIPT_ID = process.env.SOURCE_SCRIPT_ID || 'your_source_script_id';

/**
 * Authorize using OAuth2
 */
async function authorize() {
  const TOKEN_PATH = path.join(__dirname, 'token.json');
  const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
  
  let credentials;
  try {
    const content = await fs.readFile(CREDENTIALS_PATH);
    credentials = JSON.parse(content);
  } catch (err) {
    throw new Error('Error loading credentials.json: ' + err.message);
  }

  const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  try {
    const token = await fs.readFile(TOKEN_PATH);
    oAuth2Client.setCredentials(JSON.parse(token));
    return oAuth2Client;
  } catch (err) {
    console.log('Token not found or invalid. Generating new token...');
    return await getNewToken(oAuth2Client, TOKEN_PATH);
  }
}

/**
 * Get new OAuth2 token
 */
async function getNewToken(oAuth2Client, tokenPath) {
  const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/script.scriptapp',
    'https://www.googleapis.com/auth/drive'
  ];

  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this URL:', authUrl);

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve, reject) => {
    rl.question('Enter the code from that page here: ', (code) => {
      rl.close();
      oAuth2Client.getToken(code, async (err, token) => {
        if (err) {
          reject(new Error('Error retrieving access token: ' + err.message));
          return;
        }
        oAuth2Client.setCredentials(token);
        try {
          await fs.writeFile(tokenPath, JSON.stringify(token));
          console.log('Token stored to', tokenPath);
          resolve(oAuth2Client);
        } catch (writeErr) {
          reject(new Error('Error saving token: ' + writeErr.message));
        }
      });
    });
  });
}

/**
 * Extract spreadsheet ID from various formats
 */
function extractSpreadsheetId(input) {
  if (!input) return null;
  
  if (!input.includes('/') && !input.includes('http')) {
    return input.trim();
  }
  
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
 * Read spreadsheet IDs, script IDs, button image IDs and coordinates from master sheet
 */
async function readScriptIds(auth, sourceSpreadsheetId) {
  try {
    console.log(`Reading data from ${sourceSpreadsheetId}, sheet ${SHEET_NAME}`);
    
    const sheetMetadata = await sheets.spreadsheets.get({
      auth,
      spreadsheetId: sourceSpreadsheetId,
      fields: 'sheets.properties.title'
    });
    
    const sheetNames = sheetMetadata.data.sheets.map(s => s.properties.title);
    if (!sheetNames.includes(SHEET_NAME)) {
      throw new Error(`Sheet "${SHEET_NAME}" not found. Available: ${sheetNames.join(', ')}`);
    }

    const range = `${SHEET_NAME}!${SHEET_ID_COLUMN}${START_ROW}:${BUTTON_COORDINATES_COLUMN}`;
    const response = await sheets.spreadsheets.values.get({
      auth,
      spreadsheetId: sourceSpreadsheetId,
      range,
    });

    const rows = response.data.values || [];
    const scriptIdMap = new Map();
    
    rows.forEach((row, index) => {
      const spreadsheetId = extractSpreadsheetId(row[0]);
      const scriptId = row[1] ? row[1].trim() : null;
      const buttonImageIds = row[2] ? row[2].split(',').map(id => id.trim()).filter(id => id) : [];
      const coordinates = row[3] ? row[3].split(',').map(coord => coord.trim()).filter(coord => coord) : [];
      
      // Parse coordinates into pairs [col, row]
      const coordinatePairs = [];
      for (let i = 0; i < coordinates.length; i += 2) {
        if (i + 1 < coordinates.length) {
          const col = parseInt(coordinates[i]);
          const rowNum = parseInt(coordinates[i + 1]);
          if (!isNaN(col) && !isNaN(rowNum)) {
            coordinatePairs.push([col, rowNum]);
          }
        }
      }

      if (spreadsheetId && scriptId) {
        scriptIdMap.set(spreadsheetId, { 
          scriptId,
          buttonImageIds,
          coordinates: coordinatePairs
        });
        console.log(`Row ${index + START_ROW}: Spreadsheet ID ${spreadsheetId}, Script ID ${scriptId}, Buttons: ${buttonImageIds.length}, Coordinates: ${coordinatePairs.length}`);
      } else {
        console.warn(`Invalid row ${index + START_ROW}: Spreadsheet ID ${spreadsheetId || 'missing'}, Script ID ${scriptId || 'missing'}`);
      }
    });

    return scriptIdMap;
  } catch (error) {
    throw new Error(`Error reading source spreadsheet: ${error.message}`);
  }
}

/**
 * Copy Code.gs to target script project (overwrites existing Code.gs)
 */
async function copyFunction(auth, sourceScriptId, targetScriptId) {
  try {
    console.log(`Copying Code.gs from ${sourceScriptId} to ${targetScriptId}`);

    // Get source script content
    const sourceContent = await script.projects.getContent({ auth, scriptId: sourceScriptId });
    const sourceFiles = sourceContent.data.files || [];
    const sourceFunctionFile = sourceFiles.find(f => f.name === 'Code' && f.type === 'SERVER_JS');

    if (!sourceFunctionFile) {
      throw new Error('No Code.gs file found in source script project');
    }

    // Get target script content
    let targetContent = { data: { files: [] } };
    try {
      targetContent = await script.projects.getContent({ auth, scriptId: targetScriptId });
    } catch (error) {
      console.log(`No existing content for ${targetScriptId}. Initializing empty project.`);
    }

    let targetFiles = targetContent.data.files || [];
    
    // Remove existing Code.gs if it exists
    targetFiles = targetFiles.filter(f => f.name !== 'Code');
    
    // Add Code.gs (this will overwrite the existing Code.gs)
    targetFiles.push({
      name: 'Code',
      type: 'SERVER_JS',
      source: sourceFunctionFile.source
    });

    // Update target script project
    await script.projects.updateContent({
      auth,
      scriptId: targetScriptId,
      requestBody: { files: targetFiles },
    });

    console.log(`Successfully overwrote Code.gs in ${targetScriptId}`);
    return { scriptId: targetScriptId };
  } catch (error) {
    throw new Error(`Error copying function to ${targetScriptId}: ${error.message}`);
  }
}

/**
 * Validate script ID format and test access
 */
async function validateAndTestScriptAccess(auth, scriptId) {
  // Valid script IDs can be various lengths, including long ones
  if (!scriptId || scriptId.length < 20) {
    return { valid: false, error: 'Script ID too short' };
  }
  
  const validPattern = /^[A-Za-z0-9_-]+$/;
  if (!validPattern.test(scriptId)) {
    return { valid: false, error: 'Script ID contains invalid characters' };
  }
  
  // Try to access the script to verify it exists and we have permission
  try {
    console.log(`Testing access to script ${scriptId}...`);
    const testContent = await script.projects.get({
      auth,
      scriptId: scriptId
    });
    console.log(`✓ Script found: ${testContent.data.title || 'Untitled'}`);
    return { valid: true, title: testContent.data.title };
  } catch (error) {
    console.error(`✗ Cannot access script ${scriptId}`);
    console.error(`  Error: ${error.message}`);
    return { 
      valid: false, 
      error: error.message,
      code: error.code,
      details: error.errors 
    };
  }
}

/**
 * Copy buttons using image IDs and coordinates from master sheet
 */
async function copyButtonsFromSheet(auth, masterSpreadsheetId, sourceSpreadsheetId, sourceSheetTab, targetSpreadsheetId, targetSheetTab, buttonScript) {
  try {
    console.log(`Copying buttons to ${targetSpreadsheetId}/${targetSheetTab}`);

    // Get button data (image IDs and coordinates) from master sheet
    const scriptIdMap = await readScriptIds(auth, masterSpreadsheetId);
    const targetData = scriptIdMap.get(targetSpreadsheetId);
    
    if (!targetData || !targetData.scriptId) {
      throw new Error(`No script ID found for target spreadsheet ${targetSpreadsheetId}`);
    }
    
    // Validate and test script access
    const targetScriptId = targetData.scriptId;
    console.log(`Validating script ID: ${targetScriptId} (length: ${targetScriptId.length})`);
    
    const validation = await validateAndTestScriptAccess(auth, targetScriptId);
    if (!validation.valid) {
      console.error(`❌ Script validation failed!`);
      console.error(`   Script ID: ${targetScriptId}`);
      console.error(`   Error: ${validation.error}`);
      console.error(`   Code: ${validation.code || 'N/A'}`);
      
      if (validation.code === 404 || validation.error.includes('not found')) {
        console.error(`\n💡 The script project doesn't exist or isn't accessible. Solutions:`);
        console.error(`   1. RECOMMENDED: Create a container-bound script:`);
        console.error(`      • Open: https://docs.google.com/spreadsheets/d/${targetSpreadsheetId}/edit`);
        console.error(`      • Go to: Extensions → Apps Script`);
        console.error(`      • Copy the script ID from the URL`);
        console.error(`      • Update column E in master sheet with this new ID`);
        console.error(`   2. Check you're using the same Google account that owns the script`);
        console.error(`   3. Verify the script still exists at:`);
        console.error(`      https://script.google.com/home/projects/${targetScriptId}/edit`);
      } else if (validation.code === 403 || validation.error.includes('PERMISSION')) {
        console.error(`\n💡 Permission denied. Solutions:`);
        console.error(`   1. Make sure you're logged into the correct Google account`);
        console.error(`   2. Delete token.json and re-authenticate`);
        console.error(`   3. Grant editor access to the script project`);
      }
      
      throw new Error(`Cannot access script ${targetScriptId}: ${validation.error}`);
    }
    
    console.log(`✓ Script validated: "${validation.title}"`);
    
    if (!targetData.buttonImageIds || targetData.buttonImageIds.length === 0) {
      throw new Error(`No button image IDs found for ${targetSpreadsheetId}`);
    }
    
    if (!targetData.coordinates || targetData.coordinates.length === 0) {
      throw new Error(`No coordinates found for ${targetSpreadsheetId}`);
    }
    
    if (targetData.buttonImageIds.length !== targetData.coordinates.length) {
      throw new Error(`Mismatch: ${targetData.buttonImageIds.length} button IDs but ${targetData.coordinates.length} coordinate pairs`);
    }
    
    const buttonImageIds = targetData.buttonImageIds;
    const coordinates = targetData.coordinates;
    
    console.log(`Found ${buttonImageIds.length} button(s) to copy`);

    // Create button data as JSON string for the script
    const buttonData = JSON.stringify(buttonImageIds.map((id, i) => ({
      imageId: id,
      col: coordinates[i][0],
      row: coordinates[i][1]
    })));

    // Create a script to insert buttons at specified positions
    const tempScriptContent = `
      function copyButtonsWithFunctions() {
        Logger.clear();
        Logger.log('=== Starting button insertion process ===');
        
        try {
          Logger.log('Target Spreadsheet ID: ${targetSpreadsheetId}');
          Logger.log('Target Sheet Tab: ${targetSheetTab}');
          
          Logger.log('Opening target spreadsheet...');
          var targetSpreadsheet = SpreadsheetApp.openById('${targetSpreadsheetId}');
          Logger.log('Target spreadsheet opened: ' + targetSpreadsheet.getName());
          
          Logger.log('Getting target sheet...');
          var targetSheet = targetSpreadsheet.getSheetByName('${targetSheetTab}');
          
          if (!targetSheet) {
            Logger.log('ERROR: Target sheet "${targetSheetTab}" not found!');
            Logger.log('Available sheets: ' + targetSpreadsheet.getSheets().map(function(s) { return s.getName(); }).join(', '));
            throw new Error('Target sheet "${targetSheetTab}" not found');
          }
          Logger.log('Target sheet found: ' + targetSheet.getName());
          
          // Button data from master sheet
          var buttons = ${buttonData};
          Logger.log('Found ' + buttons.length + ' button(s) to insert');
          
          var copiedCount = 0;
          var buttonDetails = [];
          
          for (var i = 0; i < buttons.length; i++) {
            Logger.log('\\n--- Processing button ' + (i + 1) + ' of ' + buttons.length + ' ---');
            var button = buttons[i];
            
            try {
              var imageId = button.imageId;
              var col = button.col;
              var row = button.row;
              
              Logger.log('Image ID: ' + imageId);
              Logger.log('Position: Column ' + col + ', Row ' + row);
              
              // Create image URL from Drive ID
              var imageUrl = 'https://drive.google.com/uc?id=' + imageId;
              Logger.log('Image URL: ' + imageUrl);
              
              // Insert image into target sheet
              Logger.log('Inserting image into target sheet...');
              var newImage = targetSheet.insertImage(imageUrl, col, row);
              Logger.log('Image inserted successfully!');
              
              ${buttonScript ? `
              // Automatically assign script function to button
              newImage.assignScript('${buttonScript}');
              Logger.log('Assigned script function: ${buttonScript}');
              ` : '// No script function assigned - manual assignment required'}
              
              copiedCount++;
              buttonDetails.push({
                position: col + ',' + row,
                imageId: imageId
              });
              
              Logger.log('Button ' + (i + 1) + ' inserted successfully!');
            } catch (imageError) {
              Logger.log('ERROR inserting button ' + (i + 1) + ': ' + imageError.toString());
              buttonDetails.push({
                position: col + ',' + row,
                imageId: imageId,
                error: imageError.toString()
              });
            }
          }
          
          Logger.log('\\n=== Insertion complete ===');
          Logger.log('Total buttons inserted: ' + copiedCount + ' out of ' + buttons.length);
          
          var resultMessage = copiedCount + ' button(s) inserted successfully. ';
          ${buttonScript ? `
          resultMessage += 'Script function "${buttonScript}" has been automatically assigned to all buttons.';
          ` : `
          resultMessage += 'IMPORTANT: You need to manually assign script functions to each button. ';
          resultMessage += 'Right-click each button > Assign script > Enter function name (e.g., "calcularPagamentos").';
          `}
          
          Logger.log(resultMessage);
          
          return { 
            success: true, 
            copiedCount: copiedCount,
            buttons: buttonDetails,
            message: resultMessage
          };
        } catch (e) {
          Logger.log('\\n=== ERROR ===');
          Logger.log('Error: ' + e.toString());
          Logger.log('Stack: ' + e.stack);
          return { 
            success: false, 
            error: e.toString(),
            stack: e.stack
          };
        }
      }
    `;

    // Get existing script content for target
    let existingContent = { data: { files: [] } };
    try {
      existingContent = await script.projects.getContent({ auth, scriptId: targetScriptId });
    } catch (error) {
      console.log(`No existing content for ${targetScriptId}. Initializing empty project.`);
    }

    let updatedFiles = existingContent.data.files || [];
    
    // Add temporary script
    updatedFiles = updatedFiles.filter(f => f.name !== 'tempCopyButtons');
    updatedFiles.push({
      name: 'tempCopyButtons',
      type: 'SERVER_JS',
      source: tempScriptContent
    });

    // Update script project
    await script.projects.updateContent({
      auth,
      scriptId: targetScriptId,
      requestBody: { files: updatedFiles },
    });

    console.log(`Created tempCopyButtons in ${targetScriptId}`);

    // Execute the script automatically with enhanced logging
    let executionResult;
    try {
      console.log(`\n=== EXECUTING tempCopyButtons ===`);
      console.log(`Target Script ID: ${targetScriptId}`);
      console.log(`Function: copyButtonsWithFunctions`);
      console.log(`Starting execution...`);
      
      const runResponse = await script.scripts.run({
        auth,
        scriptId: targetScriptId,
        resource: { 
          function: 'copyButtonsWithFunctions', 
          parameters: [],
          devMode: false
        }
      });
      
      // Check for execution errors
      if (runResponse.data.error) {
        const error = runResponse.data.error;
        console.error(`\n❌ Execution Error:`);
        console.error(`  Message: ${error.message}`);
        console.error(`  Details: ${JSON.stringify(error.details, null, 2)}`);
        throw new Error(`Execution failed: ${error.message}`);
      }
      
      // Extract result
      executionResult = runResponse.data.response?.result || { success: true };
      
      console.log(`\n✅ Execution completed successfully!`);
      console.log(`  Buttons inserted: ${executionResult.copiedCount || 0}`);
      console.log(`  Message: ${executionResult.message || 'No message'}`);
      
      if (executionResult.buttons) {
        console.log(`  Button details:`);
        executionResult.buttons.forEach((btn, idx) => {
          console.log(`    ${idx + 1}. Position: ${btn.position}, Image: ${btn.imageId}${btn.error ? ` (ERROR: ${btn.error})` : ''}`);
        });
      }
      
      console.log(`\n📝 Keeping tempCopyButtons in ${targetScriptId} for reference`);
      console.log(`   You can view execution logs at: https://script.google.com/d/${targetScriptId}/edit`);
      
    } catch (execError) {
      console.error(`\n❌ Failed to execute copyButtonsWithFunctions:`);
      console.error(`  Error: ${execError.message}`);
      
      // Check if it's an auth/permission error
      if (execError.message.includes('PERMISSION_DENIED') || execError.message.includes('403')) {
        console.error(`\n💡 This might be a permissions issue. Make sure:`);
        console.error(`  1. Apps Script API is enabled: https://console.cloud.google.com/apis/library/script.googleapis.com`);
        console.error(`  2. OAuth scope includes: https://www.googleapis.com/auth/script.scriptapp`);
        console.error(`  3. Delete token.json and re-authenticate if you added new scopes`);
      }
      
      throw execError;
    }
    
    // NOT deleting the temporary script - keeping it for reference
    console.log(`\n✓ Button copy operation completed`);

    return {
      success: true,
      scriptId: targetScriptId,
      functionName: 'copyButtonsWithFunctions',
      copiedCount: executionResult?.copiedCount || 0,
      buttons: executionResult?.buttons || [],
      message: executionResult?.message || 'Buttons copied successfully'
    };
  } catch (error) {
    throw new Error(`Error copying buttons to ${targetSpreadsheetId}: ${error.message}`);
  }
}

/**
 * Helper function to get script ID for a spreadsheet by reading the master sheet
 */
async function getScriptIdForSpreadsheet(auth, masterSpreadsheetId, targetSpreadsheetId) {
  try {
    const scriptIdMap = await readScriptIds(auth, masterSpreadsheetId);
    const targetData = scriptIdMap.get(targetSpreadsheetId);
    return targetData ? targetData.scriptId : null;
  } catch (error) {
    console.error(`Error getting script ID for ${targetSpreadsheetId}: ${error.message}`);
    return null;
  }
}

/**
 * GET endpoint
 */
app.get('/', (req, res) => {
  res.status(200).json({
    message: 'API is running. Use POST to / to copy functions and/or buttons.',
    note: 'Button image IDs and coordinates are read from master sheet columns F and G',
    features: {
      autoExecution: 'Buttons are automatically inserted after creation',
      enhancedLogging: 'Detailed execution logs with error tracking'
    },
    expectedPayload: {
      copyFunctions: {
        enable: "true/false",
        sourceSheet: "master_spreadsheet_id (with script ID mappings)",
        targetSheets: ["spreadsheet_id_or_url_1", "spreadsheet_id_or_url_2"]
      },
      copyButtons: {
        enable: "true/false",
        sourceSheet: "master_spreadsheet_id (same as copyFunctions)",
        targetSheets: ["spreadsheet_id_or_url_1", "spreadsheet_id_or_url_2"],
        targetSheetTab: "target_sheet_tab_name",
        buttonScript: "function_name (optional, e.g., 'calcularPagamentos')"
      }
    }
  });
});

/**
 * POST endpoint
 */
app.post('/', async (req, res) => {
  const payload = req.body;

  if (!payload.copyFunctions && !payload.copyButtons) {
    return res.status(400).json({
      error: 'Invalid payload. Expected at least one of: copyFunctions, copyButtons'
    });
  }

  // Validate copyFunctions
  let functionsEnable = 'false';
  let functionsSourceSheet = null;
  let functionsTargetSheets = [];
  if (payload.copyFunctions) {
    if (typeof payload.copyFunctions.enable !== 'string' ||
        !payload.copyFunctions.sourceSheet ||
        !Array.isArray(payload.copyFunctions.targetSheets)) {
      return res.status(400).json({
        error: 'Invalid copyFunctions payload. Expected: { enable: "true/false", sourceSheet: string, targetSheets: string[] }'
      });
    }
    functionsEnable = payload.copyFunctions.enable;
    functionsSourceSheet = payload.copyFunctions.sourceSheet;
    functionsTargetSheets = payload.copyFunctions.targetSheets;
    if (functionsEnable === 'true' && (!functionsSourceSheet || functionsTargetSheets.length === 0)) {
      return res.status(400).json({ error: 'sourceSheet and targetSheets (non-empty) required for copyFunctions' });
    }
  }

  // Validate copyButtons
  let buttonsEnable = 'false';
  let buttonsSourceSheet = null;
  let buttonsTargetSheets = [];
  let targetSheetTab = null;
  let buttonScript = null;
  if (payload.copyButtons) {
    if (typeof payload.copyButtons.enable !== 'string' ||
        !payload.copyButtons.sourceSheet ||
        !Array.isArray(payload.copyButtons.targetSheets) ||
        !payload.copyButtons.targetSheetTab) {
      return res.status(400).json({
        error: 'Invalid copyButtons payload. Expected: { enable: "true/false", sourceSheet: string, targetSheets: string[], targetSheetTab: string }'
      });
    }
    buttonsEnable = payload.copyButtons.enable;
    buttonsSourceSheet = payload.copyButtons.sourceSheet;
    buttonsTargetSheets = payload.copyButtons.targetSheets;
    targetSheetTab = payload.copyButtons.targetSheetTab;
    buttonScript = payload.copyButtons.buttonScript || null;
    if (buttonsEnable === 'true' && (!buttonsSourceSheet || buttonsTargetSheets.length === 0 || !targetSheetTab)) {
      return res.status(400).json({ error: 'All fields required for copyButtons when enabled' });
    }
  }

  let auth;
  try {
    auth = await authorize();
  } catch (error) {
    return res.status(401).json({ error: 'Authentication failed: ' + error.message });
  }

  const response = {
    copyFunctions: { total: 0, successful: 0, failed: 0, details: [] },
    copyButtons: { total: 0, successful: 0, failed: 0, details: [] }
  };

  // Process copyFunctions
  if (functionsEnable === 'true') {
    const sourceSpreadsheetId = extractSpreadsheetId(functionsSourceSheet);
    response.copyFunctions.total = functionsTargetSheets.length;

    let scriptIdMap;
    try {
      scriptIdMap = await readScriptIds(auth, sourceSpreadsheetId);
    } catch (error) {
      response.copyFunctions.failed = functionsTargetSheets.length;
      response.copyFunctions.details.push({
        error: `Failed to read script IDs: ${error.message}`
      });
      return res.status(500).json(response);
    }

    for (let i = 0; i < functionsTargetSheets.length; i++) {
      const targetSpreadsheetId = extractSpreadsheetId(functionsTargetSheets[i]);
      try {
        const targetData = scriptIdMap.get(targetSpreadsheetId);
        if (!targetData || !targetData.scriptId) {
          throw new Error(`No script ID found for ${targetSpreadsheetId}`);
        }

        const result = await copyFunction(auth, SOURCE_SCRIPT_ID, targetData.scriptId);
        response.copyFunctions.successful++;
        response.copyFunctions.details.push({
          spreadsheetId: targetSpreadsheetId,
          status: 'success',
          scriptId: result.scriptId
        });
      } catch (error) {
        response.copyFunctions.failed++;
        response.copyFunctions.details.push({
          spreadsheetId: targetSpreadsheetId,
          status: 'failed',
          error: error.message
        });
      }
    }
  } else {
    response.copyFunctions.message = 'Copy functions disabled.';
  }

  // Process copyButtons
  if (buttonsEnable === 'true') {
    const masterSpreadsheetId = extractSpreadsheetId(buttonsSourceSheet);
    response.copyButtons.total = buttonsTargetSheets.length;

    for (let i = 0; i < buttonsTargetSheets.length; i++) {
      const targetSpreadsheetId = extractSpreadsheetId(buttonsTargetSheets[i]);
      try {
        const result = await copyButtonsFromSheet(
          auth,
          masterSpreadsheetId,
          null,
          null,
          targetSpreadsheetId,
          targetSheetTab,
          buttonScript
        );
        response.copyButtons.successful++;
        response.copyButtons.details.push({
          spreadsheetId: targetSpreadsheetId,
          status: 'success',
          scriptId: result.scriptId,
          functionName: result.functionName,
          copiedCount: result.copiedCount,
          buttons: result.buttons,
          message: result.message
        });
      } catch (error) {
        response.copyButtons.failed++;
        response.copyButtons.details.push({
          spreadsheetId: targetSpreadsheetId,
          status: 'failed',
          error: error.message
        });
      }
    }
  } else {
    response.copyButtons.message = 'Copy buttons disabled.';
  }

  res.status(200).json(response);
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Enhanced features: Auto-execution with detailed logging`);
});