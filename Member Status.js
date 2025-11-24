/**
 * @OnlyCurrentDoc
 *
 * This script updates the membership status and personal details on 'Sunday Service'
 * and 'Event Attendance' sheets by cross-referencing with a central 'Directory' spreadsheet.
 */

/**
 * Main function to be run manually.
 * It reads the Directory spreadsheet, builds a map of members,
 * and then processes both 'Sunday Service' and 'Event Attendance' sheets.
 */
function processMemberStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    Logger.log("Error: 'Config' sheet not found. Please create one with the Directory link in cell B2.");
    return;
  }
  
  // Get the URL or ID of the Directory Spreadsheet from Config!B2
  const directoryUrl = configSheet.getRange('B2').getValue();
  if (!directoryUrl) {
    Logger.log("Error: Cell B2 in 'Config' sheet is empty. Please provide the Directory Spreadsheet URL or ID.");
    return;
  }

  // Build the fast lookup map from the Directory
  const directoryMap = buildDirectoryMap(directoryUrl);
  if (!directoryMap) {
    // Error is already logged in the buildDirectoryMap function
    Logger.log('Failed to build directory map. Aborting.');
    return;
  }

  Logger.log(`Directory map built successfully with ${Object.keys(directoryMap).length} entries.`);

  // Process the 'Sunday Service' sheet, starting from row 4
  processSheet(ss, 'Sunday Service', 4, directoryMap);
  
  // Process the 'Event Attendance' sheet, starting from row 5
  processSheet(ss, 'Event Attendance', 5, directoryMap);
  
  Logger.log('Processing complete for all sheets.');
  SpreadsheetApp.flush(); // Ensure all changes are saved
}

/**
 * Creates a lookup map from the Directory spreadsheet.
 * @param {string} directoryUrl The URL or ID of the Directory spreadsheet.
 * @return {Object|null} A map where keys are normalized names and
 * values are objects with {gender, lineage, age}, or null on failure.
 */
function buildDirectoryMap(directoryUrl) {
  let directorySpreadsheet;
  
  // Try opening by URL, then by ID, for flexibility
  try {
    directorySpreadsheet = SpreadsheetApp.openByUrl(directoryUrl);
  } catch (e) {
    try {
      directorySpreadsheet = SpreadsheetApp.openById(directoryUrl);
    } catch (e2) {
      Logger.log('Error: Could not open Directory spreadsheet. Invalid URL/ID in Config B2: ' + directoryUrl);
      Logger.log('Details: ' + e2);
      return null;
    }
  }

  const dirSheet = directorySpreadsheet.getSheetByName('Directory');
  if (!dirSheet) {
    Logger.log("Error: 'Directory' sheet not found in the linked spreadsheet.");
    return null;
  }

  const directoryMap = {};
  const lastRow = dirSheet.getLastRow();
  
  // Check if there is any data to read
  if (lastRow < 2) {
     Logger.log('Directory sheet is empty (no data found after row 1).');
     return directoryMap; // Return an empty map
  }
  
  // Get all data from Col C (Last Name) to Col H (Age). This is 6 columns.
  const range = dirSheet.getRange(2, 3, lastRow - 1, 6); // Starts at C2, goes to H[lastRow]
  const values = range.getValues();

  for (const row of values) {
    const lastName = row[0];  // Col C (index 0)
    const firstName = row[1]; // Col D (index 1)
    const gender = row[2];    // Col E (index 2)
    const lineage = row[3];   // Col F (index 3)
    const age = row[5];       // Col H (index 5)

    // Only add to map if we have a valid first and last name
    if (lastName && firstName) {
      const normLast = normalizeString(lastName);
      const firstNameStr = String(firstName);
      const normFirst = normalizeString(firstNameStr.split(' ')[0]); 
      
      const key = normLast + '_' + normFirst;
      
      // Avoid adding blank keys or partial keys
      if (normLast && normFirst && key !== '_') {
        directoryMap[key] = {
          gender: gender,
          lineage: lineage,
          age: age
        };
      }
    }
  }
  return directoryMap;
}

/**
 * Processes a single sheet ('Sunday Service' or 'Event Attendance') to update member status.
 * @param {Spreadsheet} ss The active spreadsheet object.
 * @param {string} sheetName The name of the sheet to process.
 * @param {number} startRow The first row containing data.
 * @param {Object} directoryMap The lookup map of directory members.
 */
function processSheet(ss, sheetName, startRow, directoryMap) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Warning: Sheet '${sheetName}' not found. Skipping.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    Logger.log(`Sheet '${sheetName}' has no data to process starting from row ${startRow}.`);
    return;
  }

  // Define the range to read and write: Col C (Last) to Col H (Type)
  const numRows = lastRow - startRow + 1;
  const numCols = 6; // C, D, E, F, G, H
  const range = sheet.getRange(startRow, 3, numRows, numCols); 
  const values = range.getValues();

  let membersFound = 0;
  let guestsFound = 0;

  // Loop through all rows in memory
  for (let i = 0; i < values.length; i++) {
    const lastName = values[i][0];  // Col C (index 0)
    const firstName = values[i][1]; // Col D (index 1)
    let key = '';
    let match = null; // Store the match result

    // Skip if name cells are blank
    if (!lastName && !firstName) {
      values[i][5] = 'Guest'; // Set Col H to Guest
      guestsFound++;
      continue;
    }
    
    // Build the key from the separate Last and First name columns
    const normLast = normalizeString(lastName);
    const firstNameStr = String(firstName);
    const normFirst = normalizeString(firstNameStr.split(' ')[0]); 

    if (normLast && normFirst && key !== '_') {
        key = normLast + '_' + normFirst;
        match = directoryMap[key]; // Look up the key
    }
    
    // Log the first key on the 'Sunday Service' sheet for debugging
    if (sheetName === 'Sunday Service' && i === 0) {
      Logger.log(`[${sheetName}] First generated key: '${key}' (from name: '${lastName}, ${firstName}')`);
    }
    
    // --- This section now works for the new structure ---
    if (match) {
      // Found: Update all fields
      values[i][2] = match.gender;  // Col E
      values[i][3] = match.lineage; // Col F
      values[i][4] = match.age;     // Col G
      values[i][5] = 'Member';      // Col H
      membersFound++;
    } else {
      // Not Found: Mark as Guest, existing E, F, G data is preserved
      values[i][5] = 'Guest';       // Col H
      guestsFound++;
    }
  }

  // Write all the updated values back to the sheet in one operation
  range.setValues(values);

  // --- Alignment updates (exactly as requested) ---
  // Vertically center all columns C:H
  sheet.getRange(startRow, 3, numRows, 6).setVerticalAlignment('middle');
  // Horizontally align C:D to left
  sheet.getRange(startRow, 3, numRows, 2).setHorizontalAlignment('left');
  // Horizontally align E:H to center
  sheet.getRange(startRow, 5, numRows, 4).setHorizontalAlignment('center');

  Logger.log(`Processed ${numRows} rows for '${sheetName}'. Found: ${membersFound} Members, ${guestsFound} Guests.`);
}

/**
 * Normalizes a string for matching.
 * Converts to lowercase, trims whitespace, and removes non-alphabetic characters.
 * @param {string} str The string to normalize.
 * @return {string} The normalized string.
 */
function normalizeString(str) {
  if (!str || typeof str !== 'string') {
    return '';
  }
  return str.toLowerCase().trim().replace(/[^a-z]/g, '');
}
