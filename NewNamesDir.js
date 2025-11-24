/**
 * Syncs names from external Directory into 6 tabs:
 * - Event Attendance
 * - Sunday Service
 * - Appsheet Sunserv
 * - Appsheet Event
 * - Appsheet Pastoral
 * - Pastoral Check-In
 *
 * Directory is in an external spreadsheet whose URL/ID is in Config!B2,
 * sheet name "Directory", with:
 *   - Col C: Last Name
 *   - Col D: First Name
 *   - Headers on rows 1–3
 *
 * The tabs have:
 *   - Event Attendance: Col C last name, Col D first name, rows 1–4 headers
 *   - Sunday Service:   Col C last name, Col D first name, rows 1–3 headers
 *   - Appsheet Sunserv: Col B last name, Col C first name, row 1 header
 *                      + when appending, put sequential ID in Col A
 *   - Appsheet Event:   Col B last name, Col C first name, row 1 header
 *                      + when appending, put sequential ID in Col A
 *   - Appsheet Pastoral: Col B last name, Col C first name, row 1 header
 *                      + when appending, put sequential ID in Col A
 *   - Pastoral Check-In: Col C last name, Col D first name, rows 1–3 headers
 */
function syncDirectoryNamesToAllTabs() {
  const CONFIG_SHEET_NAME = 'Config';
  const DIRECTORY_SHEET_NAME = 'Directory';
  const DIRECTORY_LAST_NAME_COL = 3; // C
  const DIRECTORY_FIRST_NAME_COL = 4; // D
  const DIRECTORY_HEADER_ROWS = 3;    // 1–3 are headers

  const SHEETS_CONFIG = [
    {
      name: 'Event Attendance',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 4
      // no ID column
    },
    {
      name: 'Sunday Service',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
      // no ID column
    },
    {
      name: 'Appsheet Sunserv',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1,
      idCol: 1          // A: assign sequential number when appending
    },
    {
      name: 'Appsheet Event',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1,
      idCol: 1          // A: assign sequential number when appending
    },
    {
      name: 'Appsheet Pastoral',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1,
      idCol: 1          // A: assign sequential number when appending
    },
    {
      name: 'Pastoral Check-In',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
      // no ID column
    }
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    throw new Error("Config sheet '" + CONFIG_SHEET_NAME + "' not found.");
  }

  // --- Get external Directory spreadsheet ID from Config!B2 ---
  const externalRef = configSheet.getRange('B2').getValue();
  if (!externalRef) {
    throw new Error("Config!B2 is empty. Please put the Directory spreadsheet URL or ID there.");
  }
  const externalId = extractSpreadsheetIdFromString_(String(externalRef));

  const externalSs = SpreadsheetApp.openById(externalId);
  const directorySheet = externalSs.getSheetByName(DIRECTORY_SHEET_NAME);
  if (!directorySheet) {
    throw new Error("Directory sheet '" + DIRECTORY_SHEET_NAME + "' not found in external spreadsheet.");
  }

  const lastDirRow = directorySheet.getLastRow();
  if (lastDirRow <= DIRECTORY_HEADER_ROWS) {
    Logger.log('No data rows found in Directory.');
    return;
  }

  const dirNumRows = lastDirRow - DIRECTORY_HEADER_ROWS;
  const dirLastNames = directorySheet
    .getRange(DIRECTORY_HEADER_ROWS + 1, DIRECTORY_LAST_NAME_COL, dirNumRows, 1)
    .getValues();
  const dirFirstNames = directorySheet
    .getRange(DIRECTORY_HEADER_ROWS + 1, DIRECTORY_FIRST_NAME_COL, dirNumRows, 1)
    .getValues();

  // Build a clean list of directory names and keys
  const directoryEntries = [];
  for (let i = 0; i < dirNumRows; i++) {
    const lastName = (dirLastNames[i][0] || '').toString().trim();
    const firstName = (dirFirstNames[i][0] || '').toString().trim();
    if (!lastName && !firstName) continue;

    const key = buildNameKey_(lastName, firstName);
    if (!key) continue;

    directoryEntries.push({
      lastName: lastName,
      firstName: firstName,
      key: key
    });
  }

  if (directoryEntries.length === 0) {
    Logger.log('Directory has no valid (Last, First) name rows.');
    return;
  }

  // --- For each sheet, append any missing names from Directory ---
  SHEETS_CONFIG.forEach(function (cfg) {
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      Logger.log("Sheet '" + cfg.name + "' not found. Skipping.");
      return;
    }

    const lastRow = sheet.getLastRow();
    const numCols = sheet.getLastColumn();
    const dataStartRow = cfg.headerRows + 1;

    const existingKeys = new Set();

    if (lastRow >= dataStartRow) {
      const numRows = lastRow - cfg.headerRows;

      const lastNameValues = sheet
        .getRange(dataStartRow, cfg.lastNameCol, numRows, 1)
        .getValues();
      const firstNameValues = sheet
        .getRange(dataStartRow, cfg.firstNameCol, numRows, 1)
        .getValues();

      for (let r = 0; r < numRows; r++) {
        const ln = (lastNameValues[r][0] || '').toString().trim();
        const fn = (firstNameValues[r][0] || '').toString().trim();
        if (!ln && !fn) continue;

        const key = buildNameKey_(ln, fn);
        if (key) existingKeys.add(key);
      }
    }

    // Determine next ID for sheets that have an ID column (Appsheet tabs)
    let nextId = null;
    if (cfg.idCol) {
      let maxId = 0;
      if (lastRow >= dataStartRow) {
        const idNumRows = lastRow - cfg.headerRows;
        const idValues = sheet
          .getRange(dataStartRow, cfg.idCol, idNumRows, 1)
          .getValues();

        for (let r = 0; r < idNumRows; r++) {
          const raw = idValues[r][0];
          if (raw === '' || raw === null) continue;
          const n = Number(raw);
          if (!isNaN(n) && n > maxId) maxId = n;
        }
      }
      nextId = maxId; // we will increment before assigning
    }

    const rowsToAppend = [];
    directoryEntries.forEach(function (entry) {
      if (!existingKeys.has(entry.key)) {
        existingKeys.add(entry.key);

        const newRow = new Array(numCols).fill('');

        if (cfg.idCol && nextId !== null) {
          nextId += 1;
          newRow[cfg.idCol - 1] = nextId;
        }

        newRow[cfg.lastNameCol - 1] = entry.lastName;
        newRow[cfg.firstNameCol - 1] = entry.firstName;

        rowsToAppend.push(newRow);
      }
    });

    if (rowsToAppend.length > 0) {
      const appendStartRow = getNextAvailableRow_(sheet, dataStartRow, cfg.lastNameCol, cfg.firstNameCol);
      sheet.getRange(appendStartRow, 1, rowsToAppend.length, numCols)
        .setValues(rowsToAppend);

      Logger.log("Sheet '" + cfg.name + "': appended " + rowsToAppend.length + " names from Directory.");
    } else {
      Logger.log("Sheet '" + cfg.name + "': no new names needed.");
    }
  });
}

/**
 * Finds the next truly empty row (based on last/first name columns),
 * starting from dataStartRow.
 */
function getNextAvailableRow_(sheet, dataStartRow, lastNameCol, firstNameCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return dataStartRow;

  const numRows = lastRow - dataStartRow + 1;
  const lastNames = sheet.getRange(dataStartRow, lastNameCol, numRows, 1).getValues();
  const firstNames = sheet.getRange(dataStartRow, firstNameCol, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    const ln = (lastNames[i][0] || '').toString().trim();
    const fn = (firstNames[i][0] || '').toString().trim();
    if (!ln && !fn) {
      return dataStartRow + i;
    }
  }
  return lastRow + 1;
}

/**
 * Builds a normalized name key from Last Name + First Name.
 * This helps avoid duplicates caused by spacing/case/punctuation differences.
 */
function buildNameKey_(lastName, firstName) {
  const ln = (lastName || '').toString().trim().toLowerCase();
  const fn = (firstName || '').toString().trim().toLowerCase();
  if (!ln && !fn) return null;

  const cleanLn = ln.replace(/[^A-Za-z\u00C0-\u024F]/g, '');
  const cleanFn = fn.replace(/[^A-Za-z\u00C0-\u024F]/g, '');
  if (!cleanLn && !cleanFn) return null;

  return cleanLn + '|' + cleanFn;
}

/**
 * Extracts a spreadsheet ID from either a raw ID or a full URL string.
 */
function extractSpreadsheetIdFromString_(input) {
  if (/^[\w-]{25,}$/.test(input)) return input;

  const match = input.match(/[-\w]{25,}/);
  if (match && match[0]) return match[0];

  throw new Error('Could not extract Spreadsheet ID from: ' + input);
}
