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
 *
 * EXTRA:
 * - Also scans Event Attendance, Sunday Service, Pastoral Check-In,
 *   Appsheet Sunserv, Appsheet Event, Appsheet Pastoral
 *   and ensures all unique names (Directory + those 6) exist in:
 *   - Event Attendance
 *   - Sunday Service
 *   - Pastoral Check-In
 *   - Appsheet Sunserv
 *   - Appsheet Event
 *   - Appsheet Pastoral
 *   (no duplicates in each tab)
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
    // We still continue later, because we may sync names from other tabs
  }

  const dirNumRows = Math.max(0, lastDirRow - DIRECTORY_HEADER_ROWS);
  let dirLastNames = [];
  let dirFirstNames = [];
  if (dirNumRows > 0) {
    dirLastNames = directorySheet
      .getRange(DIRECTORY_HEADER_ROWS + 1, DIRECTORY_LAST_NAME_COL, dirNumRows, 1)
      .getValues();
    dirFirstNames = directorySheet
      .getRange(DIRECTORY_HEADER_ROWS + 1, DIRECTORY_FIRST_NAME_COL, dirNumRows, 1)
      .getValues();
  }

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
    // Do not return; we still want to sync names from other tabs
  }

  // --- For each sheet, append any missing names from Directory (one-way) ---
  SHEETS_CONFIG.forEach(function (cfg) {
    const sheet = getSheetByNameLoose_(ss, cfg.name);
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

      Logger.log("Sheet '" + sheet.getName() + "': appended " + rowsToAppend.length + " names from Directory.");
    } else {
      Logger.log("Sheet '" + sheet.getName() + "': no new names needed from Directory.");
    }
  });

  // --- Build union of names (Directory + ALL 6 tabs) ---
  const unionEntries = buildUnionEntries_(ss, directoryEntries);

  // --- Ensure all union names are present in Event Attendance / Sunday Service / Pastoral Check-In ---
  syncUnionNamesIntoAttendanceTabs_(ss, unionEntries);

  // --- Ensure all union names are present in Appsheet tabs (no duplicates) ---
  syncUnionNamesIntoAppsheetTabs_(ss, unionEntries);

  // --- After syncing, sort the 6 tabs by Attendance Stats status + name ---
  sortSyncedTabsByAttendanceStatus();
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

/**
 * Loose sheet matcher:
 * - tries exact getSheetByName first
 * - then matches by normalized name (trim, collapse spaces, lowercase)
 */
function getSheetByNameLoose_(ss, targetName) {
  const exact = ss.getSheetByName(targetName);
  if (exact) return exact;

  const normTarget = normalizeSheetName_(targetName);
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    if (normalizeSheetName_(s.getName()) === normTarget) {
      return s;
    }
  }
  return null;
}

/**
 * Normalizes sheet names to avoid hidden-space/case issues.
 */
function normalizeSheetName_(name) {
  return String(name || '')
    .replace(/\s+/g, ' ')  // collapse any weird whitespace
    .trim()
    .toLowerCase();
}

/**
 * Returns a numeric rank for a given status string.
 * Lower number = higher priority in sorting.
 *
 * 0 = core
 * 1 = active
 * 2 = inactive
 * 3 = archived
 * 4 = not found / other
 */
function getStatusRank_(rawStatus) {
  const s = String(rawStatus || '').toLowerCase().trim();
  if (s === 'core') return 0;
  if (s === 'active') return 1;
  if (s === 'inactive') return 2;
  if (s === 'archived') return 3;
  return 4; // not found or other
}

/**
 * Sorts the following tabs by status from 'Attendance Stats' and then by name:
 * - Event Attendance
 * - Sunday Service
 * - Appsheet Sunserv
 * - Appsheet Event
 * - Appsheet Pastoral
 * - Pastoral Check-In
 *
 * 'Attendance Stats' sheet:
 *   - Col C: Last Name
 *   - Col D: First Name
 *   - Col F: Status (core, active, inactive, archived)
 *   - Rows 1–2 headers, data from row 3
 *
 * Sorting order:
 *   1) core (top, A–Z by last, first)
 *   2) active (A–Z by last, first)
 *   3) inactive (A–Z by last, first)
 *   4) archived or not found in Attendance Stats (A–Z by last, first)
 */
function sortSyncedTabsByAttendanceStatus() {
  const ATTENDANCE_STATS_SHEET_NAME = 'Attendance Stats';
  const ATT_LAST_NAME_COL = 3; // C
  const ATT_FIRST_NAME_COL = 4; // D
  const ATT_STATUS_COL = 6; // F
  const ATT_HEADER_ROWS = 2; // rows 1–2

  const SORT_SHEETS_CONFIG = [
    {
      name: 'Event Attendance',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 4
    },
    {
      name: 'Sunday Service',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
    },
    {
      name: 'Appsheet Sunserv',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1
    },
    {
      name: 'Appsheet Event',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1
    },
    {
      name: 'Appsheet Pastoral',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1
    },
    {
      name: 'Pastoral Check-In',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
    }
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(ATTENDANCE_STATS_SHEET_NAME);
  if (!attSheet) {
    throw new Error("Attendance Stats sheet '" + ATTENDANCE_STATS_SHEET_NAME + "' not found.");
  }

  const lastRowStats = attSheet.getLastRow();
  if (lastRowStats <= ATT_HEADER_ROWS) {
    Logger.log('Attendance Stats has no data rows to use for sorting.');
    return;
  }

  const statsNumRows = lastRowStats - ATT_HEADER_ROWS;

  const statsLastNames = attSheet
    .getRange(ATT_HEADER_ROWS + 1, ATT_LAST_NAME_COL, statsNumRows, 1)
    .getValues();
  const statsFirstNames = attSheet
    .getRange(ATT_HEADER_ROWS + 1, ATT_FIRST_NAME_COL, statsNumRows, 1)
    .getValues();
  const statsStatuses = attSheet
    .getRange(ATT_HEADER_ROWS + 1, ATT_STATUS_COL, statsNumRows, 1)
    .getValues();

  // Build map: key (Last|First normalized) -> statusRank
  const statusMap = {};
  for (let i = 0; i < statsNumRows; i++) {
    const ln = (statsLastNames[i][0] || '').toString().trim();
    const fn = (statsFirstNames[i][0] || '').toString().trim();
    const status = statsStatuses[i][0];

    const key = buildNameKey_(ln, fn);
    if (!key) continue;

    statusMap[key] = getStatusRank_(status);
  }

  // Sort each of the target sheets
  SORT_SHEETS_CONFIG.forEach(function (cfg) {
    const sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Sheet '" + cfg.name + "' not found. Skipping sort.");
      return;
    }

    const lastRow = sheet.getLastRow();
    const headerRows = cfg.headerRows;
    if (lastRow <= headerRows) {
      Logger.log("Sheet '" + cfg.name + "' has no data rows to sort.");
      return;
    }

    const numRows = lastRow - headerRows;
    const numCols = sheet.getLastColumn();

    const dataRange = sheet.getRange(headerRows + 1, 1, numRows, numCols);
    const dataValues = dataRange.getValues();

    const rowsWithMeta = dataValues.map(function (row, index) {
      const lnRaw = (row[cfg.lastNameCol - 1] || '').toString().trim();
      const fnRaw = (row[cfg.firstNameCol - 1] || '').toString().trim();
      const key = buildNameKey_(lnRaw, fnRaw);

      let rank = 4; // default: not found / other
      if (key && Object.prototype.hasOwnProperty.call(statusMap, key)) {
        rank = statusMap[key];
      }

      return {
        row: row,
        rank: rank,
        ln: lnRaw.toLowerCase(),
        fn: fnRaw.toLowerCase(),
        originalIndex: index
      };
    });

    rowsWithMeta.sort(function (a, b) {
      if (a.rank !== b.rank) return a.rank - b.rank;
      if (a.ln !== b.ln) return a.ln.localeCompare(b.ln);
      if (a.fn !== b.fn) return a.fn.localeCompare(b.fn);
      return a.originalIndex - b.originalIndex;
    });

    const sortedValues = rowsWithMeta.map(function (item) {
      return item.row;
    });

    dataRange.setValues(sortedValues);
    Logger.log("Sheet '" + cfg.name + "' sorted by status and name.");
  });
}

/**
 * Build a union of names from:
 * - Directory (directoryEntries)
 * - Event Attendance
 * - Sunday Service
 * - Pastoral Check-In
 *
 * NOTE (Option 1):
 * Appsheet tabs are NOT used as a source here.
 * They only receive the union (targets), so
 * names that exist only in Appsheet are not
 * pushed into the other 3 attendance tabs.
 */
function buildUnionEntries_(ss, directoryEntries) {
  const EXTRA_SOURCE_SHEETS = [
    {
      name: 'Event Attendance',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 4
    },
    {
      name: 'Sunday Service',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
    },
    {
      name: 'Pastoral Check-In',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
    }
  ];

  const unionMap = new Map();

  // Start with Directory entries
  directoryEntries.forEach(function (entry) {
    if (!unionMap.has(entry.key)) {
      unionMap.set(entry.key, {
        lastName: entry.lastName,
        firstName: entry.firstName,
        key: entry.key
      });
    }
  });

  // Add names from extra source sheets (attendance tabs only)
  EXTRA_SOURCE_SHEETS.forEach(function (cfg) {
    const sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Extra source sheet '" + cfg.name + "' not found. Skipping.");
      return;
    }

    const lastRow = sheet.getLastRow();
    const dataStartRow = cfg.headerRows + 1;
    if (lastRow < dataStartRow) {
      Logger.log("Extra source sheet '" + cfg.name + "' has no data rows.");
      return;
    }

    const numRows = lastRow - cfg.headerRows;
    const lastNames = sheet
      .getRange(dataStartRow, cfg.lastNameCol, numRows, 1)
      .getValues();
    const firstNames = sheet
      .getRange(dataStartRow, cfg.firstNameCol, numRows, 1)
      .getValues();

    for (let i = 0; i < numRows; i++) {
      const ln = (lastNames[i][0] || '').toString().trim();
      const fn = (firstNames[i][0] || '').toString().trim();
      if (!ln && !fn) continue;

      const key = buildNameKey_(ln, fn);
      if (!key) continue;

      if (!unionMap.has(key)) {
        unionMap.set(key, {
          lastName: ln,
          firstName: fn,
          key: key
        });
      }
    }
  });

  return Array.from(unionMap.values());
}


/**
 * Ensure all union names exist in:
 * - Event Attendance
 * - Sunday Service
 * - Pastoral Check-In
 * (no duplicates in each tab)
 */
function syncUnionNamesIntoAttendanceTabs_(ss, unionEntries) {
  const ATTENDANCE_TABS_CONFIG = [
    {
      name: 'Event Attendance',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 4
    },
    {
      name: 'Sunday Service',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
    },
    {
      name: 'Pastoral Check-In',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
    }
  ];

  ATTENDANCE_TABS_CONFIG.forEach(function (cfg) {
    const sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Attendance tab '" + cfg.name + "' not found. Skipping union sync.");
      return;
    }

    const lastRow = sheet.getLastRow();
    const dataStartRow = cfg.headerRows + 1;
    const numCols = sheet.getLastColumn();
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

    const rowsToAppend = [];
    unionEntries.forEach(function (entry) {
      if (!existingKeys.has(entry.key)) {
        existingKeys.add(entry.key);

        const newRow = new Array(numCols).fill('');
        newRow[cfg.lastNameCol - 1] = entry.lastName;
        newRow[cfg.firstNameCol - 1] = entry.firstName;

        rowsToAppend.push(newRow);
      }
    });

    if (rowsToAppend.length > 0) {
      const appendStartRow = getNextAvailableRow_(sheet, dataStartRow, cfg.lastNameCol, cfg.firstNameCol);
      sheet.getRange(appendStartRow, 1, rowsToAppend.length, numCols)
        .setValues(rowsToAppend);

      Logger.log("Attendance tab '" + sheet.getName() + "': appended " + rowsToAppend.length + " union names.");
    } else {
      Logger.log("Attendance tab '" + cfg.name + "': no union names needed.");
    }
  });
}

/**
 * Take the union of names and make sure ALL of them exist
 * in the 3 Appsheet tabs (no duplicates per tab):
 * - Appsheet Sunserv
 * - Appsheet Event
 * - Appsheet Pastoral
 */
function syncUnionNamesIntoAppsheetTabs_(ss, unionEntries) {
  const APPSHEET_TABS_CONFIG = [
    {
      name: 'Appsheet Sunserv',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1,
      idCol: 1          // A
    },
    {
      name: 'Appsheet Event',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1,
      idCol: 1          // A
    },
    {
      name: 'Appsheet Pastoral',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1,
      idCol: 1          // A
    }
  ];

  APPSHEET_TABS_CONFIG.forEach(function (cfg) {
    const sheet = getSheetByNameLoose_(ss, cfg.name);
    if (!sheet) {
      Logger.log("Appsheet tab '" + cfg.name + "' not found. Skipping union sync.");
      return;
    }

    const lastRow = sheet.getLastRow();
    const dataStartRow = cfg.headerRows + 1;
    const numCols = sheet.getLastColumn();
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

    // Determine next ID for Appsheet tabs
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
      nextId = maxId;
    }

    const rowsToAppend = [];
    unionEntries.forEach(function (entry) {
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

      Logger.log("Appsheet tab '" + sheet.getName() + "': appended " + rowsToAppend.length + " union names.");
    } else {
      Logger.log("Appsheet tab '" + cfg.name + "': no union names needed.");
    }
  });
}
/**
 * Debug helper:
 * - Counts how many unique names each tab has.
 * - Shows how many names in Appsheet tabs are NOT in Directory.
 *
 * Usage:
 * - Run checkNameCountsAndExtras() from the Script Editor.
 * - Open View > Logs to see the summary.
 */
function checkNameCountsAndExtras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const CONFIG_SHEET_NAME = 'Config';
  const DIRECTORY_SHEET_NAME = 'Directory';

  // Tab configuration for reading names
  const tabConfigs = [
    {
      id: 'directory',
      label: 'Directory',
      sheetName: DIRECTORY_SHEET_NAME,
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3,
      external: true    // in external spreadsheet
    },
    {
      id: 'event',
      label: 'Event Attendance',
      sheetName: 'Event Attendance',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 4
    },
    {
      id: 'sunday',
      label: 'Sunday Service',
      sheetName: 'Sunday Service',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
    },
    {
      id: 'pastoralCheck',
      label: 'Pastoral Check-In',
      sheetName: 'Pastoral Check-In',
      lastNameCol: 3,   // C
      firstNameCol: 4,  // D
      headerRows: 3
    },
    {
      id: 'appSunserv',
      label: 'Appsheet Sunserv',
      sheetName: 'Appsheet Sunserv',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1
    },
    {
      id: 'appEvent',
      label: 'Appsheet Event',
      sheetName: 'Appsheet Event',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1
    },
    {
      id: 'appPastoral',
      label: 'Appsheet Pastoral',
      sheetName: 'Appsheet Pastoral',
      lastNameCol: 2,   // B
      firstNameCol: 3,  // C
      headerRows: 1
    }
  ];

  // --- Open external Directory spreadsheet (same logic as main sync) ---
  let externalSs = null;
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (configSheet) {
    const externalRef = configSheet.getRange('B2').getValue();
    if (externalRef) {
      const externalId = extractSpreadsheetIdFromString_(String(externalRef));
      externalSs = SpreadsheetApp.openById(externalId);
    } else {
      Logger.log("checkNameCountsAndExtras: Config!B2 is empty, cannot read external Directory.");
    }
  } else {
    Logger.log("checkNameCountsAndExtras: Config sheet not found.");
  }

  // Store results per tab
  const results = {};

  tabConfigs.forEach(function (cfg) {
    let sheet = null;

    if (cfg.external) {
      if (!externalSs) {
        Logger.log("checkNameCountsAndExtras: External Directory spreadsheet not available.");
        return;
      }
      sheet = externalSs.getSheetByName(cfg.sheetName);
    } else {
      sheet = getSheetByNameLoose_(ss, cfg.sheetName);
    }

    if (!sheet) {
      Logger.log("checkNameCountsAndExtras: Sheet '" + cfg.sheetName + "' not found. Skipping.");
      return;
    }

    const lastRow = sheet.getLastRow();
    const dataStartRow = cfg.headerRows + 1;

    if (lastRow < dataStartRow) {
      results[cfg.id] = {
        label: cfg.label,
        totalRows: 0,
        nonEmptyRows: 0,
        uniqueCount: 0,
        map: new Map()
      };
      Logger.log(cfg.label + ": no data rows.");
      return;
    }

    const numRows = lastRow - cfg.headerRows;
    const lastNames = sheet
      .getRange(dataStartRow, cfg.lastNameCol, numRows, 1)
      .getValues();
    const firstNames = sheet
      .getRange(dataStartRow, cfg.firstNameCol, numRows, 1)
      .getValues();

    const nameMap = new Map();
    let nonEmptyRows = 0;

    for (let i = 0; i < numRows; i++) {
      const ln = (lastNames[i][0] || '').toString().trim();
      const fn = (firstNames[i][0] || '').toString().trim();
      if (!ln && !fn) {
        continue;
      }
      nonEmptyRows++;

      const key = buildNameKey_(ln, fn);
      if (!key) continue;

      if (!nameMap.has(key)) {
        nameMap.set(key, { lastName: ln, firstName: fn });
      }
    }

    results[cfg.id] = {
      label: cfg.label,
      totalRows: numRows,        // rows under headers (including blank)
      nonEmptyRows: nonEmptyRows, // rows that have at least one of last/first
      uniqueCount: nameMap.size, // unique normalized names
      map: nameMap
    };

    Logger.log(
      cfg.label +
      ": rows=" + numRows +
      ", nonEmptyRows=" + nonEmptyRows +
      ", uniqueNames=" + nameMap.size
    );
  });

  // --- Compare Appsheet tabs vs Directory (to explain "2x bigger") ---
  const dirRes = results['directory'];
  if (!dirRes) {
    Logger.log("checkNameCountsAndExtras: Directory results not available, cannot compare extras.");
    return;
  }

  const dirMap = dirRes.map;
  const appsheetIds = ['appSunserv', 'appEvent', 'appPastoral'];

  appsheetIds.forEach(function (id) {
    const res = results[id];
    if (!res) return;

    const extras = [];
    res.map.forEach(function (value, key) {
      if (!dirMap.has(key)) {
        extras.push(value.lastName + ', ' + value.firstName);
      }
    });

    Logger.log(res.label + ": names NOT in Directory = " + extras.length);

    // Show a preview (first 30) to help debugging without flooding the log
    if (extras.length > 0) {
      const previewCount = Math.min(30, extras.length);
      const preview = extras.slice(0, previewCount).join(" | ");
      Logger.log(
        res.label +
        " (first " + previewCount + " extras not in Directory): " +
        preview
      );
    }
  });
}
