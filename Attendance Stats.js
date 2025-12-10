/**
 * A more robust function to clean and standardize names.
 * Removes all spaces and keeps only letters/numbers/commas for a consistent key.
 */
function normalizeName(name) {
  if (!name) return '';
  return name.toString().toLowerCase().replace(/[^a-z0-9,]/g, '').trim();
}

/**
 * Extracts a purely numeric ID.
 */
function extractNumericBel(codeValue) {
  if (typeof codeValue === 'number' && Number.isInteger(codeValue) && codeValue >= 0) { return codeValue; }
  if (typeof codeValue === 'string') {
    const numStr = codeValue.trim();
    if (numStr === '' || !/^\d+$/.test(numStr)) { return null; }
    const num = parseInt(numStr, 10);
    if (!isNaN(num) && num >= 0) { return num; }
  }
  return null;
}

/**
 * Fetches all raw data from sheets.
 * Reads Directory from external sheet ID in Config!B2.
 * Reads Event Attendance and Sunday Service.
 * Reads Attendance Log for Pastoral Check-In.
 */
function getDataFromSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let externalDirectorySs = null;

  // --- External Directory ---
  try {
    const configSheet = ss.getSheetByName("Config");
    if (configSheet) {
      const directoryId = configSheet.getRange("B2").getValue();
      if (directoryId) {
        externalDirectorySs = SpreadsheetApp.openById(directoryId);
      } else {
        Logger.log('❌ "Config!B2" is empty.');
      }
    } else {
      Logger.log('❌ "Config" sheet not found.');
    }
  } catch (e) {
    Logger.log(`❌ Error opening Directory sheet: ${e}`);
  }

  const getSheetData = (sheetName, spreadsheet = ss) => {
    const sheet = spreadsheet ? spreadsheet.getSheetByName(sheetName) : null;
    if (!sheet) {
      Logger.log(`❌ Sheet "${sheetName}" not found.`);
      return [];
    }
    return sheet.getDataRange().getValues();
  };

  return {
    dData: getSheetData("Directory", externalDirectorySs),
    eData: getSheetData("Event Attendance", ss),
    sData: getSheetData("Sunday Service", ss),
    lData: getSheetData("Attendance Log", ss)
  };
}

/**
 * Assigns BEL codes and merges attendance from:
 *  - Directory
 *  - Sunday Service
 *  - Event Attendance
 *  - Attendance Log (Pastoral Check-In only)
 */
function matchOrAssignBelCodes() {
  const data = getDataFromSheets();
  if (!data) return { rawData: [], dData: [] };

  const { sData, eData, dData, lData } = data;

  const belMap = new Map();
  const allUsedCodes = new Set();

  // --- Populate from Directory ---
  const populateFromDirectory = (sourceData, startRow) => {
    if (!sourceData || sourceData.length <= startRow) return;
    sourceData.slice(startRow).forEach(row => {
      const id = extractNumericBel(row[0]);
      const last = row[2];
      const first = row[3];
      if (id !== null && last && first) {
        const key = normalizeName(`${last}, ${first}`);
        allUsedCodes.add(id);
        if (!belMap.has(key)) belMap.set(key, id);
      }
    });
  };

  // --- Populate from Attendance (Sunday Service + Event Attendance) ---
  const populateFromAttendance = (sourceData, startRow) => {
    if (!sourceData || sourceData.length <= startRow) return;
    sourceData.slice(startRow).forEach(row => {
      const id = extractNumericBel(row[0]);
      const last = row[2];
      const first = row[3];
      const key = normalizeName(`${last}, ${first}`);
      if (id !== null) {
        allUsedCodes.add(id);
        if (key && !belMap.has(key)) belMap.set(key, id);
      }
    });
  };

  populateFromDirectory(dData, 1);
  populateFromAttendance(eData, 3);
  populateFromAttendance(sData, 2);

  // --- ID generator ---
  let codeCounter = allUsedCodes.size > 0 ? Math.max(...allUsedCodes) + 1 : 1;
  const generateBEL = () => {
    while (allUsedCodes.has(codeCounter)) codeCounter++;
    allUsedCodes.add(codeCounter);
    return codeCounter++;
  };

  const results = [];

  // --- EVENT ATTENDANCE ---
  if (eData && eData.length > 3) {
    const dates = eData[1];
    const names = eData[2];
    eData.slice(3).forEach(row => {
      const last = row[2];
      const first = row[3];
      const key = normalizeName(`${last}, ${first}`);
      if (!key) return;

      let bel = belMap.get(key) || generateBEL();
      belMap.set(key, bel);

      for (let c = 8; c < row.length; c++) {
        if (row[c] === true) {
          const date = dates[c];
          const eventName = names[c];
          if (date && eventName) {
            results.push([bel, first, last, eventName, eventName, date, false]);
          }
        }
      }
    });
  }

  // --- SUNDAY SERVICE ---
  if (sData && sData.length > 2) {
    const dates = sData[1];
    sData.slice(2).forEach(row => {
      const last = row[2];
      const first = row[3];
      const key = normalizeName(`${last}, ${first}`);
      if (!key) return;

      let bel = belMap.get(key) || generateBEL();
      belMap.set(key, bel);

      for (let c = 8; c < row.length; c++) {
        if (row[c] === true) {
          const date = dates[c];
          if (date) {
            results.push([bel, first, last, "Sunday Service", "Sunday Service", date, false]);
          }
        }
      }
    });
  }

  // --- PASTORAL CHECK-IN FROM ATTENDANCE LOG (DEDUPLICATED PER DATE) ---
  if (lData && lData.length > 1) {

    // Keep map of "BEL|DATE" to prevent duplicates
    const pastoralSeen = new Set();

    lData.slice(1).forEach(row => {
      const event = row[5];
      if (!event) return;

      if (String(event).toLowerCase().trim() !== "pastoral check-in") return;

      const last = row[2];
      const first = row[3];
      const key = normalizeName(`${last}, ${first}`);
      if (!key) return;

      let bel = belMap.get(key) || generateBEL();
      belMap.set(key, bel);

      const date = row[6];
      if (!date) return;

      // Create unique key per date
      const dedupeKey = `${bel}|${new Date(date).toDateString()}`;
      if (pastoralSeen.has(dedupeKey)) return;   // skip duplicates

      pastoralSeen.add(dedupeKey);

      // Add only once
      results.push([bel, first, last, "Pastoral Check-In", "Pastoral Check-In", date, false]);
    });
  }

  return { rawData: results, dData };
}

/**
 * Calculates stats (Q1–Q4, Total, Last Event, Guest Flag).
 * FIXED: Column M now returns full event name even if it contains hyphens,
 * and Pastoral Check-In is forced to exact label.
 */
function calculateAttendanceStats() {
  const { rawData, dData } = matchOrAssignBelCodes();
  if (rawData.length === 0) return [];

  // --- Build Directory Name Lookup Set ---
  const directoryNamesSet = new Set();
  if (dData && dData.length > 1) {
    dData.slice(1).forEach(row => {
      const lastName = row[2];
      const firstName = row[3];
      if (lastName && firstName) {
        directoryNamesSet.add(normalizeName(`${lastName}, ${firstName}`));
      }
    });
  }

  const now = new Date();
  const currentYear = now.getFullYear();

  const q1_start = new Date(currentYear, 0, 1);
  const q1_end   = new Date(currentYear, 3, 0);
  const q2_start = new Date(currentYear, 3, 1);
  const q2_end   = new Date(currentYear, 6, 0);
  const q3_start = new Date(currentYear, 6, 1);
  const q3_end   = new Date(currentYear, 9, 0);
  const q4_start = new Date(currentYear, 9, 1);
  const q4_end   = new Date(currentYear, 11, 31);

  const grouped = new Map();

  rawData.forEach(row => {
    if (row.length < 7) return;
    const [bel, firstName, lastName, eventName, eventId, dateVal, isVolunteer] = row;

    const date = dateVal instanceof Date ? dateVal : new Date(String(dateVal));
    if (isNaN(date.getTime())) return;

    const isSundayService = /sunday service/i.test(eventName);
    const eventKey = isSundayService
      ? `Sunday Service-${date.toDateString()}`
      : `${eventName}-${date.toDateString()}`;

    const record = {
      firstName,
      lastName,
      date,
      eventKey,
      isVolunteer: isVolunteer === true
    };

    if (!grouped.has(bel)) grouped.set(bel, []);
    grouped.get(bel).push(record);
  });

  const summary = [];
  grouped.forEach((records, bel) => {
    const q1Events = new Set(),
          q2Events = new Set(),
          q3Events = new Set(),
          q4Events = new Set();

    records.forEach(r => {
      if (r.date >= q1_start && r.date <= q1_end) q1Events.add(r.eventKey);
      if (r.date >= q2_start && r.date <= q2_end) q2Events.add(r.eventKey);
      if (r.date >= q3_start && r.date <= q3_end) q3Events.add(r.eventKey);
      if (r.date >= q4_start && r.date <= q4_end) q4Events.add(r.eventKey);
    });

    records.sort((a, b) => b.date.getTime() - a.date.getTime());
    const mostRecentRecord = records[0];

    // FIX: get full event name even if it has hyphens (use LAST hyphen)
    const lastDashIndex = mostRecentRecord.eventKey.lastIndexOf("-");
    let lastEventName = lastDashIndex > -1
      ? mostRecentRecord.eventKey.substring(0, lastDashIndex)
      : mostRecentRecord.eventKey;

    // Force exact Pastoral Check-In label
    if (/pastoral\s*check[-\s]*in/i.test(lastEventName)) {
      lastEventName = "Pastoral Check-In";
    }

    const standardizedName = normalizeName(`${mostRecentRecord.lastName}, ${mostRecentRecord.firstName}`);
    const isGuest = !directoryNamesSet.has(standardizedName);
    const guestStatus = isGuest ? "Guest" : "";

    summary.push([
      bel,
      mostRecentRecord.firstName,
      mostRecentRecord.lastName,
      q1Events.size,
      q2Events.size,
      q3Events.size,
      q4Events.size,
      q1Events.size + q2Events.size + q3Events.size + q4Events.size,
      mostRecentRecord.date,
      lastEventName,     // ✅ Column M now EXACTLY "Pastoral Check-In"
      guestStatus
    ]);
  });

  return summary;
}

/**
 * Update activity level (Column F) based on attendance in the past 90 days.
 * If last attendance is over 12 months ago → "Archive".
 */
function updateActivityLevels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance Stats");
  if (!sheet || sheet.getLastRow() < 3) return;

  // Get raw attendance records from all sources (already merged/deduped)
  const { rawData } = matchOrAssignBelCodes();
  const today = new Date();

  const cutoff90 = new Date(today);
  cutoff90.setDate(cutoff90.getDate() - 90);

  const cutoff12mo = new Date(today);
  cutoff12mo.setFullYear(cutoff12mo.getFullYear() - 1);

  // Build map: normalizedName -> { count90, lastDate }
  const attMap = new Map();

  rawData.forEach(r => {
    if (!r || r.length < 6) return;
    const bel = r[0]; // kept for completeness, not used for matching to sheet
    const first = r[1];
    const last = r[2];
    const dateVal = r[5];

    if (!first || !last || !dateVal) return;

    const key = normalizeName(`${last}, ${first}`);
    if (!key) return;

    const d = dateVal instanceof Date ? dateVal : new Date(String(dateVal));
    if (isNaN(d.getTime())) return;

    if (!attMap.has(key)) attMap.set(key, { count90: 0, lastDate: null });
    const obj = attMap.get(key);

    if (!obj.lastDate || d > obj.lastDate) obj.lastDate = d;
    if (d >= cutoff90) obj.count90++;
  });

  const lastRow = sheet.getLastRow();

  // Read names from Attendance Stats (Last in C, First in D)
  const nameData = sheet.getRange(3, 3, lastRow - 2, 2).getValues(); // C3:D
  const out = nameData.map(([last, first]) => {
    const key = normalizeName(`${last}, ${first}`);
    const info = key ? attMap.get(key) : null;

    const lastDate = info ? info.lastDate : null;
    const count90 = info ? info.count90 : 0;

    if (lastDate && lastDate < cutoff12mo) return ["Archive"];
    if (count90 >= 12) return ["Core"];
    if (count90 >= 3) return ["Active"];
    return ["Inactive"];
  });

  sheet.getRange(3, 6, out.length, 1).setValues(out);
}

/**
 * Sort + formatting for final output.
 * FIXED: Column M is now horizontally aligned LEFT.
 */
function performFinalSort() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance Stats");
  if (!sheet || sheet.getLastRow() <= 2) return;

  const range = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn());
  const data = range.getValues();

  const order = { "Core": 1, "Active": 2, "Inactive": 3, "Archive": 4 };

  data.sort((a, b) => {
    const guestA = a[4] === "Guest";
    const guestB = b[4] === "Guest";
    if (guestA !== guestB) return guestA ? -1 : 1; // Guests on top

    const aLvl = order[a[5]] || 99;
    const bLvl = order[b[5]] || 99;
    if (aLvl !== bLvl) return aLvl - bLvl; // Core, Active, Inactive, Archive

    // Then sort alphabetically by Last Name (Column C, index 2)
    const lastA = (a[2] || "").toString().toLowerCase();
    const lastB = (b[2] || "").toString().toLowerCase(); // ✅ FIXED HERE
    if (lastA < lastB) return -1;
    if (lastA > lastB) return 1;

    // Optional tie-breaker by First Name (Column D, index 3)
    const firstA = (a[3] || "").toString().toLowerCase();
    const firstB = (b[3] || "").toString().toLowerCase();
    if (firstA < firstB) return -1;
    if (firstA > firstB) return 1;

    return 0;
  });

  range.setValues(data);
  range.clearFormat();

  const numRows = range.getNumRows();

  // Center E–L
  const centerRange = sheet.getRange(3, 5, numRows, 8); // E (5) to L (12)
  centerRange.setHorizontalAlignment("center");
  centerRange.setVerticalAlignment("middle");

  // Left align M only
  const colMRange = sheet.getRange(3, 13, numRows, 1); // Column M
  colMRange.setHorizontalAlignment("left");
  colMRange.setVerticalAlignment("middle");

  Logger.log("✅ Final sort and alignment complete (Column M left-aligned).");
}

/**
 * Main update flow.
 */
function updateAttendanceStatsSheet() {
  Logger.log("🚀 Starting the process to update the 'Attendance Stats' sheet...");

  const data = calculateAttendanceStats();
  if (data.length === 0) {
    Logger.log("No data to update.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance Stats");
  if (!sheet) return;

  const output = data.map(row => {
    const [
      bel, first, last, q1, q2, q3, q4, total, lastDate, lastEvent, guest
    ] = row;

    const formattedDate = lastDate instanceof Date
      ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "MM/dd/yyyy")
      : "";

    return [
      "",
      `${first} ${last}`,
      last,
      first,
      guest,
      "",
      q1,
      q2,
      q3,
      q4,
      total,
      formattedDate,
      lastEvent
    ];
  });

  const maxRows = sheet.getMaxRows();
  if (sheet.getLastRow() > 2) {
    sheet.getRange(3, 1, maxRows - 2, 13).clearContent().clearFormat();
  }

  sheet.getRange(3, 1, output.length, 13).setValues(output);

  updateActivityLevels();
  performFinalSort();

  Logger.log("✅ Finished updating Attendance Stats.");
}

/**
 * Manual run.
 */
function runManualUpdate() {
  updateAttendanceStatsSheet();
  SpreadsheetApp.getUi().alert('The "Attendance Stats" sheet has been successfully updated.');
}
