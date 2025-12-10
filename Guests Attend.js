/**
 * Main function to be run manually to update the "Guests" tab.
 * This script gathers data from "Sunday Service", "Event Attendance",
 * and an external "Directory" sheet to populate guest information.
 * This function CLEARS all data from B4:G and rewrites it.
 */
function updateGuestData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    const serviceSheet = ss.getSheetByName("Sunday Service");
    const eventSheet = ss.getSheetByName("Event Attendance");
    const guestsSheet = ss.getSheetByName("Guests");
    const attendanceLogSheet = ss.getSheetByName("Attendance Log");

    if (!configSheet || !serviceSheet || !eventSheet || !guestsSheet) {
      throw new Error("One or more required sheets are missing (Config, Sunday Service, Event Attendance, Guests).");
    }

    // 1. Get Directory Data
    const directoryId = configSheet.getRange("B2").getValue();
    if (!directoryId) {
      throw new Error("Directory Sheet ID not found in Config tab, cell B2.");
    }
    const directoryMap = getDirectoryData(directoryId);

    // 2. Get Sunday Service Data
    const serviceMap = getServiceData(serviceSheet);

    // 3. Get Event Attendance (Community Intro) Data
    const introMap = getIntroData(eventSheet);

    // 3b. Get Pastoral Check-In data from Attendance Log (for column F fallback)
    let pastoralIntroMap = new Map();
    if (attendanceLogSheet) {
      pastoralIntroMap = getPastoralIntroData(attendanceLogSheet);
    }

    // 4. Get Unique Guest List
    // ONLY those who have attendance in Sunday Service or Event Attendance
    // AND have Column H = "Guest"
    // PLUS any "Pastoral check -In" names that are not in Directory
    const uniqueGuests = getUniqueGuests(serviceSheet, eventSheet, attendanceLogSheet);

    // 5. Process and Prepare Data for Writing
    const finalData = [];

    // Filter out anyone already in Directory (they are NOT guests)
    const filteredGuests = [];
    for (const guest of uniqueGuests.values()) {
      const ln = String(guest.lastName || "").trim();
      const fn = String(guest.firstName || "").trim();
      let inDirectory = false;

      if (ln && fn) {
        const dirKey = (ln + "," + fn).toLowerCase();
        if (directoryMap.has(dirKey)) {
          inDirectory = true;
        }
      }

      if (!inDirectory) {
        filteredGuests.push(guest);
      }
    }

    // Sort guests alphabetically by last name, then first name.
    const sortedGuests = filteredGuests.sort((a, b) => {
      if (a.lastName < b.lastName) return -1;
      if (a.lastName > b.lastName) return 1;
      if (a.firstName < b.firstName) return -1;
      if (a.firstName > b.firstName) return 1;
      return 0;
    });

    // Build the final array for the "Guests" sheet.
    for (const guest of sortedGuests) {
      const key = buildGuestKey(guest.lastName, guest.firstName);
      const fullName = guest.firstName + " " + guest.lastName;

      const serviceDate = key ? (serviceMap.get(key) || "") : "";

      // Column F (Intro) = Community Intro first, else Pastoral Check-In from Attendance Log
      let introDate = key ? (introMap.get(key) || "") : "";
      if (!introDate && key && pastoralIntroMap.has(key)) {
        introDate = pastoralIntroMap.get(key);
      }

      // Registration date from Directory (will normally be blank for true guests,
      // because we filtered out those already in Directory)
      let regDate = "";
      if (guest.lastName && guest.firstName) {
        const dirKey = (guest.lastName + "," + guest.firstName).toLowerCase();
        regDate = directoryMap.get(dirKey) || "";
      }

      // Match the column order: B, C, D, E, F, G
      finalData.push([
        fullName,
        guest.lastName,
        guest.firstName,
        serviceDate,
        introDate,
        regDate
      ]);
    }

    // 6. Write Data to Guests Sheet
    const startRow = 4;
    const numRows = finalData.length;
    
    // Clear old data from row 4 downwards, columns B-G
    if (guestsSheet.getLastRow() >= startRow) {
      guestsSheet.getRange(startRow, 2, guestsSheet.getLastRow() - startRow + 1, 6).clearContent();
    }

    // Write new data if any exists
    if (numRows > 0) {
      guestsSheet.getRange(startRow, 2, numRows, 6).setValues(finalData);
      guestsSheet.getRange(startRow, 5, numRows, 3).setNumberFormat("MM-dd-yy");
    }

    Logger.log("Guest data updated successfully.");

  } catch (e) {
    Logger.log("Error in updateGuestData: " + e);
  }
}

/**
 * Finds new guests and appends them to the "Guests" tab WITH their dates.
 * This function does NOT sort the sheet or remove old guests.
 */
function addNewGuests() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    const serviceSheet = ss.getSheetByName("Sunday Service");
    const eventSheet = ss.getSheetByName("Event Attendance");
    const guestsSheet = ss.getSheetByName("Guests");
    const attendanceLogSheet = ss.getSheetByName("Attendance Log");

    if (!configSheet || !serviceSheet || !eventSheet || !guestsSheet) {
      throw new Error("One or more required sheets are missing (Config, Sunday Service, Event Attendance, Guests).");
    }

    // 1. Get all guests who are ALREADY in the "Guests" tab
    const existingGuestsSet = new Set();
    const startRow = 4;
    if (guestsSheet.getLastRow() >= startRow) {
      const existingNames = guestsSheet.getRange(startRow, 3, guestsSheet.getLastRow() - startRow + 1, 2).getValues(); // C:D
      for (const row of existingNames) {
        const lastName = String(row[0]).trim();
        const firstName = String(row[1]).trim();
        const key = buildGuestKey(lastName, firstName);
        if (key) {
          existingGuestsSet.add(key);
        }
      }
    }
    
    // 2. Get ALL unique guests from the source tabs
    const allGuestsMap = getUniqueGuests(serviceSheet, eventSheet, attendanceLogSheet);

    // 3. Find only the NEW guests
    const newGuests = [];
    for (const [key, guest] of allGuestsMap.entries()) {
      if (!existingGuestsSet.has(key)) {
        newGuests.push(guest); // guest is { firstName, lastName }
      }
    }

    // 4. If no new guests, stop here
    if (newGuests.length === 0) {
      Logger.log("No new guests found.");
      return;
    }

    // 5. Get all date information to populate for the new guests
    const directoryId = configSheet.getRange("B2").getValue();
    if (!directoryId) {
      throw new Error("Directory Sheet ID not found in Config tab, cell B2.");
    }
    const directoryMap = getDirectoryData(directoryId);
    const serviceMap = getServiceData(serviceSheet);
    const introMap = getIntroData(eventSheet);

    let pastoralIntroMap = new Map();
    if (attendanceLogSheet) {
      pastoralIntroMap = getPastoralIntroData(attendanceLogSheet);
    }

    // 6. Sort new guests and format them for the sheet
    const finalData = [];
    newGuests.sort((a, b) => {
      if (a.lastName < b.lastName) return -1;
      if (a.lastName > b.lastName) return 1;
      if (a.firstName < b.firstName) return -1;
      if (a.firstName > b.firstName) return 1;
      return 0;
    });

    for (const guest of newGuests) {
      const ln = String(guest.lastName || "").trim();
      const fn = String(guest.firstName || "").trim();
      let inDirectory = false;

      if (ln && fn) {
        const dirKey = (ln + "," + fn).toLowerCase();
        if (directoryMap.has(dirKey)) {
          inDirectory = true;
        }
      }

      // Skip if already in Directory (not a guest)
      if (inDirectory) {
        continue;
      }

      const fullName = guest.firstName + " " + guest.lastName;
      const key = buildGuestKey(guest.lastName, guest.firstName);

      const serviceDate = key ? (serviceMap.get(key) || "") : "";

      let introDate = key ? (introMap.get(key) || "") : "";
      if (!introDate && key && pastoralIntroMap.has(key)) {
        introDate = pastoralIntroMap.get(key);
      }

      let regDate = "";
      if (guest.lastName && guest.firstName) {
        const dirKey = (guest.lastName + "," + guest.firstName).toLowerCase();
        regDate = directoryMap.get(dirKey) || "";
      }

      // B: Full Name, C: Last, D: First, E: Service, F: Intro, G: Reg
      finalData.push([
        fullName,
        guest.lastName,
        guest.firstName,
        serviceDate,
        introDate,
        regDate
      ]);
    }

    // 7. Write the new guests starting at the next available blank row.
    const existingLastRow = guestsSheet.getLastRow();
    const dataStartRow = 4;
    const blankRows = [];

    if (existingLastRow >= dataStartRow) {
      const nameValues = guestsSheet.getRange(dataStartRow, 3, existingLastRow - dataStartRow + 1, 2).getValues(); // C:D
      for (let i = 0; i < nameValues.length; i++) {
        const ln = String(nameValues[i][0]).trim();
        const fn = String(nameValues[i][1]).trim();
        if (!ln && !fn) {
          blankRows.push(dataStartRow + i);
        }
      }
    }

    let dataIndex = 0;

    // Fill existing empty rows
    for (let i = 0; i < blankRows.length && dataIndex < finalData.length; i++, dataIndex++) {
      const rowIndex = blankRows[i];
      guestsSheet
        .getRange(rowIndex, 2, 1, 6)
        .setValues([finalData[dataIndex]]);
      guestsSheet
        .getRange(rowIndex, 5, 1, 3)
        .setNumberFormat("MM-dd-yy");
    }

    // Append remaining new guests
    if (dataIndex < finalData.length) {
      const remaining = finalData.slice(dataIndex);
      const appendStartRow = Math.max(existingLastRow + 1, dataStartRow);
      const newRowsRange = guestsSheet.getRange(appendStartRow, 2, remaining.length, 6);
      newRowsRange.setValues(remaining);
      newRowsRange.offset(0, 3, remaining.length, 3).setNumberFormat("MM-dd-yy");
    }

    Logger.log(`Added ${finalData.length} new guests with their dates.`);

  } catch (e) {
    Logger.log("Error in addNewGuests: " + e);
  }
}

/**
 * Updates blank dates (E, F, G) for guests already in the "Guests" tab.
 * This function does NOT add or remove rows.
 */
function updateExistingGuestDates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    const serviceSheet = ss.getSheetByName("Sunday Service");
    const eventSheet = ss.getSheetByName("Event Attendance");
    const guestsSheet = ss.getSheetByName("Guests");
    const attendanceLogSheet = ss.getSheetByName("Attendance Log");

    if (!configSheet || !serviceSheet || !eventSheet || !guestsSheet) {
      throw new Error("One or more required sheets are missing (Config, Sunday Service, Event Attendance, Guests).");
    }

    // 1. Get all date data from all sources
    const directoryId = configSheet.getRange("B2").getValue();
    if (!directoryId) {
      throw new Error("Directory Sheet ID not found in Config tab, cell B2.");
    }
    const directoryMap = getDirectoryData(directoryId);
    const serviceMap = getServiceData(serviceSheet);
    const introMap = getIntroData(eventSheet);

    let pastoralIntroMap = new Map();
    if (attendanceLogSheet) {
      pastoralIntroMap = getPastoralIntroData(attendanceLogSheet);
    }

    // 2. Get the current data from the "Guests" sheet
    const startRow = 4;
    const lastRow = guestsSheet.getLastRow();
    if (lastRow < startRow) {
      Logger.log("No guests to update.");
      return;
    }
    
    const numRows = lastRow - startRow + 1;
    const dataRange = guestsSheet.getRange(startRow, 3, numRows, 5); // C:G
    const values = dataRange.getValues();

    let updatesMade = 0;
    const datesToWrite = [];

    // 3. Loop through each guest and fill in blank dates
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const lastName = String(row[0]).trim();
      const firstName = String(row[1]).trim();
      
      let serviceDate = row[2];
      let introDate = row[3];
      let regDate = row[4];

      if (lastName || firstName) {
        const key = buildGuestKey(lastName, firstName);

        if (key) {
          if (!serviceDate && serviceMap.has(key)) {
            serviceDate = serviceMap.get(key);
            updatesMade++;
          }
          
          if (!introDate) {
            if (introMap.has(key)) {
              introDate = introMap.get(key);
              updatesMade++;
            } else if (pastoralIntroMap.has(key)) {
              introDate = pastoralIntroMap.get(key);
              updatesMade++;
            }
          }

          if (!regDate && lastName && firstName) {
            const dirKey = (lastName + "," + firstName).toLowerCase();
            if (directoryMap.has(dirKey)) {
              regDate = directoryMap.get(dirKey);
              updatesMade++;
            }
          }
        }
      }
      
      datesToWrite.push([serviceDate, introDate, regDate]);
    }

    // 4. Write all updated dates back
    if (updatesMade > 0) {
      const dateRangeToWrite = guestsSheet.getRange(startRow, 5, numRows, 3); // E:G
      dateRangeToWrite.setValues(datesToWrite);
      dateRangeToWrite.setNumberFormat("MM-dd-yy");
      Logger.log(`Updated ${updatesMade} blank dates for existing guests.`);
    } else {
      Logger.log("No blank dates found to update.");
    }

  } catch (e) {
    Logger.log("Error in updateExistingGuestDates: " + e);
  }
}


/**
 * Gets registration data from the external Directory sheet.
 * @param {string} sheetId The ID of the external Google Sheet.
 * @returns {Map<string, Date>} A Map where key is "lastname,firstname"
 * and value is the registration Date object.
 */
function getDirectoryData(sheetId) {
  const directoryMap = new Map();
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Directory");
    if (!sheet) {
      Logger.log("Error: 'Directory' tab not found in external sheet.");
      return directoryMap;
    }
    
    // Data starts from row 2 (assuming row 1 is header)
    const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 19).getValues(); // C to U

    for (const row of data) {
      const lastName = String(row[0]).trim();  // Column C (index 0)
      const firstName = String(row[1]).trim(); // Column D (index 1)
      const regDate = row[18]; // Column U (index 18)

      if (firstName && lastName && regDate instanceof Date) {
        const key = (lastName + "," + firstName).toLowerCase();
        if (!directoryMap.has(key)) {
          directoryMap.set(key, regDate);
        }
      }
    }
  } catch (e) {
    Logger.log("Error accessing Directory sheet: " + e);
    throw new Error("Error accessing Directory sheet. Check ID and permissions. " + e.message);
  }
  return directoryMap;
}

/**
 * Gets the first service date for all GUESTS from the "Sunday Service" sheet.
 * Uses Column H = "Guest" and allows first-name-only or last-name-only.
 * @param {Sheet} sheet The "Sunday Service" Google Sheet object.
 * @returns {Map<string, Date>} A Map where key is name key (via buildGuestKey)
 * and value is the first service Date object.
 */
function getServiceData(sheet) {
  const serviceMap = new Map();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Get service dates from row 2, starting column I (index 8)
  const serviceDates = values[1].slice(8);
  
  // Data starts from row 4 (index 3)
  for (let i = 3; i < values.length; i++) {
    const row = values[i];
    const lastName = String(row[2]).trim();  // Column C (index 2)
    const firstName = String(row[3]).trim(); // Column D (index 3)
    const status = String(row[7]).trim();    // Column H (index 7)

    // Only rows marked Guest, and have at least first OR last name
    if ((firstName || lastName) && status === "Guest") {
      const key = buildGuestKey(lastName, firstName);
      if (!key) continue;

      if (!serviceMap.has(key)) {
        const attendance = row.slice(8);
        for (let j = 0; j < attendance.length; j++) {
          if (attendance[j] === true) {
            if (serviceDates[j] instanceof Date) {
              serviceMap.set(key, serviceDates[j]);
              break;
            }
          }
        }
      }
    }
  }
  return serviceMap;
}

/**
 * Gets the first "Community Intro" date for all GUESTS
 * from the "Event Attendance" sheet.
 * Uses Column H = "Guest" and allows first-name-only or last-name-only.
 * @param {Sheet} sheet The "Event Attendance" Google Sheet object.
 * @returns {Map<string, Date>} A Map where key is name key (via buildGuestKey)
 * and value is the event Date object.
 */
/**
 * Gets the first "Intro/Orientation" date for all GUESTS
 * from the "Event Attendance" sheet.
 * Uses Column H = "Guest" and allows first-name-only or last-name-only.
 * @param {Sheet} sheet The "Event Attendance" Google Sheet object.
 * @returns {Map<string, Date>} A Map where key is name key (via buildGuestKey)
 * and value is the event Date object.
 */
function getIntroData(sheet) {
  const introMap = new Map();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Get event dates from row 2, starting column I (index 8)
  const eventDates = values[1].slice(8); // Row 2
  
  // Get event names from row 3, starting column I (index 8)
  const eventNames = values[2].slice(8); // Row 3
  
  // Data starts from row 5 (index 4)
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const lastName = String(row[2]).trim();  // Column C (index 2)
    const firstName = String(row[3]).trim(); // Column D (index 3)
    const status = String(row[7]).trim();    // Column H (index 7)

    // Only rows marked Guest, and have at least first OR last name
    if ((firstName || lastName) && status === "Guest") {
      const key = buildGuestKey(lastName, firstName);
      if (!key) continue;

      if (!introMap.has(key)) {
        const attendance = row.slice(8);
        for (let j = 0; j < attendance.length; j++) {
          if (attendance[j] === true) { 
            const eventName = String(eventNames[j]).toLowerCase();

            // Match any event whose name contains "community intro",
            // "orientation", or "orient"
            if (
              eventName.includes("community intro") ||
              eventName.includes("orientation") ||
              eventName.includes("orient")
            ) {
              const eventDate = eventDates[j];
              if (eventDate instanceof Date) {
                introMap.set(key, eventDate);
                break;
              }
            }
          }
        }
      }
    }
  }
  return introMap;
}


/**
 * Gets the first "Pastoral check -In" date for all names
 * from the "Attendance Log" sheet.
 * Uses:
 * - Column C: Last Name
 * - Column D: First Name
 * - Column B: Date
 * - Column F: Event name (must be "Pastoral check -In")
 *
 * Allows first-name-only or last-name-only, using buildGuestKey.
 *
 * @param {Sheet} sheet The "Attendance Log" sheet.
 * @returns {Map<string, Date>} A Map where key is name key (via buildGuestKey)
 * and value is the first Pastoral Check-In Date object.
 */
function getPastoralIntroData(sheet) {
  const pastoralMap = new Map();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const lastName = String(row[2]).trim();  // Column C (index 2)
    const firstName = String(row[3]).trim(); // Column D (index 3)
    const dateVal = row[1];                  // Column B (index 1)
    const eventNameRaw = String(row[5]).trim(); // Column F (index 5)

    if (!firstName && !lastName) {
      continue;
    }

    const eventName = eventNameRaw.toLowerCase();
    if (eventName === "pastoral check -in") {
      if (dateVal instanceof Date) {
        const key = buildGuestKey(lastName, firstName);
        if (!key) continue;

        if (!pastoralMap.has(key)) {
          pastoralMap.set(key, dateVal);
        } else {
          const existing = pastoralMap.get(key);
          if (dateVal < existing) {
            pastoralMap.set(key, dateVal);
          }
        }
      }
    }
  }

  return pastoralMap;
}

/**
 * Build a name key that also works when only first name or only last name
 * is present. Used for de-duplicating guests and matching across functions.
 *
 * Rules:
 * - If both last & first: "lastname,firstname" (lowercase)
 * - If only last: "lastname"
 * - If only first: "firstname"
 * - If both empty: "" (no key)
 */
function buildGuestKey(lastName, firstName) {
  const ln = String(lastName || "").trim().toLowerCase();
  const fn = String(firstName || "").trim().toLowerCase();
  if (ln && fn) return ln + "," + fn;
  if (ln) return ln;
  if (fn) return fn;
  return "";
}

/**
 * Compiles a unique list of GUESTS from:
 * - Sunday Service (Column H = "Guest", with attendance)
 * - Event Attendance (Column H = "Guest", with attendance)
 * - Attendance Log: "Pastoral check -In" rows
 *
 * For Pastoral Check-In:
 * - We collect all "Pastoral check -In" names here as guest candidates.
 * - Later, the caller filters out those already in Directory.
 *
 * Allows first-name-only or last-name-only (using buildGuestKey).
 *
 * @param {Sheet} serviceSheet The "Sunday Service" sheet.
 * @param {Sheet} eventSheet The "Event Attendance" sheet.
 * @param {Sheet} attendanceLogSheet The "Attendance Log" sheet (optional).
 * @returns {Map<string, Object>} A Map of unique guests,
 * with key from buildGuestKey and {firstName, lastName} as value.
 */
function getUniqueGuests(serviceSheet, eventSheet, attendanceLogSheet) {
  const guests = new Map();

  // --- 1. Guests from Sunday Service WITH attendance ---
  const serviceValues = serviceSheet.getDataRange().getValues();
  if (serviceValues.length >= 4) {
    for (let i = 3; i < serviceValues.length; i++) { // Data from row 4
      const row = serviceValues[i];
      const lastName = String(row[2]).trim();  // Col C
      const firstName = String(row[3]).trim(); // Col D
      const status = String(row[7]).trim();    // Col H

      // Must be marked Guest, and have at least first or last name
      if ((firstName || lastName) && status === "Guest") {
        const attendance = row.slice(8); // Checkboxes from Col I onwards
        const hasAttendance = attendance.some(v => v === true);

        if (hasAttendance) {
          const key = buildGuestKey(lastName, firstName);
          if (key && !guests.has(key)) {
            guests.set(key, { firstName: firstName, lastName: lastName });
          }
        }
      }
    }
  }

  // --- 2. Guests from Event Attendance WITH attendance ---
  const eventValues = eventSheet.getDataRange().getValues();
  if (eventValues.length >= 5) {
    for (let i = 4; i < eventValues.length; i++) { // Data from row 5
      const row = eventValues[i];
      const lastName = String(row[2]).trim();  // Col C
      const firstName = String(row[3]).trim(); // Col D
      const status = String(row[7]).trim();    // Col H

      if ((firstName || lastName) && status === "Guest") {
        const attendance = row.slice(8); // Checkboxes from Col I onwards
        const hasAttendance = attendance.some(v => v === true);

        if (hasAttendance) {
          const key = buildGuestKey(lastName, firstName);
          if (key && !guests.has(key)) {
            guests.set(key, { firstName: firstName, lastName: lastName });
          }
        }
      }
    }
  }

  // --- 3. Guest candidates from Pastoral Check-In (Attendance Log) ---
  if (attendanceLogSheet) {
    const logValues = attendanceLogSheet.getDataRange().getValues();
    if (logValues.length >= 2) {
      for (let i = 1; i < logValues.length; i++) { // Data from row 2
        const row = logValues[i];
        const lastName = String(row[2]).trim();  // Col C
        const firstName = String(row[3]).trim(); // Col D
        const eventNameRaw = String(row[5]).trim(); // Col F

        if (!firstName && !lastName) {
          continue;
        }

        const eventName = eventNameRaw.toLowerCase();
        if (eventName === "pastoral check -in") {
          const key = buildGuestKey(lastName, firstName);
          if (key && !guests.has(key)) {
            guests.set(key, { firstName: firstName, lastName: lastName });
          }
        }
      }
    }
  }
  
  return guests;
}
