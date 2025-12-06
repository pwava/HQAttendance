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

    // 3. Get Event Attendance (Orientation) Data
    const introMap = getIntroData(eventSheet);

    // 3b. Get Pastoral Check-In data from Attendance Log (for column F)
    let pastoralIntroMap = new Map();
    if (attendanceLogSheet) {
      pastoralIntroMap = getPastoralIntroData(attendanceLogSheet);
    }

    // 4. Get Unique Guest List (ONLY those who have attendance)
    const uniqueGuests = getUniqueGuests(serviceSheet, eventSheet);

    // 5. Process and Prepare Data for Writing
    const finalData = [];
    
    // Sort guests alphabetically by last name, then first name.
    const sortedGuests = Array.from(uniqueGuests.values()).sort((a, b) => {
      if (a.lastName < b.lastName) return -1;
      if (a.lastName > b.lastName) return 1;
      if (a.firstName < b.firstName) return -1;
      if (a.firstName > b.firstName) return 1;
      return 0;
    });

    // Build the final array for the "Guests" sheet.
    for (const guest of sortedGuests) {
      // Create a standardized key for matching across maps.
      const key = (guest.lastName.trim() + "," + guest.firstName.trim()).toLowerCase();
      
      const fullName = guest.firstName + " " + guest.lastName;
      const serviceDate = serviceMap.get(key) || ""; // Get date or empty string

      // Column F (Intro) = Community Intro first, else Pastoral Check-In from Attendance Log
      let introDate = introMap.get(key) || "";
      if (!introDate && pastoralIntroMap.has(key)) {
        introDate = pastoralIntroMap.get(key);
      }

      const regDate = directoryMap.get(key) || ""; // Get date or empty string

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
    // Start writing from row 4, column 2 (B)
    const startRow = 4;
    const numRows = finalData.length;
    
    // Clear old data from row 4 downwards, columns B-G
    if (guestsSheet.getLastRow() >= startRow) {
      guestsSheet.getRange(startRow, 2, guestsSheet.getLastRow() - startRow + 1, 6).clearContent();
    }

    // Write new data if any exists
    if (numRows > 0) {
      guestsSheet.getRange(startRow, 2, numRows, 6).setValues(finalData);
      
      // Apply date formatting
      guestsSheet.getRange(startRow, 5, numRows, 3).setNumberFormat("MM-dd-yy");
    }

    Logger.log("Guest data updated successfully.");

  } catch (e) {
    Logger.log("Error in updateGuestData: " + e);
    // SpreadsheetApp.getUi().alert("Error: " + e.message);
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
        if (lastName && firstName) {
          existingGuestsSet.add((lastName + "," + firstName).toLowerCase());
        }
      }
    }
    
    // 2. Get ALL unique guests from the source tabs
    //    ONLY those who actually have attendance
    const allGuestsMap = getUniqueGuests(serviceSheet, eventSheet);

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
    // This is necessary to add their dates
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
      const fullName = guest.firstName + " " + guest.lastName;
      const key = (guest.lastName.trim() + "," + guest.firstName.trim()).toLowerCase();

      // Look up the dates for the new guest
      const serviceDate = serviceMap.get(key) || "";

      // Column F (Intro) = Community Intro first, else Pastoral Check-In from Attendance Log
      let introDate = introMap.get(key) || "";
      if (!introDate && pastoralIntroMap.has(key)) {
        introDate = pastoralIntroMap.get(key);
      }

      const regDate = directoryMap.get(key) || "";

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

    // 7. Write the new guests (with their dates) starting at the next available blank row,
    //    not always at the very bottom.
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

    // Fill existing completely blank rows (no last name & no first name)
    for (let i = 0; i < blankRows.length && dataIndex < finalData.length; i++, dataIndex++) {
      const rowIndex = blankRows[i];
      guestsSheet
        .getRange(rowIndex, 2, 1, 6)
        .setValues([finalData[dataIndex]]);
      guestsSheet
        .getRange(rowIndex, 5, 1, 3)
        .setNumberFormat("MM-dd-yy");
    }

    // If there are still new guests left, append them below the current last row
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
    // SpreadsheetApp.getUi().alert("Error: " + e.message);
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
      return; // No data to update
    }
    
    const numRows = lastRow - startRow + 1;
    // Get range C:G (LastName, FirstName, ServiceDate, IntroDate, RegDate)
    const dataRange = guestsSheet.getRange(startRow, 3, numRows, 5);
    const values = dataRange.getValues();

    let updatesMade = 0;
    const datesToWrite = []; // We will build a new array for writing dates

    // 3. Loop through each guest and fill in blank dates
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const lastName = String(row[0]).trim();
      const firstName = String(row[1]).trim();
      
      // Get existing dates
      let serviceDate = row[2];
      let introDate = row[3];
      let regDate = row[4];

      if (lastName && firstName) {
        const key = (lastName + "," + firstName).toLowerCase();

        // Check if Service Date (Col E) is blank and we have a new date
        if (!serviceDate && serviceMap.has(key)) {
          serviceDate = serviceMap.get(key);
          updatesMade++;
        }
        
        // Check if Intro Date (Col F) is blank and we have a new date
        if (!introDate) {
          if (introMap.has(key)) {
            introDate = introMap.get(key);
            updatesMade++;
          } else if (pastoralIntroMap.has(key)) {
            introDate = pastoralIntroMap.get(key);
            updatesMade++;
          }
        }

        // Check if Reg Date (Col G) is blank and we have a new date
        if (!regDate && directoryMap.has(key)) {
          regDate = directoryMap.get(key);
          updatesMade++;
        }
      }
      
      // Add the (potentially updated) dates to our write-back array
      datesToWrite.push([serviceDate, introDate, regDate]);
    }

    // 4. Write all the updated dates back to the sheet in one go
    if (updatesMade > 0) {
      const dateRangeToWrite = guestsSheet.getRange(startRow, 5, numRows, 3); // Cols E, F, G
      dateRangeToWrite.setValues(datesToWrite);
      dateRangeToWrite.setNumberFormat("MM-dd-yy");
      Logger.log(`Updated ${updatesMade} blank dates for existing guests.`);
    } else {
      Logger.log("No blank dates found to update.");
    }

  } catch (e) {
    Logger.log("Error in updateExistingGuestDates: " + e);
    // SpreadsheetApp.getUi().alert("Error: " + e.message);
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
    // This error will bubble up to the main function
    throw new Error("Error accessing Directory sheet. Check ID and permissions. " + e.message);
  }
  return directoryMap;
}

/**
 * Gets the first service date for all guests from the "Sunday Service" sheet.
 * @param {Sheet} sheet The "Sunday Service" Google Sheet object.
 * @returns {Map<string, Date>} A Map where key is "lastname,firstname"
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
    const status = String(row[7]).trim();  // Column H (index 7)

    if (firstName && lastName && status === "Guest") {
      const key = (lastName + "," + firstName).toLowerCase();
      
      // Only find the first date, so if guest is already processed, skip.
      if (!serviceMap.has(key)) {
        // Check attendance checkboxes starting from Column I (index 8)
        const attendance = row.slice(8);
        for (let j = 0; j < attendance.length; j++) {
          if (attendance[j] === true) { // Checkbox is checked
            if (serviceDates[j] instanceof Date) {
              serviceMap.set(key, serviceDates[j]);
              break; // Found the first date, stop checking this row
            }
          }
        }
      }
    }
  }
  return serviceMap;
}

/**
 * [UPDATED FUNCTION]
 * Gets the first "Community Intro" date for all guests
 * from the "Event Attendance" sheet.
 * @param {Sheet} sheet The "Event Attendance" Google Sheet object.
 * @returns {Map<string, Date>} A Map where key is "lastname,firstname"
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
    const status = String(row[7]).trim();  // Column H (index 7)

    if (firstName && lastName && status === "Guest") {
      const key = (lastName + "," + firstName).toLowerCase();

      // Only find the first intro date, so if guest is processed, skip.
      if (!introMap.has(key)) {
        
        // Get the attendance checkboxes for this row, starting Col I (index 8)
        const attendance = row.slice(8);
        
        // Loop through all event checkboxes for this guest
        for (let j = 0; j < attendance.length; j++) {
          
          // Check if checkbox is checked (true)
          if (attendance[j] === true) { 
            
            // Get the corresponding event name from Row 3
            const eventName = String(eventNames[j]).toLowerCase();
            
            // Check if this event is "Community Intro"
            if (eventName.includes("community intro")) {
              
              // Get the corresponding date from Row 2
              const eventDate = eventDates[j];
              
              if (eventDate instanceof Date) {
                introMap.set(key, eventDate);
                break; // Found the first intro date, stop checking this row
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
 * Gets the first "Pastoral check -In" date for all guests
 * from the "Attendance Log" sheet.
 * Uses:
 * - Column C: Last Name
 * - Column D: First Name
 * - Column B: Date
 * - Column F: Event name (must be "Pastoral check -In")
 *
 * @param {Sheet} sheet The "Attendance Log" sheet.
 * @returns {Map<string, Date>} A Map where key is "lastname,firstname"
 * and value is the first Pastoral Check-In Date object.
 */
function getPastoralIntroData(sheet) {
  const pastoralMap = new Map();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Assume headers in row 1; data starts row 2 (index 1)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const lastName = String(row[2]).trim();  // Column C (index 2)
    const firstName = String(row[3]).trim(); // Column D (index 3)
    const dateVal = row[1];                  // Column B (index 1)
    const eventNameRaw = String(row[5]).trim(); // Column F (index 5)

    if (!firstName || !lastName) {
      continue;
    }

    const eventName = eventNameRaw.toLowerCase();
    // Only for event name exactly "Pastoral check -In" (case-insensitive)
    if (eventName === "pastoral check -in") {
      if (dateVal instanceof Date) {
        const key = (lastName + "," + firstName).toLowerCase();

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
 * Compiles a unique list of guests from both service and event sheets,
 * but ONLY includes guests who actually have at least one attendance
 * (checkbox checked) in Sunday Service or Event Attendance.
 * @param {Sheet} serviceSheet The "Sunday Service" sheet.
 * @param {Sheet} eventSheet The "Event Attendance" sheet.
 * @returns {Map<string, Object>} A Map of unique guests, with "lastname,firstname"
 * as key and {firstName, lastName} as value.
 */
function getUniqueGuests(serviceSheet, eventSheet) {
  const guests = new Map();

  // 1. Guests from Sunday Service WITH attendance
  const serviceValues = serviceSheet.getDataRange().getValues();
  if (serviceValues.length >= 4) {
    for (let i = 3; i < serviceValues.length; i++) { // Data from row 4
      const row = serviceValues[i];
      const lastName = String(row[2]).trim();  // Col C
      const firstName = String(row[3]).trim(); // Col D
      const status = String(row[7]).trim();    // Col H

      if (firstName && lastName && status === "Guest") {
        const attendance = row.slice(8); // Checkboxes from Col I onwards
        const hasAttendance = attendance.some(v => v === true);
        if (hasAttendance) {
          const key = (lastName + "," + firstName).toLowerCase();
          if (!guests.has(key)) {
            guests.set(key, { firstName: firstName, lastName: lastName });
          }
        }
      }
    }
  }

  // 2. Guests from Event Attendance WITH attendance (any event)
  const eventValues = eventSheet.getDataRange().getValues();
  if (eventValues.length >= 5) {
    for (let i = 4; i < eventValues.length; i++) { // Data from row 5
      const row = eventValues[i];
      const lastName = String(row[2]).trim();  // Col C
      const firstName = String(row[3]).trim(); // Col D
      const status = String(row[7]).trim();    // Col H

      if (firstName && lastName && status === "Guest") {
        const attendance = row.slice(8); // Checkboxes from Col I onwards
        const hasAttendance = attendance.some(v => v === true);
        if (hasAttendance) {
          const key = (lastName + "," + firstName).toLowerCase();
          if (!guests.has(key)) {
            guests.set(key, { firstName: firstName, lastName: lastName });
          }
        }
      }
    }
  }
  
  return guests;
}
