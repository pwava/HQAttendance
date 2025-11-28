/**
* Processes the 'Attendance Log' sheet and updates 'Sunday Service' and 'Event Attendance' sheets.
*
* - Only processes rows where Column I (Status) is blank.
* - Marks successfully processed rows as "Logged" in Column I.
* - Logs with event "Sunday Service" go to the 'Sunday Service' sheet.
* - Logs with event "Pastoral Check-In" go to the 'Pastoral Check-In' sheet.
* - All other event logs go to the 'Event Attendance' sheet.
* - 'Event Attendance' sheet dynamically adds new columns for new event/date combinations.
* - New names found in the log ("Guests") are automatically added to the
* *next available blank row* in 'Sunday Service', 'Pastoral Check-In', or 'Event Attendance'.
* - New names are **properly capitalized** (e.g., "Smith, John").
*/
function processAttendanceLogV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- CONFIGURATION ---
  const logSheetName = 'Attendance Log';
  const sunServiceSheetName = 'Sunday Service';
  const eventSheetName = 'Event Attendance'; // REFINED: Was 'Events'
  const pastoralSheetName = 'Pastoral Check-In';
  
  const nameColumnLetter = 'B'; // This is passed to prepareSheetData but will be ignored
  const typeColumnLetter = 'F'; // Type column on all destination sheets
  const typeColumnIndex = 6;  // Column F is index 6
  
  // 'Sunday Service' sheet config
  const sunServiceDateRow = 2;
  const sunServiceDataStartCol = 9; // REFINED: Was 7 (Column G), now 9 (Column I)
  const sunServiceDataStartRow = 4; // Sunday Service data starts on row 4

  // 'Event Attendance' sheet config
  const eventDateRow = 2;
  const eventNameRow = 3;
  const eventCountRow = 4;  
  const eventDataStartCol = 9; // REFINED: Was 7 (Column G), now 9 (Column I)
  const eventDataStartRow = 5; // Events data starts on row 5
  
  // 'Pastoral Check-In' sheet config
  const pastoralDataStartRow = 4; // 3 header rows, data starts on row 4
  const pastoralRecentDateCol = 5;  // Column E
  const pastoralPreviousDateCol = 6; // Column F
  const pastoralNotesCol = 7; // Column G
  const pastoralExtraCol = 8; // Column H (from Attendance Log Col L)
  
  // --- REFINEMENT ---
  // Column I (index 9) is used for "Logged" status
  const logStatusColumn = 9;  
  // Column J (index 10) is used for "Remarks"
  const logRemarksColumn = 10;
  // Read from Col B (2) to Col L (12). 12 - 2 + 1 = 11 columns
  const logNumColsToRead = 11;  
  // --- Indices for the B:L logData array ---
  const logLastNameIndex = 1;  // Col C
  const logFirstNameIndex = 2; // Col D
  const logTypeIndex = 3;      // Col E
  const logStatusColIndex = 7; // Index of Status (Col I)
  const logRemarksColIndex = 8; // Index of Remarks (Col J)
  const logNotesColIndex = 9; // Col K (Notes for Pastoral Check-In)
  const logExtraColIndex = 10; // Col L (extra for Pastoral Check-In -> Col H)
  // --- END REFINEMENT ---

  const logSheet = ss.getSheetByName(logSheetName);
  if (!logSheet) {
    Logger.log(`Error: Source sheet "${logSheetName}" not found.`);
    return;
  }

  // 1. Read all attendance log data at once (B:L)
  const logRange = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, logNumColsToRead);  
  const logData = logRange.getValues();
  
  const attendanceRecords = [];
  
  // --- REFINEMENT: Filter for unprocessed rows ---
  for (let i = 0; i < logData.length; i++) {
    const row = logData[i];
    const status = row[logStatusColIndex]; // Get status from Col I

    if (status === 'Logged') {
      continue; // Skip this row, it's already done
    }
    
    const eventName = row[4]; // Col F (index 4)
    let eventDate = row[5]; // Col G (index 5)
    
    // --- REFINED LOGIC: Use C/D as the "source of truth" for the name ---
    const logLastName = row[logLastNameIndex];  // Col C
    const logFirstName = row[logFirstNameIndex]; // Col D
    let standardizedName;

    if (logLastName && logFirstName) {
      // Use C/D as the "source of truth"
      standardizedName = standardizeNameHelper(`${logLastName}, ${logFirstName}`);
    } else {
      // Fallback to Col B only if C/D are missing
      const nameFromColB = row[0]; // Col B
      if (!nameFromColB) {
        continue; // No name at all, skip
      }
      standardizedName = standardizeNameHelper(nameFromColB);
    }
    // --- END REFINED LOGIC ---
    
    // --- REFINED DATE CHECK: Try to fix text dates ---
    if (!(eventDate instanceof Date) && eventDate) {
      try {
        eventDate = new Date(eventDate);
        if (isNaN(eventDate.getTime())) { // Check if new Date() was valid
          throw new Error("Invalid date string");
        }
      } catch(e) {
        Logger.log(`Skipping log row ${i + 2} for "${standardizedName}": Invalid date format in Col G (${row[5]}).`);
        logData[i][logRemarksColIndex] = "Skipped: Invalid date format.";
        continue; // Date is not valid, skip this row
      }
    }
    // --- END REFINED DATE CHECK ---

    if (eventName && eventDate instanceof Date && standardizedName) {
      const formattedFullDate = `${eventDate.getMonth() + 1}-${eventDate.getDate()}-${eventDate.getFullYear()}`;
      const formattedShortDate = `${eventDate.getMonth() + 1}-${eventDate.getDate()}`;
      
      attendanceRecords.push({
        name: standardizedName,
        eventName: eventName.toString().trim(),
        eventDate: eventDate,
        formattedFullDate: formattedFullDate, // M-D-YYYY
        formattedShortDate: formattedShortDate, // M-D
        notes: row[logNotesColIndex], // Col K
        extra: row[logExtraColIndex], // Col L
        originalLogRownum: i + 2
      });
    }
  }
  // --- END REFINEMENT ---
  
  if (attendanceRecords.length === 0) {
    Logger.log('No *new* valid attendance records found in the log.');
    return;
  }

  Logger.log(`Processing ${attendanceRecords.length} new records.`);

  // 2. Prepare caches to hold sheet data
  const sunServiceSheet = ss.getSheetByName(sunServiceSheetName);
  const eventSheet = ss.getSheetByName(eventSheetName);
  const pastoralSheet = ss.getSheetByName(pastoralSheetName);

  let sunServiceData, eventSheetData, pastoralData;

  // Prepare 'Sunday Service' sheet data
  if (sunServiceSheet) {
    sunServiceData = prepareSheetData(
      sunServiceSheet,  
      sunServiceDataStartRow,  
      sunServiceDataStartCol,  
      nameColumnLetter,
      [sunServiceDateRow],  
      false,  
      false  
    );
  } else {
    Logger.log(`Warning: "${sunServiceSheetName}" not found. Skipping.`);
  }

  // Prepare 'Event Attendance' sheet data
  if (eventSheet) {
    eventSheetData = prepareSheetData(
      eventSheet,  
      eventDataStartRow,  
      eventDataStartCol,  
      nameColumnLetter,
      [eventDateRow, eventNameRow],  
      true,  
      true  
    );
  } else {
    Logger.log(`Warning: "${eventSheetName}" not found. Skipping.`);
  }

  // Prepare 'Pastoral Check-In' sheet data
  if (pastoralSheet) {
    pastoralData = preparePastoralSheetData(pastoralSheet, pastoralDataStartRow);
  } else {
    Logger.log(`Warning: "${pastoralSheetName}" not found. Skipping.`);
  }

  // 3. Process all *new* records in memory
  let recordsWereLogged = false;
  
  // --- This set tracks unique name/date pairs to prevent duplicates ---
  const processedLogs = new Set();

  for (const record of attendanceRecords) {
    const { name, eventName, eventDate, formattedFullDate, formattedShortDate, notes, extra } = record;
    const logDataIndex = record.originalLogRownum - 2;

    // --- Check for duplicates IN THE LOG FILE ---
    const logKey = `${name}|${eventName}|${formattedFullDate}`;
    if (processedLogs.has(logKey)) {
        logData[logDataIndex][logStatusColIndex] = 'Logged';
        logData[logDataIndex][logRemarksColIndex] = 'Duplicate log entry processed.';
        recordsWereLogged = true;
        continue;
    }

    try {
      // --- REFINED: Make "Sunday Service" check case-insensitive ---
      if (/sunday service/i.test(eventName)) {
        // --- CASE 1: Sunday Service ---
        if (!sunServiceData) continue;
        
        let row = sunServiceData.nameMap.get(name);
        const col = sunServiceData.dateMap.get(formattedShortDate);
        
        if (row && col) {
          const arrayRow = row - sunServiceDataStartRow;  
          const arrayCol = col - sunServiceDataStartCol;
          if (sunServiceData.checkboxes[arrayRow] && sunServiceData.checkboxes[arrayRow][arrayCol] !== undefined) {
            sunServiceData.checkboxes[arrayRow][arrayCol] = true;
            logData[logDataIndex][logStatusColIndex] = 'Logged';
            logData[logDataIndex][logRemarksColIndex] = '';
            recordsWereLogged = true;
            processedLogs.add(logKey);
          }
        } else if (!row) {
          const guestLogEntry = logData[logDataIndex];
          const lastName = guestLogEntry[logLastNameIndex];
          const firstName = guestLogEntry[logFirstNameIndex];

          const capitalizedLastName = capitalizeName(lastName);
          const capitalizedFirstName = capitalizeName(firstName);
          
          const nextRow = sunServiceData.nextBlankRow;
          
          sunServiceSheet.getRange(nextRow, 3).setValue(capitalizedLastName); // Col C
          sunServiceSheet.getRange(nextRow, 4).setValue(capitalizedFirstName); // Col D
          
          SpreadsheetApp.flush(); 
          
          sunServiceData.nameMap.set(name, nextRow);
          
          const numCols = sunServiceData.checkboxes[0] ? sunServiceData.checkboxes[0].length : 0;
          const newCheckboxRow = Array(numCols).fill(false);

          if (col) {
            const arrayCol = col - sunServiceDataStartCol;
            newCheckboxRow[arrayCol] = true;  
            logData[logDataIndex][logStatusColIndex] = 'Logged';
            logData[logDataIndex][logRemarksColIndex] = 'New person added to directory.';
            processedLogs.add(logKey);
          } else {
            logData[logDataIndex][logRemarksColIndex] = 'New person added, but event date not found.';
          }
          
          sunServiceData.checkboxes.push(newCheckboxRow);
          sunServiceData.numRows++;
          sunServiceData.nextBlankRow++;
          recordsWereLogged = true;

        } else if (row && !col) {
          logData[logDataIndex][logStatusColIndex] = '';
          logData[logDataIndex][logRemarksColIndex] = 'Date not found in Sunday Service sheet.';
          recordsWereLogged = true;
        }

      } else if (/pastoral check-?in/i.test(eventName)) {
        // --- CASE 2: Pastoral Check-In ---
        if (!pastoralData) continue;

        let row = pastoralData.nameMap.get(name);

        if (row) {
          const recentCell = pastoralSheet.getRange(row, pastoralRecentDateCol); // E
          const prevCell = pastoralSheet.getRange(row, pastoralPreviousDateCol); // F
          const existingRecent = recentCell.getValue();

          // shift recent → previous
          if (existingRecent) {
            prevCell.setValue(existingRecent);
            prevCell.setHorizontalAlignment("center");   // center align
            prevCell.setVerticalAlignment("middle");
          }

          // write new recent check-in
          recentCell.setValue(eventDate);
          recentCell.setHorizontalAlignment("center");   // center align
          recentCell.setVerticalAlignment("middle");

          // write notes (always replace, even if blank)
          const notesCell = pastoralSheet.getRange(row, pastoralNotesCol); // G
          notesCell.setValue(notes);
          notesCell.setHorizontalAlignment("left");    // left align
          notesCell.setVerticalAlignment("middle");

          // write extra from Col L into Col H (pastor, always replace)
          const extraCell = pastoralSheet.getRange(row, pastoralExtraCol); // H
          extraCell.setValue(extra);
          extraCell.setHorizontalAlignment("left");    // LEFT ALIGN

          logData[logDataIndex][logStatusColIndex] = 'Logged';
          logData[logDataIndex][logRemarksColIndex] = '';
          recordsWereLogged = true;
          processedLogs.add(logKey);

        } else {
          // --- New Person ---
          const guestLogEntry = logData[logDataIndex];
          const lastName = guestLogEntry[logLastNameIndex];
          const firstName = guestLogEntry[logFirstNameIndex];

          const capitalizedLastName = capitalizeName(lastName);
          const capitalizedFirstName = capitalizeName(firstName);

          const nextRow = pastoralData.nextBlankRow;

          pastoralSheet.getRange(nextRow, 3).setValue(capitalizedLastName);  // C
          pastoralSheet.getRange(nextRow, 4).setValue(capitalizedFirstName); // D

          // Insert recent date in E
          const recentCell = pastoralSheet.getRange(nextRow, pastoralRecentDateCol);
          recentCell.setValue(eventDate);
          recentCell.setHorizontalAlignment("center");
          recentCell.setVerticalAlignment("middle");

          // Insert notes in G (always set, even if blank)
          const notesCell = pastoralSheet.getRange(nextRow, pastoralNotesCol);
          notesCell.setValue(notes);
          notesCell.setHorizontalAlignment("left");
          notesCell.setVerticalAlignment("middle");

          // Insert extra in H (pastor, always set)
          const extraCell = pastoralSheet.getRange(nextRow, pastoralExtraCol);
          extraCell.setValue(extra);
          extraCell.setHorizontalAlignment("left");    // LEFT ALIGN

          SpreadsheetApp.flush();

          pastoralData.nameMap.set(name, nextRow);
          pastoralData.nextBlankRow++;
          pastoralData.numRows++;

          logData[logDataIndex][logStatusColIndex] = 'Logged';
          logData[logDataIndex][logRemarksColIndex] = 'New person added to directory.';
          recordsWereLogged = true;
          processedLogs.add(logKey);
        }

      } else {

        // --- CASE 3: Other Events (to 'Event Attendance' sheet) ---
        if (!eventSheetData) continue;

        const eventKey = `${formattedFullDate}_${eventName.trim().toLowerCase()}`;
        let col = eventSheetData.dateMap.get(eventKey);

        if (!col) {
          col = eventSheetData.lastDataCol + 1;
          
          eventSheet.getRange(eventDateRow, col).setValue(eventDate);
          eventSheet.getRange(eventNameRow, col).setValue(eventName);
          
          const colLetter = eventSheet.getRange(1, col).getA1Notation().replace(/\d+/g, '');
          const formula = `=COUNTIF(${colLetter}${eventDataStartRow}:${colLetter}, TRUE)`;
          eventSheet.getRange(eventCountRow, col).setFormula(formula);

          if (eventSheetData.numRows > 0) {
            const newCheckboxRange = eventSheet.getRange(eventDataStartRow, col, eventSheetData.numRows, 1);
            newCheckboxRange.insertCheckboxes();  
          }
          
          SpreadsheetApp.flush();

          eventSheetData.dateMap.set(eventKey, col);
          eventSheetData.lastDataCol = col;
          eventSheetData.checkboxes.forEach(r => r.push(false));
          Logger.log(`Added new column ${col} for ${eventName} on ${formattedFullDate} to sheet "${eventSheetName}".`);
        }

        let row = eventSheetData.nameMap.get(name);

        if (row) {
          const arrayRow = row - eventDataStartRow;  
          const arrayCol = col - eventDataStartCol;
          if (eventSheetData.checkboxes[arrayRow] && eventSheetData.checkboxes[arrayRow][arrayCol] !== undefined) {
            eventSheetData.checkboxes[arrayRow][arrayCol] = true;
            logData[logDataIndex][logStatusColIndex] = 'Logged';
            logData[logDataIndex][logRemarksColIndex] = '';
            recordsWereLogged = true;
            processedLogs.add(logKey);
          }
        } else {
          const guestLogEntry = logData[logDataIndex];
          const lastName = guestLogEntry[logLastNameIndex];
          const firstName = guestLogEntry[logFirstNameIndex];
          const type = guestLogEntry[logTypeIndex];
          
          const capitalizedLastName = capitalizeName(lastName);
          const capitalizedFirstName = capitalizeName(firstName);
          
          const nextRow = eventSheetData.nextBlankRow;
          
          eventSheet.getRange(nextRow, 3).setValue(capitalizedLastName); // Col C
          eventSheet.getRange(nextRow, 4).setValue(capitalizedFirstName); // Col D
          eventSheet.getRange(nextRow, typeColumnIndex).setValue(type); // Col F
          
          SpreadsheetApp.flush(); 

          eventSheetData.nameMap.set(name, nextRow);
          
          const numCols = eventSheetData.checkboxes[0] ? eventSheetData.checkboxes[0].length : 0;
          const newCheckboxRow = Array(numCols).fill(false);
          
          const arrayCol = col - eventDataStartCol;
          newCheckboxRow[arrayCol] = true;  
          
          eventSheetData.checkboxes.push(newCheckboxRow);
          eventSheetData.numRows++;
          eventSheetData.nextBlankRow++;
          
          logData[logDataIndex][logStatusColIndex] = 'Logged';
          logData[logDataIndex][logRemarksColIndex] = 'New person added to directory.';
          recordsWereLogged = true;
          processedLogs.add(logKey);
        }
      }
    } catch (e) {
      Logger.log(`Error processing record for ${name} at ${eventName}: ${e.message} ${e.stack}`);
    }
  }

  // 4. Write all updates to the sheets
  // Write to Sunday Service
  if (sunServiceData && sunServiceData.checkboxes.length > 0 && sunServiceData.checkboxes[0].length > 0) {
    const finalSunRange = sunServiceSheet.getRange(
      sunServiceDataStartRow,
      sunServiceDataStartCol,
      sunServiceData.checkboxes.length,
      sunServiceData.checkboxes[0].length
    );
    finalSunRange.setValues(sunServiceData.checkboxes);
    
    SpreadsheetApp.flush();
    Logger.log(`Successfully updated "${sunServiceSheetName}".`);
  }

  // Write to Event Attendance
  if (eventSheetData && eventSheetData.checkboxes.length > 0 && eventSheetData.checkboxes[0].length > 0) {
    const finalRange = eventSheet.getRange(
      eventDataStartRow,  
      eventDataStartCol,  
      eventSheetData.checkboxes.length,
      eventSheetData.checkboxes[0].length
    );
    finalRange.setValues(eventSheetData.checkboxes);
    
    SpreadsheetApp.flush();
    Logger.log(`Successfully updated "${eventSheetName}".`);
  }

  // --- REFINEMENT: Write back "Logged" status to 'Attendance Log' ---
  if (recordsWereLogged) {
    const statusData = logData.map(row => [row[logStatusColIndex], row[logRemarksColIndex]]);
    logSheet.getRange(2, logStatusColumn, statusData.length, 2).setValues(statusData);
    Logger.log('Updated "Logged" and "Guest" status in Attendance Log.');
  }
  // --- END REFINEMENT ---

  Logger.log('Attendance processing complete.');
}

/**
 * Helper function to read a sheet's layout and existing checkbox values.
 * FINDS THE *ACTUAL* LAST ROW with data in the name column.
 */
function prepareSheetData(sheet, dataStartRow, dataStartCol, nameColumnLetter, dateKeyRows, useFullDate, isEventSheet) {
  
  // --- REFINED: Find last row based on data in Col C or Col D ---
  const nameValuesAll = sheet.getRange(`C1:D${sheet.getMaxRows()}`).getValues();
  let actualLastDataRow = 0;
  for (let i = nameValuesAll.length - 1; i >= 0; i--) {
    const lastName = nameValuesAll[i][0]; // Col C
    const firstName = nameValuesAll[i][1]; // Col D
    if (lastName || firstName) {
      actualLastDataRow = i + 1;
      break;
    }
  }
  
  const nextBlankRow = actualLastDataRow < dataStartRow ? dataStartRow : actualLastDataRow + 1;
  const dataRowCount = actualLastDataRow >= dataStartRow ? (actualLastDataRow - dataStartRow + 1) : 0;

  // --- REFINED: Build Name-to-Row Map from Columns C and D ---
  const nameMap = new Map();
  if (dataRowCount > 0) {
    const namesData = nameValuesAll.slice(dataStartRow - 1, actualLastDataRow);
    
    for (let i = 0; i < namesData.length; i++) {
      const lastName = namesData[i][0];
      const firstName = namesData[i][1];
      
      if (lastName || firstName) {
        const name = standardizeNameHelper(`${lastName}, ${firstName}`);
        if (name && !nameMap.has(name)) {
          nameMap.set(name, i + dataStartRow);
        }
      }
    }
  } else {
     Logger.log(`Sheet "${sheet.getName()}" has no names in columns C or D starting at row ${dataStartRow}.`);
  }

  // 2. Build Date-to-Column Map
  const dateMap = new Map();
  const lastSheetCol = sheet.getLastColumn() || dataStartCol;
  
  const dateValues = sheet.getRange(dateKeyRows[0], 1, 1, lastSheetCol).getValues()[0];
  const nameValues = dateKeyRows[1] ? sheet.getRange(dateKeyRows[1], 1, 1, lastSheetCol).getValues()[0] : null;

  let lastDataCol = dataStartCol - 1;

  for (let i = dataStartCol - 1; i < lastSheetCol; i++) {
    const dateVal = dateValues[i];
    
    if (dateVal instanceof Date) {
      const formattedDate = useFullDate
        ? `${dateVal.getMonth() + 1}-${dateVal.getDate()}-${dateVal.getFullYear()}`
        : `${dateVal.getMonth() + 1}-${dateVal.getDate()}`;
        
      let key;
      if (nameValues) {
        const eventName = nameValues[i] ? nameValues[i].toString().trim().toLowerCase() : '';
        key = `${formattedDate}_${eventName}`;
      } else {
        key = formattedDate;
      }
      
      dateMap.set(key, i + 1);  
      lastDataCol = i + 1;

    } else if (dateVal === '' && (!nameValues || nameValues[i] === '')) {
      break;
    } else if (i >= dataStartCol -1) {
      lastDataCol = i + 1;
    }
  }

  // 3. Get all existing checkbox values
  const numCols = lastDataCol >= dataStartCol ? (lastDataCol - dataStartCol + 1) : 0;
  let checkboxes = [];
  let range;  

  if (dataRowCount > 0) {
    if (numCols > 0) {
      range = sheet.getRange(dataStartRow, dataStartCol, dataRowCount, numCols);
      checkboxes = range.getValues();
      if (isEventSheet) {
        range.insertCheckboxes();
      }
    } else {
      range = sheet.getRange(dataStartRow, dataStartCol, dataRowCount, 1);  
      checkboxes = Array(dataRowCount).fill(0).map(() => []);
    }
  }

  return {
    sheet: sheet,
    nameMap: nameMap,
    dateMap: dateMap,
    checkboxes: checkboxes,
    lastDataCol: lastDataCol,  
    numRows: dataRowCount,
    nextBlankRow: nextBlankRow
  };
}

/**
 * Helper function for Pastoral Check-In sheet.
 * Builds nameMap and nextBlankRow only.
 */
function preparePastoralSheetData(sheet, dataStartRow) {
  
  const nameValuesAll = sheet.getRange(`C1:D${sheet.getMaxRows()}`).getValues();
  let actualLastDataRow = 0;
  for (let i = nameValuesAll.length - 1; i >= 0; i--) {
    const lastName = nameValuesAll[i][0];
    const firstName = nameValuesAll[i][1];
    if (lastName || firstName) {
      actualLastDataRow = i + 1;
      break;
    }
  }

  const nextBlankRow = actualLastDataRow < dataStartRow ? dataStartRow : actualLastDataRow + 1;
  const dataRowCount = actualLastDataRow >= dataStartRow ? (actualLastDataRow - dataStartRow + 1) : 0;

  const nameMap = new Map();
  if (dataRowCount > 0) {
    const namesData = nameValuesAll.slice(dataStartRow - 1, actualLastDataRow);
    for (let i = 0; i < namesData.length; i++) {
      const lastName = namesData[i][0];
      const firstName = namesData[i][1];
      if (lastName || firstName) {
        const name = standardizeNameHelper(`${lastName}, ${firstName}`);
        if (name && !nameMap.has(name)) {
          nameMap.set(name, i + dataStartRow);
        }
      }
    }
  }

  return {
    sheet: sheet,
    nameMap: nameMap,
    numRows: dataRowCount,
    nextBlankRow: nextBlankRow
  };
}

/**
 * Standardizes a name into 'lastname,firstname' format (no spaces).
 * REFINED: This is the new, consistent, and smarter version.
 */
function standardizeNameHelper(nameStr) {
  if (!nameStr) return '';
  
  // 1. Clean junk (like (Guest), (Deacon), etc.), lowercase
  let name = nameStr.toString().trim().toLowerCase().replace(/[^a-z,\s]/g, "");

  let lastName = '';
  let firstName = '';

  if (name.includes(',')) {
    const parts = name.split(',');
    lastName = parts[0] ? parts[0].trim() : '';
    firstName = parts[1] ? parts[1].trim() : '';
  } else {
    const parts = name.split(/\s+/).filter(Boolean);
    if (parts.length === 0) return '';
    if (parts.length === 1) {
      firstName = parts[0] ? parts[0].trim() : '';
    } else {
      firstName = parts.slice(0, parts.length - 1).join(' ');
      lastName = parts[parts.length - 1].trim();
    }
  }
  
  lastName = lastName.replace(/\s+/g, '');
  firstName = firstName.replace(/\s+/g, '');

  return `${lastName},${firstName}`;
}

/**
* Capitalizes the first letter of each part of a name (e.g., "john doe" -> "John Doe").
 * Handles hyphenated names (e.g., "arai-joseph" -> "Arai-Joseph").
 */
function capitalizeName(nameStr) {
  if (!nameStr) return '';
  
  return nameStr.toLowerCase()
    .replace(/\b(\w)|(-(\w))/g, (match, p1, p2, p3) => {
      if (p1) return p1.toUpperCase();
      if (p3) return '-' + p3.toUpperCase();
      return match;
    });
}
