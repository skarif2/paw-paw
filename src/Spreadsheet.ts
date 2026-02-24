/** Matches sheet names in `yyyy-MM` format — used to identify dated roster sheets 🐾 */
const SHEET_DATE_PATTERN = /^\d{4}-\d{2}$/;

/** Valid attendance statuses for a standard working day */
const ATTENDANCE_OPTIONS = ['Office', 'WFH', 'Leave'] as const;

/** Valid attendance statuses for a team off-day (office is not an option) */
const OFFDAY_OPTIONS = ['WFH', 'Leave'] as const;

/** Row index of the hidden Slack ID row — stores immutable user IDs beside display names */
const ID_ROW = 2;

/** First row index of actual attendance data (below the hidden ID row) */
const DATA_START = 3;

/**
 * Extracts year, month, and days-in-month from a `yyyy-MM` named sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - A dated roster sheet
 * @returns {{ year: number, month: number, daysInMonth: number }}
 */
function getSheetInfo(sheet: GoogleAppsScript.Spreadsheet.Sheet): { year: number; month: number; daysInMonth: number } {
  const [year, month] = sheet.getName().split('-').map(Number);
  return { year, month, daysInMonth: new Date(year, month, 0).getDate() };
}

/**
 * Converts a 1-based column index to its Excel-style letter notation.
 * @param {number} column - 1-based column index
 * @returns {string} Column letter(s), e.g. 1 → `A`, 2 → `B`, 27 → `AA`
 */
function columnToLetter(column: number): string {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Writes the Office/WFH/Leave summary section below the data rows.
 * Sets labels, a top border, and COUNTIF formulas for each member column.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {number} daysInMonth - Number of data rows (used to anchor the section)
 * @param {number} memberCount - Number of member columns to fill
 * @param {number} totalCols - Total columns (used for the top border span)
 */
function writeSummarySection(sheet: GoogleAppsScript.Spreadsheet.Sheet, daysInMonth: number, memberCount: number, totalCols: number): void {
  // +1 for header, +1 for ID row, +1 for gap row, +1 for 1-based = DATA_START + daysInMonth + 1 🐾
  const summaryStartRow = DATA_START + daysInMonth + 1;
  sheet.getRange(summaryStartRow, 1, 3, 1).setValues([["Office"], ["WFH"], ["Leave"]]).setFontWeight("bold");
  sheet.getRange(summaryStartRow, 1, 1, totalCols).setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  for (let i = 0; i < memberCount; i++) {
    const colLetter = columnToLetter(i + 2);
    // Data starts at DATA_START and runs for daysInMonth rows 🐾
    const dataRange = `${colLetter}${DATA_START}:${colLetter}${DATA_START + daysInMonth - 1}`;
    sheet.getRange(summaryStartRow, i + 2).setFormula(`=COUNTIF(${dataRange}, "Office")`);
    sheet.getRange(summaryStartRow + 1, i + 2).setFormula(`=COUNTIF(${dataRange}, "WFH")`);
    sheet.getRange(summaryStartRow + 2, i + 2).setFormula(`=COUNTIF(${dataRange}, "Leave")`);
  }
}

/**
 * Applies conditional formatting rules and weekly border styling to a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {number} totalRows - Number of data rows (excluding header)
 * @param {number} totalCols - Total number of columns
 */
function addFormattingRules(sheet: GoogleAppsScript.Spreadsheet.Sheet, totalRows: number, totalCols: number): void {
  // Data rows start at DATA_START — skip the header and hidden ID row 🐾
  const dataRange = sheet.getRange(DATA_START, 1, totalRows, totalCols);
  const rules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] = [];

  // 1. Weekend Rule — grey out Saturday and Sunday rows
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=WEEKDAY($A${DATA_START}, 2) > 5`)
    .setBackground("#EFEFEF")
    .setFontColor("#999999")
    .setRanges([dataRange])
    .build());

  // 2. Holiday Rule — light green background for national holidays
  const { holidays, offdays } = getDateConfig();
  if (holidays.length > 0) {
    const holidayStrings = holidays.map(d => `"${d}"`).join(",");
    const holidayFormula = `=AND(ISNUMBER(MATCH(TEXT($A${DATA_START}, "yyyy-mm-dd"), {${holidayStrings}}, 0)), WEEKDAY($A${DATA_START}, 2) <= 5)`;

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(holidayFormula)
      .setBackground("#D9EAD3") // Light Green
      .setFontColor("#274E13") // Dark Green
      .setRanges([dataRange])
      .build());
  }

  // 3. Offday Rule — light orange background for team off-days
  if (offdays.length > 0) {
    const offdayStrings = offdays.map(d => `"${d}"`).join(",");
    const offdayFormula = `=AND(ISNUMBER(MATCH(TEXT($A${DATA_START}, "yyyy-mm-dd"), {${offdayStrings}}, 0)), WEEKDAY($A${DATA_START}, 2) <= 5)`;

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(offdayFormula)
      .setBackground("#FCE5CD") // Light Orange
      .setFontColor("#B45F06") // Dark Orange
      .setRanges([dataRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);

  // 4. Weekly Borders — draw a thick bottom border after each Sunday to visually group weeks
  const dateValues = sheet.getRange(DATA_START, 1, totalRows, 1).getValues();
  for (let i = 0; i < dateValues.length; i++) {
    const date = new Date(dateValues[i][0]);
    if (date.getDay() === 0) { // Sunday
      sheet.getRange(i + DATA_START, 1, 1, totalCols).setBorder(
        null, null, true, null, null, null,
        "#666666",
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    }
  }
}

/**
 * Creates the next month's sheet based on the latest existing dated sheet.
 * Populates it with dates, member columns, data validation, row protections,
 * column protections, summary section, and styling.
 *
 * @remarks
 * Fetches current Slack channel members to populate the roster.
 * If no dated sheet exists yet, defaults to the current calendar month.
 */
function createNextMonthSheet(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let latestDate: Date | null = null;
  for (const sh of sheets) {
    const name = sh.getName();
    if (SHEET_DATE_PATTERN.test(name)) {
      const [year, month] = name.split('-').map(Number);
      const sheetDate = new Date(year, month - 1, 1);
      if (!latestDate || sheetDate > latestDate) latestDate = sheetDate;
    }
  }

  const now = new Date();
  const targetDate = !latestDate
    ? new Date(now.getFullYear(), now.getMonth(), 1)
    : new Date(latestDate.getFullYear(), latestDate.getMonth() + 1, 1);

  const newSheetName = Utilities.formatDate(targetDate, CONFIG.TIMEZONE, "yyyy-MM");
  if (ss.getSheetByName(newSheetName)) return;

  const userList = Object.values(getChannelUsers());
  const teamMembers = userList.sort((a, b) => a.name.localeCompare(b.name));

  const sheet = ss.insertSheet(newSheetName);
  const headers = ["Date", ...teamMembers.map(m => m.name), "Total"];
  // ID row mirrors the header positions — "—" for non-member columns 🐾
  const idRow = ["—", ...teamMembers.map(m => m.id), "—"];
  const daysInMonth = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0).getDate();

  const rows: any[][] = [];
  const { holidays: holidayList, offdays: offdayList } = getDateConfig();

  // Pre-compute per-day metadata once — reused for both row generation and validation 🐾
  const lastMemberCol = columnToLetter(teamMembers.length + 1);
  const dayInfos = Array.from({ length: daysInMonth }, (_, idx) => {
    const d = idx + 1;
    const date = new Date(targetDate.getFullYear(), targetDate.getMonth(), d);
    const dateStr = Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy-MM-dd");
    return {
      date,
      dateStr,
      isWeekend: date.getDay() === 0 || date.getDay() === 6,
      isHoliday: holidayList.includes(dateStr),
      isOffday: offdayList.includes(dateStr),
    };
  });

  // 1. GENERATE ROWS — one row per calendar day (data begins at DATA_START)
  for (let d = 1; d <= daysInMonth; d++) {
    const { date: currentDate, isWeekend, isHoliday, isOffday } = dayInfos[d - 1];
    const rowData: any[] = [currentDate];
    const rowNum = d + DATA_START - 1; // absolute sheet row for this day

    if (isWeekend) {
      teamMembers.forEach(() => rowData.push("-"));
      rowData.push("-");
    } else {
      const defaultValue = (isHoliday || isOffday) ? "Leave" : "Office";
      teamMembers.forEach(() => rowData.push(defaultValue));
      rowData.push(`=COUNTIF(B${rowNum}:${lastMemberCol}${rowNum}, "Office")`);
    }
    rows.push(rowData);
  }

  // Write header (Row 1), hidden ID row (Row 2), and data rows (Row DATA_START onward)
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground('#a4c2f4').setWrap(true);
  sheet.getRange(ID_ROW, 1, 1, idRow.length).setValues([idRow]);
  sheet.hideRows(ID_ROW); // 😼 Out of sight — this row is for machines, not humans
  sheet.getRange(DATA_START, 1, rows.length, headers.length).setValues(rows);

  // 2. SELECTIVE DATA VALIDATION & ROW PROTECTIONS (weekends and holidays are locked)
  const offdayRule = SpreadsheetApp.newDataValidation().requireValueInList([...OFFDAY_OPTIONS], true).build();
  const standardRule = SpreadsheetApp.newDataValidation().requireValueInList([...ATTENDANCE_OPTIONS], true).build();

  for (let d = 1; d <= daysInMonth; d++) {
    const { dateStr, isWeekend, isHoliday, isOffday } = dayInfos[d - 1];
    const rowNum = d + DATA_START - 1;
    const rowRange = sheet.getRange(rowNum, 2, 1, teamMembers.length);

    if (isWeekend) {
      const p = sheet.getRange(rowNum, 1, 1, headers.length).protect().setDescription(`Weekend ${dateStr}`);
      p.removeEditors(p.getEditors());
    } else if (isHoliday || isOffday) {
      rowRange.setDataValidation(offdayRule);
    } else {
      rowRange.setDataValidation(standardRule);
    }

    if (isHoliday || isOffday) {
      const p = sheet.getRange(rowNum, 1, 1, headers.length).protect().setDescription(`Offday/Holiday ${dateStr}`);
      p.removeEditors(p.getEditors());
    }
  }

  // 3. INDIVIDUAL COLUMN PROTECTIONS — each member can only edit their own column
  // Description keyed by user ID — survives display name changes 😸
  for (let i = 0; i < teamMembers.length; i++) {
    const member = teamMembers[i];
    const colIndex = i + 2; // +1 for Date col, +1 for 1-based index

    const colRange = sheet.getRange(DATA_START, colIndex, daysInMonth, 1);
    const protection = colRange.protect().setDescription(`member:${member.id}`);

    // Remove everyone, then add only this member back as an editor
    protection.removeEditors(protection.getEditors());
    if (member.email) {
      protection.addEditor(member.email);
    }
  }

  // 4. STRUCTURAL PROTECTIONS — lock header, ID row, date column, and totals column
  const headerProt = sheet.getRange(1, 1, 1, headers.length).protect().setDescription("Headers");
  headerProt.removeEditors(headerProt.getEditors());

  const idRowProt = sheet.getRange(ID_ROW, 1, 1, idRow.length).protect().setDescription("Member IDs");
  idRowProt.removeEditors(idRowProt.getEditors());

  const dateProt = sheet.getRange(DATA_START, 1, daysInMonth, 1).protect().setDescription("Dates");
  dateProt.removeEditors(dateProt.getEditors());

  const totalsProt = sheet.getRange(DATA_START, headers.length, daysInMonth, 1).protect().setDescription("Totals");
  totalsProt.removeEditors(totalsProt.getEditors());

  // 5. SUMMARY SECTION — per-member Office/WFH/Leave counts below the data
  writeSummarySection(sheet, daysInMonth, teamMembers.length, headers.length);

  const summaryStartRow = DATA_START + daysInMonth + 1;
  const summaryProt = sheet.getRange(summaryStartRow, 1, 3, headers.length).protect().setDescription("Summary");
  summaryProt.removeEditors(summaryProt.getEditors());

  // 6. FINAL STYLING — number format, fonts, alignment, frozen panes
  sheet.getRange(DATA_START, 1, rows.length, 1).setNumberFormat("yyyy-mm-dd").setFontWeight("bold");
  addFormattingRules(sheet, daysInMonth, headers.length);

  const fullRange = sheet.getRange(1, 1, summaryStartRow + 2, headers.length);
  fullRange.setFontSize(11).setVerticalAlignment("middle").setHorizontalAlignment("center");
  sheet.setRowHeights(1, summaryStartRow + 2, 32);
  sheet.hideRows(ID_ROW); // Re-hide after bulk setRowHeights resets it 😼
  sheet.setColumnWidth(1, 100);
  if (headers.length > 2) sheet.setColumnWidths(2, headers.length - 1, 90);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

/**
 * Syncs the current month and any future month sheets with the live Slack roster.
 * Exits early (no-op) if no membership changes are detected.
 */
function syncAllActiveSheets(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const today = new Date();
  const currentMonthStr = Utilities.formatDate(today, CONFIG.TIMEZONE, "yyyy-MM");

  // 1. Fetch current Slack roster
  const userMap = getChannelUsers();
  const currentSlackUsers = Object.values(userMap).sort((a, b) => a.name.localeCompare(b.name));
  const currentSlackIds = currentSlackUsers.map(u => u.id);

  // 2. GATEKEEPER CHECK — compare ID row against live Slack IDs (immune to renames 😸)
  const currentSheet = ss.getSheetByName(currentMonthStr);
  if (currentSheet) {
    const idRowValues = currentSheet.getRange(ID_ROW, 1, 1, currentSheet.getLastColumn()).getValues()[0];
    // Exclude first ("—") and last ("—") sentinel values to get member IDs only
    const sheetMemberIds = idRowValues.slice(1, -1) as string[];

    const isSynced = sheetMemberIds.length === currentSlackIds.length &&
      sheetMemberIds.every(id => currentSlackIds.includes(id));

    if (isSynced) {
      // 2a. Even if membership is in sync, check for display name changes 🐾
      const nameHeaders = currentSheet.getRange(1, 1, 1, currentSheet.getLastColumn()).getValues()[0];
      const sheetMemberNames = nameHeaders.slice(1, -1) as string[];
      let nameUpdated = false;
      sheetMemberIds.forEach((id, idx) => {
        const liveUser = currentSlackUsers.find(u => u.id === id);
        if (liveUser && liveUser.name !== sheetMemberNames[idx]) {
          currentSheet.getRange(1, idx + 2).setValue(liveUser.name);
          nameUpdated = true;
        }
      });
      if (nameUpdated) {
        console.log("😸 Roster IDs unchanged — display name(s) silently updated.");
      } else {
        console.log("😸 Roster is purrfect — no sync needed.");
      }
      return;
    }
  }

  // 3. Roster has changed — process each active or future sheet
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (SHEET_DATE_PATTERN.test(sheetName)) {
      if (sheetName >= currentMonthStr) {
        const isFutureSheet = sheetName > currentMonthStr;
        processSheetSync(sheet, currentSlackUsers, currentSlackIds, isFutureSheet);
      }
    }
  });

  console.log("😸 Multi-month sync complete — Paw-Paw's work here is done.");
}

/**
 * Applies the member sync diff (leavers, joiners, and renames) to a single sheet.
 * Identity is established via the hidden Slack ID row — immune to display name changes. 😸
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to update
 * @param {SlackUser[]} currentSlackUsers - Full list of current Slack members
 * @param {string[]} currentSlackIds - Slack IDs of current members (for fast lookup)
 * @param {boolean} isFutureSheet - If true, leaver columns are deleted instead of zeroed out
 */
function processSheetSync(sheet: GoogleAppsScript.Spreadsheet.Sheet, currentSlackUsers: SlackUser[], currentSlackIds: string[], isFutureSheet: boolean): void {
  const today = new Date();
  const todayDay = today.getDate();
  const { year, month, daysInMonth } = getSheetInfo(sheet);
  const { holidays: holidayList, offdays: offdayList } = getDateConfig();

  // 0. PRE-CLEAN: Remove the existing summary area to prevent stale rows or gaps
  const lastRow = sheet.getMaxRows();
  const lastCol = sheet.getMaxColumns();
  // Summary starts at DATA_START + daysInMonth + 1 (i.e., the gap row is DATA_START + daysInMonth)
  const oldSummaryStart = DATA_START + daysInMonth;
  if (lastRow > oldSummaryStart) {
    sheet.getRange(oldSummaryStart, 1, lastRow - oldSummaryStart + 1, lastCol).clear().breakApart();
  }

  // Read IDs from the hidden row for identity comparison — names in row 1 are display-only 😸
  let idRowValues = sheet.getRange(ID_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  let sheetMemberIds = idRowValues.slice(1, -1) as string[];

  // --- 1. HANDLE LEAVERS (matched by Slack ID) ---
  const leaverIds = sheetMemberIds.filter(id => !currentSlackIds.includes(id));
  leaverIds.forEach(leaverId => {
    // Re-read the id row each iteration in case columns shifted after a deletion
    const freshIdRow = sheet.getRange(ID_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = freshIdRow.indexOf(leaverId) + 1;
    if (colIndex > 0) {
      if (isFutureSheet) {
        sheet.deleteColumn(colIndex);
      } else {
        const remainingDays = daysInMonth - todayDay;
        if (remainingDays > 0) {
          // Rows for remaining days: start at DATA_START + todayDay, run remainingDays rows
          const range = sheet.getRange(DATA_START + todayDay, colIndex, remainingDays, 1);
          range.setDataValidation(null); // Strip dropdown so hyphens are allowed
          range.setValues(Array(remainingDays).fill(["-"]));
        }
        // Strip column protection using the stable ID-based description
        const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        const userProt = protections.find(p => p.getDescription() === `member:${leaverId}`);
        if (userProt) userProt.removeEditors(userProt.getEditors());
      }
    }
  });

  // Refresh after potential column deletions
  idRowValues = sheet.getRange(ID_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheetMemberIds = idRowValues.slice(1, -1) as string[];

  // --- 1b. HANDLE RENAMES — same ID but display name has changed 🐾 ---
  const nameHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheetMemberIds.forEach((id, idx) => {
    const liveUser = currentSlackUsers.find(u => u.id === id);
    const currentName = nameHeaders[idx + 1] as string;
    if (liveUser && liveUser.name !== currentName) {
      sheet.getRange(1, idx + 2).setValue(liveUser.name);
      console.log(`😸 Name updated in ${sheet.getName()}: "${currentName}" → "${liveUser.name}"`);
    }
  });

  // --- 2. HANDLE JOINERS (matched by Slack ID) ---
  const newJoiners = currentSlackUsers.filter(u => !sheetMemberIds.includes(u.id));
  newJoiners.forEach(user => {
    const insertPos = sheet.getLastColumn(); // Insert before the Totals column
    sheet.insertColumnBefore(insertPos);
    // Write display name to Row 1 and immutable ID to the hidden Row 2
    sheet.getRange(1, insertPos).setValue(user.name).setFontWeight("bold");
    sheet.getRange(ID_ROW, insertPos).setValue(user.id);

    const dropdownRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([...ATTENDANCE_OPTIONS], true)
      .build();

    const offdayRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([...OFFDAY_OPTIONS], true)
      .build();

    // Pre-compute per-day metadata once — no need to loop twice over the same dates 🐾
    const joinerDayInfos = Array.from({ length: daysInMonth }, (_, idx) => {
      const d = idx + 1;
      const rowDate = new Date(year, month - 1, d);
      const dateStr = Utilities.formatDate(rowDate, CONFIG.TIMEZONE, "yyyy-MM-dd");
      return {
        isPast: !isFutureSheet && d < todayDay,
        isWeekend: rowDate.getDay() === 0 || rowDate.getDay() === 6,
        isHoliday: holidayList.includes(dateStr),
        isOffday: offdayList.includes(dateStr),
      };
    });

    // Build column values from pre-computed info
    const columnValues = joinerDayInfos.map(({ isPast, isWeekend, isHoliday }) =>
      isPast || isWeekend ? ["-"] : [isHoliday ? "Leave" : "Office"]
    );

    sheet.getRange(DATA_START, insertPos, daysInMonth, 1).setValues(columnValues);

    // Apply data validation using pre-computed info — weekends and past days get no dropdown
    for (let d = 1; d <= daysInMonth; d++) {
      const { isPast, isWeekend, isHoliday, isOffday } = joinerDayInfos[d - 1];
      const cell = sheet.getRange(d + DATA_START - 1, insertPos);
      if (isWeekend || isPast) {
        cell.setDataValidation(null);
      } else if (isHoliday || isOffday) {
        cell.setDataValidation(offdayRule);
      } else {
        cell.setDataValidation(dropdownRule);
      }
    }

    // Description keyed by user ID — survives future renames 😸
    const prot = sheet.getRange(DATA_START, insertPos, daysInMonth, 1).protect().setDescription(`member:${user.id}`);
    prot.removeEditors(prot.getEditors());
    if (user.email) prot.addEditor(user.email);
  });

  // --- 3. RE-CALCULATE TOTALS & SUMMARY ---
  const finalLastCol = sheet.getLastColumn();
  const memberColLetter = columnToLetter(finalLastCol - 1);
  const summaryStartRow = DATA_START + daysInMonth + 1;

  const totalFormulas: any[][] = [];
  for (let r = DATA_START; r < DATA_START + daysInMonth; r++) {
    totalFormulas.push([`=COUNTIF(B${r}:${memberColLetter}${r}, "Office")`]);
  }
  sheet.getRange(DATA_START, finalLastCol, daysInMonth, 1).setFormulas(totalFormulas);

  const totalMembers = finalLastCol - 2;
  writeSummarySection(sheet, daysInMonth, totalMembers, finalLastCol);

  // --- 4. FINAL STYLING & RULES ---
  sheet.getRange(1, 1, 1, finalLastCol).setBackground('#a4c2f4').setHorizontalAlignment("center");
  sheet.getRange(summaryStartRow, 2, 3, totalMembers).setHorizontalAlignment("center");
  sheet.hideRows(ID_ROW); // Re-hide after row height resets — machines only 😼

  addFormattingRules(sheet, daysInMonth, finalLastCol);

  console.log(`🐾 Sheet ${sheet.getName()} processed and looking sharp.`);
}

/**
 * Applies holiday/offday cell values, data validation, protections, and
 * conditional formatting to a single sheet based on the current CONFIG.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to refresh
 * @param {boolean} isFutureSheet - If true, all days are refreshed (day 1 onward); otherwise only today onward
 */
function refreshSheetHolidayFormatting(sheet: GoogleAppsScript.Spreadsheet.Sheet, isFutureSheet: boolean): void {
  const today = new Date();
  const todayDay = today.getDate();
  const { year, month, daysInMonth } = getSheetInfo(sheet);

  const { holidays: holidayList, offdays: offdayList } = getDateConfig();

  const offdayRule = SpreadsheetApp.newDataValidation().requireValueInList([...OFFDAY_OPTIONS], true).build();
  const standardRule = SpreadsheetApp.newDataValidation().requireValueInList([...ATTENDANCE_OPTIONS], true).build();

  // Fetch protections and column count once — cheaper than per-row API calls 🐾
  const allProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const totalCols = sheet.getLastColumn();
  const memberCols = totalCols - 2; // exclude Date and Total columns

  // 1. UPDATE VALUES & LOCKS (Future Sheets Only) 🐾
  // For the current month, we only update colors to avoid overwriting manual attendance.
  if (isFutureSheet) {
    for (let d = 1; d <= daysInMonth; d++) {
      // rowNum maps calendar day d to its absolute sheet row (data starts at DATA_START) 🐾
      const rowNum = d + DATA_START - 1;
      const date = new Date(year, month - 1, d);
      const dateStr = Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy-MM-dd");

      const isWeekend = date.getDay() === 0 || date.getDay() === 6;
      const isHoliday = holidayList.includes(dateStr);
      const isOffday = offdayList.includes(dateStr);

      if (isWeekend) continue;

      const memberRange = sheet.getRange(rowNum, 2, 1, memberCols);

      // A. Update Cell Values — force 'Leave' for holidays/offdays, restore 'Office' otherwise
      const currentValues = memberRange.getValues()[0];
      const newValues = currentValues.map((val: string) => {
        if (val === '-') return val;
        if (isHoliday || isOffday) return 'Leave';
        return (val === 'Leave') ? 'Office' : val; // restore if it was Leave, preserve WFH/Office otherwise
      });
      memberRange.setValues([newValues]);

      // B. Update Data Validation — consistent dropdowns for all team off-days
      memberRange.setDataValidation(isHoliday || isOffday ? offdayRule : standardRule);

      // C. Update Row Protections — remove stale lock first, then re-add if needed
      const rowDesc = `Offday/Holiday ${dateStr}`;
      allProtections.forEach(p => {
        if (p.getDescription() === rowDesc) p.remove();
      });

      if (isHoliday || isOffday) {
        const p = sheet.getRange(rowNum, 1, 1, totalCols).protect().setDescription(rowDesc);
        p.removeEditors(p.getEditors());
      }
    }
  }

  // 4. RE-APPLY CONDITIONAL FORMATTING — reads fresh HOLIDAYS/OFFDAYS from CONFIG 🐾
  addFormattingRules(sheet, daysInMonth, totalCols);
  console.log(`🐾 Sheet ${sheet.getName()} refreshed — holidays are purr-fectly up to date.`);
}

/**
 * Refreshes holiday/offday formatting, cell values, data validation, and row
 * protections across all active sheets (current month onward). Run this manually
 * after updating the `Holidays` sheet tab mid-month. 🐾
 *
 * @remarks
 * - Current month: only rows from **today onward** are touched (past data preserved).
 * - Future month sheets: all rows are refreshed from day 1.
 * - Weekend rows are always skipped — the weekend rule owns those.
 */
function refreshHolidayFormatting(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const currentMonthStr = Utilities.formatDate(today, CONFIG.TIMEZONE, "yyyy-MM");

  let refreshed = 0;
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (SHEET_DATE_PATTERN.test(name) && name >= currentMonthStr) {
      refreshSheetHolidayFormatting(sheet, name > currentMonthStr);
      refreshed++;
    }
  });

  if (refreshed === 0) {
    console.warn(`🙀 No active sheets found from ${currentMonthStr} onward — nothing to refresh!`);
  } else {
    console.log(`😸 Holiday refresh purrfect — ${refreshed} sheet(s) all tidied up.`);
  }
}

/**
 * Creates the next month's sheet if today is on or after the 25th.
 * Called nightly by `sendTomorrowHeadcount` as a monthly admin check.
 */
function checkAndCreateFutureSheet(): void {
  const today = new Date();
  if (today.getDate() >= 25) {
    createNextMonthSheet();
  }
}
