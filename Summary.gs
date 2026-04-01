/**
 * GOOGLE APPS SCRIPT: EMPLOYEE REPORT GENERATOR
 * Cycle: 16th (Prev Month) to 15th (Selected Month)
 */

// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Employee Reports')
    .addItem('Create Employee Spreadsheet...', 'showCreateDialog')
    .addToUi();
}

// ============================================================
//  SHOW DIALOG
// ============================================================
function showCreateDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(460)
    .setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Create Employee Spreadsheet');
}

// ============================================================
//  CALLED FROM DIALOG
// ============================================================
function runCreateScript(month, year) {
  try {
    const result = createEmployeeSeparatedSpreadsheet(month, year);
    return { success: true, url: result.url, name: result.name, count: result.count };
  } catch (e) {
    console.error(e.stack);
    return { success: false, error: e.toString() };
  }
}

// ============================================================
//  CORE FUNCTION
// ============================================================
function createEmployeeSeparatedSpreadsheet(month, year) {
  // 1. Setup Parameters
  if (!month || !year) {
    const now = new Date();
    month = now.getMonth() + 1;
    year  = now.getFullYear();
  }

  const SOURCE_SHEET_NAME = "CONS_DATA";
  const EMPLOYEE_COL      = 0;  // Column A
  const DATE_COL          = 1;  // Column B

  // 2. Define Cycle (16th Prev Month to 15th Current Month)
  const startDate = new Date(year, month - 2, 16, 0, 0, 0, 0);
  const endDate   = new Date(year, month - 1, 15, 23, 59, 59, 999);
  
  const startTs = startDate.getTime();
  const endTs   = endDate.getTime();
  const monthLabel = getMonthLabel(month, year);

  // 3. Get Source Data
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) throw new Error(`Sheet "${SOURCE_SHEET_NAME}" not found.`);

  const allData = sourceSheet.getDataRange().getValues();
  if (allData.length < 2) throw new Error("No data found in CONS_DATA.");

  const headers  = allData[0];
  const dataRows = allData.slice(1);

  // 4. Filter Rows
  const filteredRows = dataRows.filter(row => {
    const cell = row[DATE_COL];
    if (cell === "" || cell === null || cell === undefined) return false;

    let d;
    if (cell instanceof Date) {
      d = new Date(cell.getTime()); 
    } else if (typeof cell === 'number') {
      d = new Date(Math.round((cell - 25569) * 86400 * 1000));
    } else {
      d = parseDateDDMMMYY(String(cell).trim());
    }

    if (!d || isNaN(d.getTime())) return false;

    d.setHours(0, 0, 0, 0);
    const currentTs = d.getTime();

    return currentTs >= startTs && currentTs <= endTs;
  });

  // 5. Group by Employee
  const employeeMap = {};
  filteredRows.forEach(row => {
    const empName = String(row[EMPLOYEE_COL]).trim();
    if (!empName) return;
    if (!employeeMap[empName]) employeeMap[empName] = [];
    employeeMap[empName].push(row);
  });

  const employeeNames = Object.keys(employeeMap).sort();
  if (employeeNames.length === 0) {
    throw new Error(`No data found between ${startDate.toDateString()} and ${endDate.toDateString()}.`);
  }

  // 6. Create New Spreadsheet
  const newName = `Employee Time Records — ${monthLabel}`;
  const newSS   = SpreadsheetApp.create(newName);
  let isFirst   = true;

  employeeNames.forEach(empName => {
    let sheet;
    if (isFirst) {
      sheet = newSS.getSheets()[0];
      sheet.setName(sanitizeName(empName));
      isFirst = false;
    } else {
      sheet = newSS.insertSheet(sanitizeName(empName));
    }

    const rows = [headers, ...employeeMap[empName]];
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  });

  // 7. Summary Tab (Using your OT Logic)
  buildOTSummaryTab(newSS, employeeNames, monthLabel, startDate, endDate);

  return { url: newSS.getUrl(), name: newName, count: employeeNames.length };
}

// ============================================================
//  OT SUMMARY LOGIC (Replaces buildSummaryTab)
// ============================================================
function buildOTSummaryTab(ss, employeeNames, monthLabel, startDate, endDate) {
  const summarySheet = ss.getSheetByName("📊 Summary") || ss.insertSheet("📊 Summary", 0);
  summarySheet.clear(); 

  const tz = Session.getScriptTimeZone();
  let totalAllOT = 0;
  let summaryRows = [];

  // FIX: Ensure startDate and endDate are Date objects for Utilities.formatDate
  const sDate = new Date(startDate);
  const eDate = new Date(endDate);

  summarySheet.getRange("A1").setValue(`OT Summary Report — ${monthLabel}`)
    .setFontSize(14).setFontWeight("bold").setFontColor("#1565C0");
  
  summarySheet.getRange("A2").setValue(
    `Cycle: ${Utilities.formatDate(sDate, tz, "dd MMM yyyy")} to ${Utilities.formatDate(eDate, tz, "dd MMM yyyy")}`
  ).setFontColor("#5f6368");

  const headers = ["Employee Name", "Total Overtime", "WeekDay Overtime (Subtotal)", "Weekend Overtime (Subtotal)", "Weekend Double Overtime (Subtotal)", "Pubic Holiday Overtime (Subtotal)", "Pubic Holiday Double Overtime (Subtotal)", "Overtime-Leave Compensate"];
  const headersforeachname = ["Weekday OT", "OT 1.5 (Hrs)", "Weekend OT (8hrs.)", "OT 1 (Hrs)", "Weekend OT", "OT 3 (Hrs)"];
  summarySheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  styleHeader(summarySheet, headers.length);

  employeeNames.forEach(name => {
    const sh = ss.getSheetByName(sanitizeName(name));
    if (!sh) return;

    // VBA Range B12:Q42 translated to GAS
    sh.getRange(1, 12, 1, headersforeachname.length).setValues([headersforeachname]);
    const range = sh.getRange(2, 2, 31, 16);
    const data = range.getValues();
    
    let stats = { ot1: 0, w1: 0, w3: 0, h1: 0, h2: 0, comp: 0, total: 0 };

    const updatedData = data.map(row => {
      let hrs = row[7];       // Col I (Hours)
      let type = row[8];      // Col J (Type)
      let cal = String(row[2]).trim(); // Col D (Calendar)
      
      let r1 = 0, r2 = 0, r3 = 0;

      if (hrs && type) {
        if (hrs > 8) hrs -= 1; // Lunch Deduction

        if (type !== "Overtime-Leave Compensate") {
          if (cal === "Weekday" && hrs > 8) {
            r1 = Number(((hrs - 8) * 1.5).toFixed(2));
            stats.ot1 += r1;
          } 
          else if (cal === "Weekend" || cal === "Holiday") {
            let base = hrs > 8 ? 8 : hrs;
            r2 = base; 
            if (cal === "Weekend") stats.w1 += r2; else stats.h1 += r2;

            if (hrs > 8) {
              r3 = Number(((hrs - 8) * 3).toFixed(2));
              if (cal === "Weekend") stats.w3 += r3; else stats.h2 += r3;
            }
          }
        } else {
          stats.comp += hrs;
        }
      }
      // Fill the calculated columns (L to Q)
      row[10] = r1 > 0 ? (r1 / 1.5).toFixed(2) : 0; row[11] = r1; 
      row[12] = r2 > 0 ? r2 : 0;                     row[13] = r2;
      row[14] = r3 > 0 ? (r3 / 3).toFixed(2) : 0;   row[15] = r3;
      return row;
    });

    range.setValues(updatedData); 
    
    stats.total = stats.ot1 + stats.w1 + stats.w3 + stats.h1 + stats.h2;
    totalAllOT += stats.total;

    summaryRows.push([name, stats.total, stats.ot1, stats.w1, stats.w3, stats.h1, stats.h2, stats.comp]);
  });

  if (summaryRows.length > 0) {
    summarySheet.getRange(5, 1, summaryRows.length, 8).setValues(summaryRows);
    
    const totalRow = 5 + summaryRows.length;
    summarySheet.getRange(totalRow, 1).setValue("GRAND TOTAL").setFontWeight("bold");
    summarySheet.getRange(totalRow, 2).setValue(totalAllOT).setFontWeight("bold")
      .setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.DOUBLE);
  }

  summarySheet.autoResizeColumns(1, 8);
}

// ============================================================
//  DATE PARSER
// ============================================================
function parseDateDDMMMYY(str) {
  const MONTHS = {
    jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
    jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
  };
  const match = str.match(/^(\d{1,2})[\-\/\s\.]?([A-Za-z]{3})[\-\/\s\.]?(\d{2,4})$/);
  if (!match) return null;

  const day   = parseInt(match[1], 10);
  const month = MONTHS[match[2].toLowerCase()];
  let   year  = parseInt(match[3], 10);
  if (month === undefined) return null;
  if (year < 100) year += 2000; 

  return new Date(year, month, day);
}

// ============================================================
//  HELPERS
// ============================================================
function getMonthLabel(month, year) {
  const M = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return M[month - 1] + String(year).slice(2);
}

function styleHeader(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
    .setBackground("#1565C0")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function sanitizeName(name) {
  return name.replace(/[\\\/\?\*\[\]\:\']/g, "").substring(0, 100).trim();
}
