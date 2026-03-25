/**
 * Workers' Comp Data Cleaner
 *
 * Reads the "raw" sheet, consolidates to one row per employee (by File #),
 * summing Reg Hours and O/T Hours, and writes the result to a "clean" sheet.
 *
 * Col J:  Co Code + "0" + File # (e.g. JA2010366)
 * Col K:  XLOOKUP → 'Personnel File'!A:A → H:H (SSN)
 * Col L:  Hire Date from Contracts (start date of first active contract)
 * Col M:  State from Contracts (first active contract for Config!B1)
 * Col N:  Hourly Wage from Contracts
 * Col O:  WC Code from tbl_WC_rates
 * Col P:  WC Rate from tbl_WC_rates
 * Col Q:  WC Value = Hourly Wage * WC Rate * (Reg + OT Hours)
 *
 * Raw sheet columns:
 * A: Co Code | B: Batch ID | C: File # | D: Tax Frequency |
 * E: Temp Dept | F: Temp Rate | G: Reg Hours | H: O/T Hours | I: Employee Name
 *
 * Config sheet:  B1 = Report Date
 * Contracts sheet: A=Co Code+File# | B=Employee ID | D=State | E=State | F=Start | G=End | J=Wage
 */

type CellValue = string | number | boolean | Date;

const SHARED_FOLDER_ID = "1aOLGi-izMEUo-IvPpwugFotG8BIl6AN2";

// ─── Menu ─────────────────────────────────────────────────────────────────────

function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu("Workers' Comp")
    .addItem("1. Run Clean Report", "cleanWorkersCompData")
    .addItem("2. Build Data Needs", "buildDataNeeds")
    .addItem("3. Build Final Report", "buildFinalReport")
    .addItem("4. Publish Final Report", "publishFinalReport")
    .addToUi();
}

// ─── Step 1: Clean ────────────────────────────────────────────────────────────

interface EmployeeRecord {
  coCode: CellValue;
  batchId: CellValue;
  fileNum: CellValue;
  taxFreq: CellValue;
  tempDept: CellValue;
  tempRate: CellValue;
  regHours: number;
  otHours: number;
  empName: CellValue;
}

function cleanWorkersCompData(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const rawSheet = ss.getSheetByName("raw");
  if (!rawSheet) {
    ui.alert(
      'No sheet named "raw" found. Please rename your source sheet to "raw" and try again.',
    );
    return;
  }

  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    ui.alert(
      'No sheet named "Config" found. Expected the report date in Config!B1.',
    );
    return;
  }

  const contractsSheet = ss.getSheetByName("Contracts");
  if (!contractsSheet) {
    ui.alert('No sheet named "Contracts" found. State column will be blank.');
  }

  let cleanSheet = ss.getSheetByName("clean");
  if (cleanSheet) {
    const filter = cleanSheet.getFilter();
    if (filter) filter.remove();
    cleanSheet.clearContents();
  } else {
    cleanSheet = ss.insertSheet("clean");
  }

  const reportDate = configSheet.getRange("B1").getValue() as Date;
  if (!(reportDate instanceof Date) || isNaN(reportDate.getTime())) {
    ui.alert(
      "Config!B1 does not contain a valid date. Please enter the report date there and try again.",
    );
    return;
  }

  const rawData = rawSheet.getDataRange().getValues();

  const COL = {
    coCode: 0,
    batchId: 1,
    fileNum: 2,
    taxFreq: 3,
    tempDept: 4,
    tempRate: 5,
    regHours: 6,
    otHours: 7,
    empName: 8,
  };

  const employeeMap = new Map<string, EmployeeRecord>();

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const fileNum = row[COL.fileNum];
    if (!fileNum && fileNum !== 0) continue;

    const key = String(fileNum);

    if (!employeeMap.has(key)) {
      employeeMap.set(key, {
        coCode: row[COL.coCode],
        batchId: row[COL.batchId],
        fileNum: fileNum,
        taxFreq: row[COL.taxFreq],
        tempDept: row[COL.tempDept],
        tempRate: row[COL.tempRate],
        regHours: 0,
        otHours: 0,
        empName: row[COL.empName],
      });
    }

    const emp = employeeMap.get(key)!;
    emp.regHours +=
      typeof row[COL.regHours] === "number" ? (row[COL.regHours] as number) : 0;
    emp.otHours +=
      typeof row[COL.otHours] === "number" ? (row[COL.otHours] as number) : 0;

    if (!emp.taxFreq && row[COL.taxFreq]) emp.taxFreq = row[COL.taxFreq];
    if (!emp.tempDept && row[COL.tempDept]) emp.tempDept = row[COL.tempDept];
    if (!emp.tempRate && row[COL.tempRate]) emp.tempRate = row[COL.tempRate];
    if (!emp.empName && row[COL.empName]) emp.empName = row[COL.empName];
  }

  const outputHeaders = [
    "Co Code",
    "Batch ID",
    "File #",
    "Tax Frequency",
    "Temp Dept",
    "Temp Rate",
    "Reg Hours",
    "O/T Hours",
    "Employee Name",
    "Co Code + File #",
    "SSN",
    "Hire Date",
    "State",
    "Hourly Wage",
    "WC Code",
    "WC Rate",
    "WC Value",
    "Birthday",
    "Needs Review",
  ];
  const outputRows: CellValue[][] = [outputHeaders];

  employeeMap.forEach((emp) => {
    const concatKey =
      String(emp.coCode) + "0" + String(emp.fileNum).padStart(5, "0");
    outputRows.push([
      emp.coCode,
      emp.batchId,
      emp.fileNum,
      emp.taxFreq,
      emp.tempDept,
      emp.tempRate,
      emp.regHours,
      emp.otHours,
      emp.empName,
      concatKey,
      "", // K — SSN
      "", // L — Hire Date
      "", // M — State
      "", // N — Hourly Wage
      "", // O — WC Code
      "", // P — WC Rate
      "", // Q — WC Value
      "", // R — Birthday
      "", // S — Needs Review
    ]);
  });

  cleanSheet
    .getRange(1, 1, outputRows.length, outputHeaders.length)
    .setValues(outputRows);

  const dataRowCount = outputRows.length - 1;

  if (dataRowCount > 0) {
    const ssnFormulas: string[][] = [];
    const hireDateFormulas: string[][] = [];
    const stateFormulas: string[][] = [];
    const hourlyWageFormulas: string[][] = [];
    const wcCodeFormulas: string[][] = [];
    const wcRateFormulas: string[][] = [];
    const wcValueFormulas: string[][] = [];
    const birthdayFormulas: string[][] = [];
    const needsReviewFormulas: string[][] = [];

    for (let r = 2; r <= dataRowCount + 1; r++) {
      ssnFormulas.push([
        `=IFERROR(LET(v,XLOOKUP($J${r},'Personnel File'!$A:$A,'Personnel File'!$G:$G),IF(v="","NOT FOUND",v)),"NOT FOUND")`,
      ]);
      hireDateFormulas.push([
        `=IFERROR(QUERY(Contracts!$A$2:$G,"SELECT F WHERE B = '"&$J${r}&"' AND F <= date '"&TEXT(Config!$B$2,"yyyy-mm-dd")&"' AND G >= date '"&TEXT(Config!$B$1,"yyyy-mm-dd")&"' LIMIT 1",0),"NOT FOUND")`,
      ]);
      stateFormulas.push([
        `=IFERROR(QUERY(Contracts!$A$2:$G,"SELECT E WHERE B = '"&$J${r}&"' AND F <= date '"&TEXT(Config!$B$2,"yyyy-mm-dd")&"' AND G >= date '"&TEXT(Config!$B$1,"yyyy-mm-dd")&"' LIMIT 1",0),"NOT FOUND")`,
      ]);
      hourlyWageFormulas.push([
        `=IFERROR(QUERY(Contracts!$A$2:$J,"SELECT J WHERE B = '"&$J${r}&"' AND F <= date '"&TEXT(Config!$B$2,"yyyy-mm-dd")&"' AND G >= date '"&TEXT(Config!$B$1,"yyyy-mm-dd")&"' LIMIT 1",0),"NOT FOUND")`,
      ]);
      wcCodeFormulas.push([
        `=IFERROR(LET(v,XLOOKUP($M${r},'WC Rates'!$A$2:$A$31,'WC Rates'!$B$2:$B$31),IF(v="","NOT FOUND",v)),"NOT FOUND")`,
      ]);
      wcRateFormulas.push([
        `=IFERROR(LET(v,XLOOKUP($M${r},'WC Rates'!$A$2:$A$31,'WC Rates'!$C$2:$C$31),IF(v="","NOT FOUND",v)),"NOT FOUND")`,
      ]);
      wcValueFormulas.push([
        `=IFERROR($N${r}*$P${r}*($G${r}+$H${r}),"NOT FOUND")`,
      ]);
      birthdayFormulas.push([
        `=IFERROR(LET(v,XLOOKUP($J${r},'Personnel File'!$A:$A,'Personnel File'!$H:$H),IF(v="","NOT FOUND",v)),"NOT FOUND")`,
      ]);
      needsReviewFormulas.push([
        `=IF(OR($K${r}="NOT FOUND",$L${r}="NOT FOUND",$M${r}="NOT FOUND",$N${r}="NOT FOUND",$R${r}="NOT FOUND"),"YES","")`,
      ]);
    }

    cleanSheet.getRange(2, 11, dataRowCount, 1).setFormulas(ssnFormulas);
    cleanSheet.getRange(2, 12, dataRowCount, 1).setFormulas(hireDateFormulas);
    cleanSheet.getRange(2, 13, dataRowCount, 1).setFormulas(stateFormulas);
    cleanSheet.getRange(2, 14, dataRowCount, 1).setFormulas(hourlyWageFormulas);
    cleanSheet.getRange(2, 15, dataRowCount, 1).setFormulas(wcCodeFormulas);
    cleanSheet.getRange(2, 16, dataRowCount, 1).setFormulas(wcRateFormulas);
    cleanSheet.getRange(2, 17, dataRowCount, 1).setFormulas(wcValueFormulas);
    cleanSheet.getRange(2, 18, dataRowCount, 1).setFormulas(birthdayFormulas);
    cleanSheet
      .getRange(2, 19, dataRowCount, 1)
      .setFormulas(needsReviewFormulas);

    // Conditional formatting: highlight "Needs Review" rows in red
    const reviewRange = cleanSheet.getRange(2, 19, dataRowCount, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("YES")
      .setBackground("#FFCDD2")
      .setFontColor("#B71C1C")
      .setRanges([reviewRange])
      .build();
    const existingRules = cleanSheet
      .getConditionalFormatRules()
      .filter((r) => !r.getRanges().some((rng) => rng.getColumn() === 19));
    cleanSheet.setConditionalFormatRules([...existingRules, rule]);
  }

  // --- Formatting ---
  const headerRange = cleanSheet.getRange(1, 1, 1, outputHeaders.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4A90D9");
  headerRange.setFontColor("#FFFFFF");

  if (dataRowCount > 0) {
    cleanSheet.getRange(2, 7, dataRowCount, 2).setNumberFormat("0.00"); // Reg/OT Hours
    cleanSheet.getRange(2, 12, dataRowCount, 1).setNumberFormat("MM/DD/YYYY"); // Hire Date
    cleanSheet.getRange(2, 14, dataRowCount, 1).setNumberFormat("$#,##0.00"); // Hourly Wage
    cleanSheet.getRange(2, 16, dataRowCount, 1).setNumberFormat("0.000%"); // WC Rate
    cleanSheet.getRange(2, 17, dataRowCount, 1).setNumberFormat("$#,##0.00"); // WC Value
    cleanSheet.getRange(2, 18, dataRowCount, 1).setNumberFormat("MM/DD/YYYY"); // Birthday
  }

  // Flush to ensure formulas are resolved before auto-fitting
  SpreadsheetApp.flush();

  // Auto-fit all columns
  cleanSheet.autoResizeColumns(1, outputHeaders.length);

  // Apply filter to the full data range
  cleanSheet.getDataRange().createFilter();

  ss.toast(
    `Done! ${rawData.length - 1} raw rows → ${dataRowCount} employees. Report date: ${reportDate.toLocaleDateString()}`,
    "Workers' Comp Cleaner",
    6,
  );
}

// ─── Step 2: Final ────────────────────────────────────────────────────────────

// Final sheet column layout (1-indexed):
//   A(1)  Name          B(2)  SS#           C(3)  Employee Hire Date
//   D(4)  WKC Code      E(5)  Check Week    F(6)  Hourly Wage
//   G(7)  DOB           H(8)  Reg Hours     I(9)  OT Hours
//   J(10) Hours Worked  K(11) OT Pay        L(12) Total Pay
//   M(13) WC Rate       N(14) WC Value
//
// Columns J, K, L, N are Sheets formulas so the math is fully auditable.
// Subtotal rows carry =SUM(N{start}:N{end}) rather than a hardcoded value.

interface DataRow {
  wcCode: string;
  state: string;
  empName: string;
  staticValues: CellValue[]; // 14 values; positions 10,11,12,14 are "" (formula slots)
}

function buildFinalReport(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const cleanSheet = ss.getSheetByName("clean");
  if (!cleanSheet) {
    ui.alert('"clean" sheet not found. Please run the Clean Report first.');
    return;
  }

  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    ui.alert('"Config" sheet not found. Expected report date in Config!B1.');
    return;
  }

  let finalSheet = ss.getSheetByName("Final");
  if (finalSheet) {
    finalSheet.clearContents();
    finalSheet.clearFormats();
  } else {
    finalSheet = ss.insertSheet("Final");
  }

  const checkWeek = configSheet.getRange("B1").getValue() as Date;
  if (!(checkWeek instanceof Date) || isNaN(checkWeek.getTime())) {
    ui.alert("Config!B1 does not contain a valid date.");
    return;
  }

  // --- Read clean sheet display values (formulas already calculated) ---
  const cleanData = cleanSheet.getDataRange().getDisplayValues();

  const C = {
    empName:    8,  // I
    concatKey:  9,  // J
    ssn:        10, // K
    hireDate:   11, // L
    state:      12, // M
    hourlyWage: 13, // N
    wcCode:     14, // O
    wcRate:     15, // P
    regHours:   6,  // G
    otHours:    7,  // H
    birthday:   17, // R
  };

  // --- Collect rows, skipping any with missing key data ---
  const dataRows: DataRow[] = [];
  for (let i = 1; i < cleanData.length; i++) {
    const row = cleanData[i];
    const empName    = row[C.empName];
    const ssn        = row[C.ssn];
    const hireDate   = row[C.hireDate];
    const wcCode     = row[C.wcCode];
    const state      = row[C.state];
    const hourlyWageStr = row[C.hourlyWage];
    const wcRateStr  = row[C.wcRate];
    const dob        = row[C.birthday];

    // Skip rows with any missing key field — they belong in Data Needs
    const isMissing = (v: string) => v === "NOT FOUND" || v === "";
    if (
      isMissing(ssn) ||
      isMissing(hireDate) ||
      isMissing(wcCode) ||
      isMissing(state) ||
      isMissing(hourlyWageStr) ||
      isMissing(dob)
    ) continue;

    const hourlyWage = parseFloat(hourlyWageStr.replace(/[$,]/g, "")) || 0;
    // WC Rate stored as decimal so formula arithmetic works (display format restores %)
    const wcRate     = parseFloat(wcRateStr.replace(/%/g, "")) / 100 || 0;
    const regHours   = parseFloat(row[C.regHours]) || 0;
    const otHours    = parseFloat(row[C.otHours])  || 0;

    dataRows.push({
      wcCode,
      state,
      empName,
      staticValues: [
        empName,    // A(1)  Name
        ssn,        // B(2)  SS#
        hireDate,   // C(3)  Employee Hire Date
        wcCode,     // D(4)  WKC Code
        checkWeek,  // E(5)  Check Week
        hourlyWage, // F(6)  Hourly Wage
        dob,        // G(7)  DOB
        regHours,   // H(8)  Reg Hours
        otHours,    // I(9)  OT Hours
        "",         // J(10) Hours Worked  ← formula
        "",         // K(11) OT Pay        ← formula
        "",         // L(12) Total Pay     ← formula
        wcRate,     // M(13) WC Rate
        "",         // N(14) WC Value      ← formula
      ],
    });
  }

  // Sort by WC Code, then Employee Name within each group
  dataRows.sort((a, b) => {
    const cc = String(a.wcCode).localeCompare(String(b.wcCode));
    return cc !== 0 ? cc : String(a.empName).localeCompare(String(b.empName));
  });

  // --- Build output rows, tracking sheet row numbers for formula insertion ---
  const finalHeaders = [
    "Name",               // A col 1
    "SS#",                // B col 2
    "Employee Hire Date", // C col 3
    "WKC Code",           // D col 4
    "Check Week",         // E col 5
    "Hourly Wage",        // F col 6
    "DOB",                // G col 7
    "Reg Hours",          // H col 8
    "OT Hours",           // I col 9
    "Hours Worked",       // J col 10  (formula: =H+I)
    "OT Pay",             // K col 11  (formula: =I*F)
    "Total Pay",          // L col 12  (formula: =J*F)
    "WC Rate",            // M col 13
    "WC Value",           // N col 14  (formula: =F*M*J)
  ];
  const numCols = finalHeaders.length; // 14

  const outputRows: CellValue[][] = [finalHeaders];
  const groupRows:    number[] = [];
  const subtotalRows: number[] = [];
  // Sheet row number (1-indexed) for each actual data row
  const dataRowNums:  number[] = [];
  // For each subtotal, the range of data rows it should sum
  const groupBounds: Array<{ subtotalSheet: number; startSheet: number; endSheet: number }> = [];

  let currentCode:     string | null = null;
  let groupStartSheet  = -1;

  const flushSubtotal = (): void => {
    if (currentCode === null) return;
    const lastDataSheet = outputRows.length; // last pushed row = last data row
    outputRows.push([`Subtotal — ${currentCode}`, ...Array(numCols - 1).fill("") as CellValue[]]);
    const subtotalSheetNum = outputRows.length;
    subtotalRows.push(subtotalSheetNum);
    groupBounds.push({ subtotalSheet: subtotalSheetNum, startSheet: groupStartSheet, endSheet: lastDataSheet });
  };

  dataRows.forEach((dr) => {
    if (dr.wcCode !== currentCode) {
      flushSubtotal();
      outputRows.push([dr.wcCode, ...Array(numCols - 1).fill("") as CellValue[]]);
      groupRows.push(outputRows.length);
      currentCode    = dr.wcCode;
      groupStartSheet = -1; // will be set on first data row of this group
    }
    outputRows.push(dr.staticValues);
    const sheetRowNum = outputRows.length;
    dataRowNums.push(sheetRowNum);
    if (groupStartSheet === -1) groupStartSheet = sheetRowNum;
  });
  flushSubtotal();

  // Write all static values at once
  finalSheet
    .getRange(1, 1, outputRows.length, numCols)
    .setValues(outputRows);

  // --- Write formulas for computed columns on each data row ---
  dataRowNums.forEach((r) => {
    finalSheet.getRange(r, 10).setFormula(`=H${r}+I${r}`);           // Hours Worked
    finalSheet.getRange(r, 11).setFormula(`=I${r}*F${r}`);           // OT Pay
    finalSheet.getRange(r, 12).setFormula(`=J${r}*F${r}`);           // Total Pay
    finalSheet.getRange(r, 14).setFormula(`=F${r}*M${r}*J${r}`);     // WC Value
  });

  // --- Write SUM formulas on subtotal rows ---
  groupBounds.forEach(({ subtotalSheet, startSheet, endSheet }) => {
    finalSheet.getRange(subtotalSheet, 14).setFormula(`=SUM(N${startSheet}:N${endSheet})`);
  });

  // --- Formatting: main header ---
  const headerRange = finalSheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#2E7D32");
  headerRange.setFontColor("#FFFFFF");

  // --- Formatting: group header rows ---
  groupRows.forEach((rowNum) => {
    const r = finalSheet.getRange(rowNum, 1, 1, numCols);
    r.setFontWeight("bold");
    r.setBackground("#E8F5E9");
    r.setFontColor("#1B5E20");
    r.setFontSize(10);
  });

  // --- Formatting: subtotal rows ---
  subtotalRows.forEach((rowNum) => {
    const r = finalSheet.getRange(rowNum, 1, 1, numCols);
    r.setFontWeight("bold");
    r.setBackground("#C8E6C9");
    r.setFontColor("#1B5E20");
    r.setFontStyle("italic");
    finalSheet.getRange(rowNum, 14, 1, 1).setNumberFormat("$#,##0.00"); // WC Value (col N)
  });

  // --- Formatting: data columns ---
  const totalRowCount = outputRows.length - 1;
  if (totalRowCount > 0) {
    finalSheet.getRange(2, 5,  totalRowCount, 1).setNumberFormat("MM/DD/YYYY"); // E Check Week
    finalSheet.getRange(2, 6,  totalRowCount, 1).setNumberFormat("$#,##0.00");  // F Hourly Wage
    finalSheet.getRange(2, 8,  totalRowCount, 1).setNumberFormat("0.00");        // H Reg Hours
    finalSheet.getRange(2, 9,  totalRowCount, 1).setNumberFormat("0.00");        // I OT Hours
    finalSheet.getRange(2, 10, totalRowCount, 1).setNumberFormat("0.00");        // J Hours Worked
    finalSheet.getRange(2, 11, totalRowCount, 1).setNumberFormat("$#,##0.00");  // K OT Pay
    finalSheet.getRange(2, 12, totalRowCount, 1).setNumberFormat("$#,##0.00");  // L Total Pay
    finalSheet.getRange(2, 13, totalRowCount, 1).setNumberFormat("0.000%");      // M WC Rate
    finalSheet.getRange(2, 14, totalRowCount, 1).setNumberFormat("$#,##0.00");  // N WC Value
  }

  // Auto-fit all columns
  finalSheet.autoResizeColumns(1, numCols);

  ss.toast(
    `Final report built — ${dataRows.length} employees in ${groupRows.length} WC Code groups. Check Week: ${checkWeek.toLocaleDateString()}`,
    "Workers' Comp Final",
    6,
  );
}

// ─── Step 3: Publish ──────────────────────────────────────────────────────────

function publishFinalReport(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const finalSheet = ss.getSheetByName("Final");
  if (!finalSheet) {
    ui.alert('"Final" sheet not found. Please run Build Final Report first.');
    return;
  }

  // Determine the Friday of the current week (Sun–Sat week)
  const today = new Date();
  const friday = new Date(today);
  friday.setDate(today.getDate() + (5 - today.getDay()));

  const yyyy = friday.getFullYear();
  const mm = String(friday.getMonth() + 1).padStart(2, "0");
  const dd = String(friday.getDate()).padStart(2, "0");
  const fileName = `Worker's Comp ${yyyy}${mm}${dd}`;

  const folder = DriveApp.getFolderById(SHARED_FOLDER_ID);

  // Remove any existing file with the same name
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  // Copy the spreadsheet as a native Google Sheet into the shared folder
  const sourceFile = DriveApp.getFileById(ss.getId());
  const copy = sourceFile.makeCopy(fileName, folder);

  // Open the copy and delete every sheet except "Final"
  const copySs = SpreadsheetApp.openById(copy.getId());
  const sheets = copySs.getSheets();
  const keepSheet = copySs.getSheetByName("Final");
  if (keepSheet) {
    sheets.forEach((sheet) => {
      if (sheet.getSheetId() !== keepSheet.getSheetId()) {
        copySs.deleteSheet(sheet);
      }
    });
  }

  ss.toast(
    `Published "${fileName}" to shared folder.`,
    "Workers' Comp Publish",
    6,
  );
}

// ─── Step 4: Data Needs ───────────────────────────────────────────────────────

function buildDataNeeds(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const cleanSheet = ss.getSheetByName("clean");
  if (!cleanSheet) {
    ui.alert('"clean" sheet not found. Please run the Clean Report first.');
    return;
  }

  // Get or create the Data Needs sheet and clear it
  let dataNeedsSheet = ss.getSheetByName("Data Needs");
  if (dataNeedsSheet) {
    const filter = dataNeedsSheet.getFilter();
    if (filter) filter.remove();
    dataNeedsSheet.clearContents();
    dataNeedsSheet.clearFormats();
  } else {
    dataNeedsSheet = ss.insertSheet("Data Needs");
  }

  // Read clean sheet display values (formulas already resolved)
  const cleanData = cleanSheet.getDataRange().getDisplayValues();

  // Clean sheet column indices (0-based)
  const C = {
    empName: 8, // I — Employee Name
    adpId: 9, // J — Co Code + File #
    ssn: 10, // K — SSN
    wage: 13, // N — Hourly Wage
    birthday: 17, // R — Birthday
  };

  const headers = ["Name", "ADP Id", "SSN", "Birthday", "Wage"];
  const outputRows: string[][] = [headers];

  for (let i = 1; i < cleanData.length; i++) {
    const row = cleanData[i];
    const ssn = row[C.ssn];
    const birthday = row[C.birthday];
    const wage = row[C.wage];

    const ssnMissing      = ssn === "NOT FOUND" || ssn === "";
    const birthdayMissing = birthday === "NOT FOUND" || birthday === "";
    const wageMissing     = wage === "NOT FOUND" || wage === "";

    // Only include rows where at least one of the three fields is missing
    if (!ssnMissing && !birthdayMissing && !wageMissing) continue;

    outputRows.push([
      row[C.empName],
      row[C.adpId],
      ssnMissing      ? "" : ssn,
      birthdayMissing ? "" : birthday,
      wageMissing     ? "" : wage,
    ]);
  }

  dataNeedsSheet
    .getRange(1, 1, outputRows.length, headers.length)
    .setValues(outputRows);

  // --- Formatting ---
  const headerRange = dataNeedsSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#E65100");
  headerRange.setFontColor("#FFFFFF");

  const dataRowCount = outputRows.length - 1;
  if (dataRowCount > 0) {
    // Highlight blank cells in the SSN, Birthday, Wage columns so HR can see what's missing
    const missingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground("#FFF9C4")
      .setRanges([dataNeedsSheet.getRange(2, 3, dataRowCount, 3)])
      .build();
    dataNeedsSheet.setConditionalFormatRules([missingRule]);
  }

  SpreadsheetApp.flush();
  dataNeedsSheet.autoResizeColumns(1, headers.length);
  dataNeedsSheet.getDataRange().createFilter();

  ss.toast(
    `Data Needs built — ${dataRowCount} employee${dataRowCount === 1 ? "" : "s"} with missing information.`,
    "Workers' Comp Data Needs",
    6,
  );
}
