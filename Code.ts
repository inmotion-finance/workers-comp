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
    .addItem("Run Clean Report", "cleanWorkersCompData")
    .addSeparator()
    .addItem("Build Final Report", "buildFinalReport")
    .addSeparator()
    .addItem("Publish Final Report", "publishFinalReport")
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
    ui.alert('No sheet named "raw" found. Please rename your source sheet to "raw" and try again.');
    return;
  }

  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    ui.alert('No sheet named "Config" found. Expected the report date in Config!B1.');
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
    ui.alert('Config!B1 does not contain a valid date. Please enter the report date there and try again.');
    return;
  }

  const rawData = rawSheet.getDataRange().getValues();

  const COL = {
    coCode:   0,
    batchId:  1,
    fileNum:  2,
    taxFreq:  3,
    tempDept: 4,
    tempRate: 5,
    regHours: 6,
    otHours:  7,
    empName:  8
  };

  const employeeMap = new Map<string, EmployeeRecord>();

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const fileNum = row[COL.fileNum];
    if (!fileNum && fileNum !== 0) continue;

    const key = String(fileNum);

    if (!employeeMap.has(key)) {
      employeeMap.set(key, {
        coCode:   row[COL.coCode],
        batchId:  row[COL.batchId],
        fileNum:  fileNum,
        taxFreq:  row[COL.taxFreq],
        tempDept: row[COL.tempDept],
        tempRate: row[COL.tempRate],
        regHours: 0,
        otHours:  0,
        empName:  row[COL.empName]
      });
    }

    const emp = employeeMap.get(key)!;
    emp.regHours += (typeof row[COL.regHours] === "number") ? (row[COL.regHours] as number) : 0;
    emp.otHours  += (typeof row[COL.otHours]  === "number") ? (row[COL.otHours] as number)  : 0;

    if (!emp.taxFreq  && row[COL.taxFreq])  emp.taxFreq  = row[COL.taxFreq];
    if (!emp.tempDept && row[COL.tempDept]) emp.tempDept = row[COL.tempDept];
    if (!emp.tempRate && row[COL.tempRate]) emp.tempRate = row[COL.tempRate];
    if (!emp.empName  && row[COL.empName])  emp.empName  = row[COL.empName];
  }

  const outputHeaders = [
    'Co Code', 'Batch ID', 'File #', 'Tax Frequency', 'Temp Dept',
    'Temp Rate', 'Reg Hours', 'O/T Hours', 'Employee Name',
    'Co Code + File #', 'SSN', 'Hire Date', 'State', 'Hourly Wage', 'WC Code', 'WC Rate', 'WC Value', 'Birthday', 'Needs Review'
  ];
  const outputRows: CellValue[][] = [outputHeaders];

  employeeMap.forEach(emp => {
    const concatKey = String(emp.coCode) + "0" + String(emp.fileNum).padStart(5, "0");
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
      "",   // K — SSN
      "",   // L — Hire Date
      "",   // M — State
      "",   // N — Hourly Wage
      "",   // O — WC Code
      "",   // P — WC Rate
      "",   // Q — WC Value
      "",   // R — Birthday
      ""    // S — Needs Review
    ]);
  });

  cleanSheet.getRange(1, 1, outputRows.length, outputHeaders.length).setValues(outputRows);

  const dataRowCount = outputRows.length - 1;

  if (dataRowCount > 0) {
    const ssnFormulas:        string[][] = [];
    const hireDateFormulas:   string[][] = [];
    const stateFormulas:      string[][] = [];
    const hourlyWageFormulas: string[][] = [];
    const wcCodeFormulas:     string[][] = [];
    const wcRateFormulas:     string[][] = [];
    const wcValueFormulas:      string[][] = [];
    const birthdayFormulas:     string[][] = [];
    const needsReviewFormulas:  string[][] = [];

    for (let r = 2; r <= dataRowCount + 1; r++) {
      ssnFormulas.push([
        `=IFERROR(LET(v,XLOOKUP($J${r},'Personnel File'!$A:$A,'Personnel File'!$G:$G),IF(v="","NOT FOUND",v)),"NOT FOUND")`
      ]);
      hireDateFormulas.push([
        `=IFERROR(QUERY(Contracts!$A$2:$G,"SELECT F WHERE B = '"&$J${r}&"' AND F <= date '"&TEXT(Config!$B$2,"yyyy-mm-dd")&"' AND G >= date '"&TEXT(Config!$B$1,"yyyy-mm-dd")&"' LIMIT 1",0),"NOT FOUND")`
      ]);
      stateFormulas.push([
        `=IFERROR(QUERY(Contracts!$A$2:$G,"SELECT E WHERE B = '"&$J${r}&"' AND F <= date '"&TEXT(Config!$B$2,"yyyy-mm-dd")&"' AND G >= date '"&TEXT(Config!$B$1,"yyyy-mm-dd")&"' LIMIT 1",0),"NOT FOUND")`
      ]);
      hourlyWageFormulas.push([
        `=IFERROR(QUERY(Contracts!$A$2:$J,"SELECT J WHERE B = '"&$J${r}&"' AND F <= date '"&TEXT(Config!$B$2,"yyyy-mm-dd")&"' AND G >= date '"&TEXT(Config!$B$1,"yyyy-mm-dd")&"' LIMIT 1",0),"NOT FOUND")`
      ]);
      wcCodeFormulas.push([
        `=IFERROR(LET(v,XLOOKUP($M${r},'WC Rates'!$A$2:$A$31,'WC Rates'!$B$2:$B$31),IF(v="","NOT FOUND",v)),"NOT FOUND")`
      ]);
      wcRateFormulas.push([
        `=IFERROR(LET(v,XLOOKUP($M${r},'WC Rates'!$A$2:$A$31,'WC Rates'!$C$2:$C$31),IF(v="","NOT FOUND",v)),"NOT FOUND")`
      ]);
      wcValueFormulas.push([
        `=IFERROR($N${r}*$P${r}*($G${r}+$H${r}),"NOT FOUND")`
      ]);
      birthdayFormulas.push([
        `=IFERROR(LET(v,XLOOKUP($J${r},'Personnel File'!$A:$A,'Personnel File'!$H:$H),IF(v="","NOT FOUND",v)),"NOT FOUND")`
      ]);
      needsReviewFormulas.push([
        `=IF(OR($K${r}="NOT FOUND",$L${r}="NOT FOUND",$M${r}="NOT FOUND",$N${r}="NOT FOUND",$O${r}="NOT FOUND",$P${r}="NOT FOUND",$R${r}="NOT FOUND"),"YES","")`
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
    cleanSheet.getRange(2, 19, dataRowCount, 1).setFormulas(needsReviewFormulas);

    // Conditional formatting: highlight "Needs Review" rows in red
    const reviewRange = cleanSheet.getRange(2, 19, dataRowCount, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("YES")
      .setBackground("#FFCDD2")
      .setFontColor("#B71C1C")
      .setRanges([reviewRange])
      .build();
    const existingRules = cleanSheet.getConditionalFormatRules().filter(
      r => !r.getRanges().some(rng => rng.getColumn() === 19)
    );
    cleanSheet.setConditionalFormatRules([...existingRules, rule]);
  }

  // --- Formatting ---
  const headerRange = cleanSheet.getRange(1, 1, 1, outputHeaders.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4A90D9");
  headerRange.setFontColor("#FFFFFF");

  if (dataRowCount > 0) {
    cleanSheet.getRange(2, 7,  dataRowCount, 2).setNumberFormat("0.00");       // Reg/OT Hours
    cleanSheet.getRange(2, 12, dataRowCount, 1).setNumberFormat("MM/DD/YYYY"); // Hire Date
    cleanSheet.getRange(2, 14, dataRowCount, 1).setNumberFormat("$#,##0.00");  // Hourly Wage
    cleanSheet.getRange(2, 16, dataRowCount, 1).setNumberFormat("0.000%");     // WC Rate
    cleanSheet.getRange(2, 17, dataRowCount, 1).setNumberFormat("$#,##0.00");  // WC Value
    cleanSheet.getRange(2, 18, dataRowCount, 1).setNumberFormat("MM/DD/YYYY"); // Birthday
  }

  // Flush to ensure formulas are resolved before auto-fitting
  SpreadsheetApp.flush();

  // Auto-fit all columns
  cleanSheet.autoResizeColumns(1, outputHeaders.length);

  ss.toast(
    `Done! ${rawData.length - 1} raw rows → ${dataRowCount} employees. Report date: ${reportDate.toLocaleDateString()}`,
    "Workers' Comp Cleaner",
    6
  );
}

// ─── Step 2: Final ────────────────────────────────────────────────────────────

interface DataRow {
  wcCode: string;
  state: string;
  values: CellValue[];
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

  const personnelSheet = ss.getSheetByName("Personnel File");
  if (!personnelSheet) {
    ui.alert('"Personnel File" sheet not found. DOB column will show NOT FOUND.');
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
    ui.alert('Config!B1 does not contain a valid date.');
    return;
  }

  // --- DOB lookup from Personnel File (col A = key, col H = DOB) ---
  const dobMap = new Map<string, CellValue>();
  if (personnelSheet) {
    const personnelData = personnelSheet.getDataRange().getValues();
    for (let i = 1; i < personnelData.length; i++) {
      const key = String(personnelData[i][0]).trim();
      if (!key) continue;
      dobMap.set(key, personnelData[i][7] ?? "NOT FOUND");
    }
  }

  // --- Read clean sheet values (formulas already calculated) ---
  const cleanData = cleanSheet.getDataRange().getDisplayValues();

  const C = {
    empName:    8,   // I
    concatKey:  9,   // J
    ssn:        10,  // K
    hireDate:   11,  // L
    state:      12,  // M
    hourlyWage: 13,  // N
    wcCode:     14,  // O
    wcRate:     15,  // P
    regHours:   6,   // G
    otHours:    7,   // H
    wcValue:    16   // Q
  };

  // --- Collect and sort rows by WC Code then Employee Name ---
  const dataRows: DataRow[] = [];
  for (let i = 1; i < cleanData.length; i++) {
    const row        = cleanData[i];
    const concatKey  = String(row[C.concatKey]).trim();
    const empName    = row[C.empName];
    const ssn        = row[C.ssn];
    const hireDate   = row[C.hireDate];
    const wcCode     = row[C.wcCode];
    const state      = row[C.state];
    const hourlyWage = parseFloat(row[C.hourlyWage].replace(/[$,]/g, "")) || 0;
    const wcRate     = row[C.wcRate];
    const wcValue    = parseFloat(row[C.wcValue].replace(/[$,]/g, "")) || 0;
    const regHours   = parseFloat(row[C.regHours]) || 0;
    const otHours    = parseFloat(row[C.otHours])  || 0;
    const dob        = dobMap.get(concatKey) ?? "NOT FOUND";
    const hoursWorked = regHours + otHours;
    const otPay      = otHours * hourlyWage;
    const totalPay   = hoursWorked * hourlyWage;

    dataRows.push({
      wcCode,
      state,
      values: [
        empName,
        ssn,
        hireDate,
        wcCode,
        checkWeek,
        hourlyWage,
        dob,
        hoursWorked,
        otPay,
        totalPay,
        wcRate,
        wcValue
      ]
    });
  }

  // Sort by WC Code, then by Employee Name within each group
  dataRows.sort((a, b) => {
    const codeCompare = String(a.wcCode).localeCompare(String(b.wcCode));
    if (codeCompare !== 0) return codeCompare;
    return String(a.values[0]).localeCompare(String(b.values[0]));
  });

  // --- Build final output with group header rows and subtotals ---
  const finalHeaders = [
    'Name', 'SS#', 'Employee Hire Date', 'WKC Code', 'Check Week',
    'Hourly Wage', 'DOB', 'Hours Worked', 'OT Pay', 'Total Pay', 'WC Rate', 'WC Value'
  ];

  const outputRows: CellValue[][] = [finalHeaders];
  const groupRows:    number[] = [];
  const subtotalRows: number[] = [];
  let currentCode: string | null = null;
  let groupWcValue = 0;

  const flushSubtotal = (): void => {
    if (currentCode === null) return;
    outputRows.push([
      `Subtotal — ${currentCode}`, "", "", "", "", "", "", "", "", "", "", groupWcValue
    ]);
    subtotalRows.push(outputRows.length);
    groupWcValue = 0;
  };

  dataRows.forEach((dr) => {
    if (dr.wcCode !== currentCode) {
      flushSubtotal();
      outputRows.push([dr.wcCode, "", "", "", "", "", "", "", "", "", "", ""]);
      groupRows.push(outputRows.length);
      currentCode = dr.wcCode;
    }
    groupWcValue += (typeof dr.values[11] === "number" ? dr.values[11] : parseFloat(String(dr.values[11]))) || 0;
    outputRows.push(dr.values);
  });
  flushSubtotal();

  finalSheet.getRange(1, 1, outputRows.length, finalHeaders.length).setValues(outputRows);

  // --- Formatting: main header ---
  const headerRange = finalSheet.getRange(1, 1, 1, finalHeaders.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#2E7D32");
  headerRange.setFontColor("#FFFFFF");

  // --- Formatting: group header rows ---
  groupRows.forEach(rowNum => {
    const groupRange = finalSheet.getRange(rowNum, 1, 1, finalHeaders.length);
    groupRange.setFontWeight("bold");
    groupRange.setBackground("#E8F5E9");
    groupRange.setFontColor("#1B5E20");
    groupRange.setFontSize(10);
  });

  // --- Formatting: subtotal rows ---
  subtotalRows.forEach(rowNum => {
    const subtotalRange = finalSheet.getRange(rowNum, 1, 1, finalHeaders.length);
    subtotalRange.setFontWeight("bold");
    subtotalRange.setBackground("#C8E6C9");
    subtotalRange.setFontColor("#1B5E20");
    subtotalRange.setFontStyle("italic");
    finalSheet.getRange(rowNum, 12, 1, 1).setNumberFormat("$#,##0.00");
  });

  // --- Formatting: data columns ---
  const dataRowCount = outputRows.length - 1;
  if (dataRowCount > 0) {
    finalSheet.getRange(2, 5,  dataRowCount, 1).setNumberFormat("MM/DD/YYYY"); // Check Week
    finalSheet.getRange(2, 6,  dataRowCount, 1).setNumberFormat("$#,##0.00");  // Hourly Wage
    finalSheet.getRange(2, 8,  dataRowCount, 1).setNumberFormat("0.00");       // Hours Worked
    finalSheet.getRange(2, 9,  dataRowCount, 1).setNumberFormat("$#,##0.00");  // OT Pay
    finalSheet.getRange(2, 10, dataRowCount, 1).setNumberFormat("$#,##0.00");  // Total Pay
    finalSheet.getRange(2, 12, dataRowCount, 1).setNumberFormat("$#,##0.00");  // WC Value
  }

  // Auto-fit all columns
  for (let c = 1; c <= finalHeaders.length; c++) {
    finalSheet.autoResizeColumn(c);
  }

  ss.toast(
    `Final report built — ${dataRows.length} employees in ${groupRows.length} WC Code groups. Check Week: ${checkWeek.toLocaleDateString()}`,
    "Workers' Comp Final",
    6
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

  const yyyy  = friday.getFullYear();
  const mm    = String(friday.getMonth() + 1).padStart(2, "0");
  const dd    = String(friday.getDate()).padStart(2, "0");
  const fileName = `Worker's Comp ${yyyy}${mm}${dd}`;

  // Export the Final sheet as XLSX via the Sheets export URL
  const ssId      = ss.getId();
  const sheetId   = finalSheet.getSheetId();
  const exportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const blob   = response.getBlob().setName(`${fileName}.xlsx`);
  const folder = DriveApp.getFolderById(SHARED_FOLDER_ID);

  // Replace any existing file with the same name
  const existing = folder.getFilesByName(`${fileName}.xlsx`);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  folder.createFile(blob);

  ss.toast(
    `Published "${fileName}.xlsx" to shared folder.`,
    "Workers' Comp Publish",
    6
  );
}
