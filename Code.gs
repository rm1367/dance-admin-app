// ===================================================================
// AWA Admin Portal — Google Apps Script Backend
// File: Code.gs
//
// SENSITIVE CONFIGURATION lives in Config.gs (not committed).
// Copy Config.gs.example → Config.gs and fill in your values before deploying.
// ===================================================================

// ---- CONFIGURATION ------------------------------------------------
// SPREADSHEET_ID and LOGO_FILE_ID are defined in Config.gs (see Config.gs.example).

// Tabs that are NOT date tabs — excluded from the date picker.
// Add any other non-date tab names here.
const SYSTEM_TABS = ['Monthlies', '2026 Weekly Template', 'StudentEmails'];

// Column positions (1-indexed: A=1, B=2, etc.)
// UPDATE these if your sheet columns differ.
const COL = {
  NAME:    1,  // A — Student name (First Last)
  CLASS:   2,  // B — Class / Social Pass type
  PAYMENT: 4,  // D — Payment method
  TOTAL:   7,  // G — Total (formula-driven — NEVER written to)
  NOTES:   8   // H — Notes + check-in timestamp
};

// Columns that contain formulas — this app will NEVER write to these.
// C(3), E(5), F(6), G(7) confirmed as formula columns.
const PROTECTED_COLS = [3, 5, 6, 7];

// Student check-in records begin on this row of every date tab.
// Rows 1–2 are reserved (headers). Data entry starts at row 3.
const CHECKIN_START_ROW = 3;

// PRICING + PAYMENT SOURCE — column Q on the weekly template tab.
// This is the same tab that drives the sheet's column B data validation,
// so the app and the sheet always show the same options.
// Q rows (class options): class/pass name in col Q, base price in col R
// Q rows (payment types): listed in col Q only, no price in col R
const PRICING_SOURCE_TAB = '2026 Weekly Template';

// Column positions for the Q/R pricing data (1-indexed)
const PRICING_COL = {
  NAME:  17,  // Q — class/pass name OR payment method label
  PRICE: 18,  // R — base price (numeric); blank for payment method rows
};

// Row range in col Q where payment methods are listed.
// Row 33 is the "PAYMENT TYPE" section header (skipped — starts at 34).
// Rows 34–37: SQUARE, CASH, VENMO, OTHER
const PAYMENT_ROWS = { START: 34, END: 37 };

// Keywords in Q text that flag a Student/Senior discount option (case-insensitive)
const STUDENT_KEYWORDS = ['student', 'discount'];

// Keywords in Q text that flag a Monthly pass option (case-insensitive)
const MONTHLY_KEYWORDS = ['month'];

// Regex pattern that identifies a tab name as a date tab.
// Accepts: 1/8  |  1/8/26  |  1/8/2026
const DATE_TAB_PATTERN = /^\d{1,2}\/\d{1,2}(\/\d{2,4})?$/;

// StudentEmails tab — created automatically if it doesn't exist
const STUDENT_EMAILS_TAB = 'StudentEmails';
const SE = { NAME: 1, EMAIL: 2, LAST_VISIT: 3 };

// Monthlies tab column layout (A–I)
const MONTHLIES_TAB = 'Monthlies';
const M = {
  NAME:          1,  // A — Student name
  EMAIL:         2,  // B — Email
  CLASS_OPTION:  3,  // C — Pass type
  PURCHASE_DATE: 4,  // D — Date pass was purchased (MM/DD/YYYY)
  CLASSES_LEFT:  5,  // E — Classes remaining (starts at 4, auto-decrements)
  CHECKIN_1:     6,  // F — Check-in date 1
  CHECKIN_2:     7,  // G — Check-in date 2
  CHECKIN_3:     8,  // H — Check-in date 3
  CHECKIN_4:     9,  // I — Check-in date 4
};

const MONTHLIES_HEADERS = [
  'Student Name', 'Email', 'Pass Type', 'Purchase Date',
  'Classes Remaining', 'Check-In 1', 'Check-In 2', 'Check-In 3', 'Check-In 4',
];

// ---- END CONFIGURATION -----------------------------------------------


// ── Custom spreadsheet menu ────────────────────────────────────────────────────
// Adds an "AWA Admin" menu to the Google Sheet toolbar automatically when the
// sheet is opened. Click "Open Check-In Portal" to launch the web app.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AWA Admin')
    .addItem('Open Check-In Portal', 'openAdminPortal')
    .addToUi();
}

// Opens the deployed web app in a new browser tab.
// Update PORTAL_URL to your current deployment URL.
const PORTAL_URL = 'YOUR_DEPLOYMENT_URL_HERE';

function openAdminPortal() {
  const html = HtmlService
    .createHtmlOutput(
      '<script>window.open("' + PORTAL_URL + '"); google.script.host.close();</script>'
    )
    .setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Admin Portal…');
}
// ── End custom menu ────────────────────────────────────────────────────────────


// Serves the web app
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AWA Admin Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fetches the business logo from Drive and returns it as a base64 data URL.
function getLogoBase64() {
  try {
    const file     = DriveApp.getFileById(LOGO_FILE_ID);
    const blob     = file.getBlob();
    const mimeType = blob.getContentType();
    const b64      = Utilities.base64Encode(blob.getBytes());
    return 'data:' + mimeType + ';base64,' + b64;
  } catch (e) {
    Logger.log('Logo fetch error: ' + e.message);
    return '';
  }
}

// Reads class names + prices from columns Q and R of the weekly template.
// Returns: { "Class Name": { price, isStudentSenior, isMonthly } }
function getPricingFromSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(PRICING_SOURCE_TAB);
  if (!sheet) { Logger.log('Pricing tab not found: ' + PRICING_SOURCE_TAB); return {}; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return {};

  const data    = sheet.getRange(1, PRICING_COL.NAME, lastRow, 2).getValues();
  const pricing = {};

  data.forEach(row => {
    const name     = String(row[0] || '').trim();
    const rawPrice = row[1];
    let price;
    if (typeof rawPrice === 'number') {
      price = rawPrice;
    } else {
      price = parseFloat(String(rawPrice).replace(/[$,\s]/g, ''));
    }
    if (!name || isNaN(price) || price < 0) return;

    const nameLower       = name.toLowerCase();
    const isStudentSenior = STUDENT_KEYWORDS.some(kw => nameLower.includes(kw));
    const isMonthly       = MONTHLY_KEYWORDS.some(kw => nameLower.includes(kw));
    pricing[name] = { price, isStudentSenior, isMonthly };
  });

  Logger.log('Loaded ' + Object.keys(pricing).length + ' pricing options from ' + PRICING_SOURCE_TAB);
  return pricing;
}

// Reads payment method labels from the weekly template.
function getPaymentMethodsFromSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(PRICING_SOURCE_TAB);
  if (!sheet) return [];

  const numRows = PAYMENT_ROWS.END - PAYMENT_ROWS.START + 1;
  const data    = sheet.getRange(PAYMENT_ROWS.START, PRICING_COL.NAME, numRows, 1).getValues();
  return data.map(row => String(row[0] || '').trim()).filter(v => v.length > 0);
}

// Returns all config to the frontend in a single round trip.
function getConfig() {
  return {
    pricing:        getPricingFromSheet(),
    paymentMethods: getPaymentMethodsFromSheet()
  };
}

function getOrCreateMonthliesTab_(ss) {
  let sheet = ss.getSheetByName(MONTHLIES_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(MONTHLIES_TAB);
    Logger.log('Created monthlies tab.');
  }
  const headerRange = sheet.getRange(1, 1, 1, MONTHLIES_HEADERS.length);
  const existing    = headerRange.getValues()[0];
  if (!existing[0]) {
    headerRange.setValues([MONTHLIES_HEADERS]);
    headerRange.setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, MONTHLIES_HEADERS.length);
  }
  return sheet;
}

function getNextMonthliesRow_(sheet) {
  return Math.max(sheet.getLastRow() + 1, 2);
}

// Returns all date tab names for the date picker (sorted newest first).
function getDateTabs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => DATE_TAB_PATTERN.test(name.trim()))
    .reverse();
}

// Returns student name/email suggestions for autocomplete (up to 8 matches).
function getStudentSuggestions(query) {
  if (!query || query.trim().length < 2) return [];

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENT_EMAILS_TAB);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const q       = query.trim().toLowerCase();
  const matches = [];

  for (const row of data) {
    const name  = String(row[SE.NAME  - 1] || '').toLowerCase();
    const email = String(row[SE.EMAIL - 1] || '').toLowerCase();
    if (name.includes(q) || email.includes(q)) {
      matches.push({ name: row[SE.NAME - 1], email: row[SE.EMAIL - 1] });
      if (matches.length >= 8) break;
    }
  }
  return matches;
}

function getNextCheckInRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CHECKIN_START_ROW) return CHECKIN_START_ROW;

  const colA = sheet.getRange(CHECKIN_START_ROW, COL.NAME, lastRow - CHECKIN_START_ROW + 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] === '' || colA[i][0] === null) return CHECKIN_START_ROW + i;
  }
  return lastRow + 1;
}

// Writes a check-in row to the selected date tab.
// Uses a script lock to prevent concurrent writes to the same row.
// Never writes to formula columns (C, E, F, G).
function submitCheckIn(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(data.tabName);
    if (!sheet) throw new Error('Tab "' + data.tabName + '" not found.');

    const name = data.name ||
      [data.firstName, data.lastName].filter(Boolean).join(' ').trim();
    if (!name) throw new Error('Student name is required.');

    const tz        = ss.getSpreadsheetTimeZone();
    const timestamp = Utilities.formatDate(new Date(), tz, 'h:mm a');
    const noteText  = data.notes ? data.notes + ' — ' + timestamp : 'Checked in ' + timestamp;
    const targetRow = getNextCheckInRow_(sheet);

    const pricing        = getPricingFromSheet();
    const classOptions   = Object.keys(pricing);
    const paymentOptions = getPaymentMethodsFromSheet();

    // Normalize values to canonical casing from the sheet
    const rawClass    = (data.classOption   || '').trim().toLowerCase();
    const classOption = classOptions.find(k => k.trim().toLowerCase() === rawClass) || (data.classOption || '');
    const rawPayment  = (data.paymentMethod || '').trim().toLowerCase();
    const paymentMethod = paymentOptions.find(p => p.trim().toLowerCase() === rawPayment) || (data.paymentMethod || '');

    // Snapshot the sheet's existing validation rules (which carry dropdown colors/styling),
    // clear them so setValue() is never rejected, write, then restore originals.
    const classRule   = sheet.getRange(targetRow, COL.CLASS).getDataValidation();
    const paymentRule = sheet.getRange(targetRow, COL.PAYMENT).getDataValidation();
    sheet.getRange(targetRow, COL.CLASS).clearDataValidations();
    sheet.getRange(targetRow, COL.PAYMENT).clearDataValidations();

    sheet.getRange(targetRow, COL.NAME).setValue(name);
    sheet.getRange(targetRow, COL.CLASS).setValue(classOption);
    sheet.getRange(targetRow, COL.PAYMENT).setValue(paymentMethod);
    sheet.getRange(targetRow, COL.NOTES).setValue(noteText);

    if (classRule)   sheet.getRange(targetRow, COL.CLASS).setDataValidation(classRule);
    if (paymentRule) sheet.getRange(targetRow, COL.PAYMENT).setDataValidation(paymentRule);

    // Write to Monthlies tab only on new pass purchase (not on existing pass check-ins)
    const classLower = classOption.toLowerCase();
    if (!data.skipMonthliesWrite && MONTHLY_KEYWORDS.some(kw => classLower.includes(kw))) {
      const mSheet     = getOrCreateMonthliesTab_(ss);
      const mTargetRow = getNextMonthliesRow_(mSheet);
      mSheet.getRange(mTargetRow, M.NAME).setValue(name);
      mSheet.getRange(mTargetRow, M.EMAIL).setValue(data.email || '');
      mSheet.getRange(mTargetRow, M.CLASS_OPTION).setValue(classOption);
      mSheet.getRange(mTargetRow, M.PURCHASE_DATE).setValue(Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy'));
      mSheet.getRange(mTargetRow, M.CLASSES_LEFT).setValue(4);
    }

    SpreadsheetApp.flush();
    updateStudentRecord(name, data.email || '');
    return { success: true, timestamp, name, row: targetRow };
  } catch (e) {
    Logger.log('submitCheckIn error: ' + e.message);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── CONNECTION TEST — run from Apps Script editor to verify setup ──
function testSheetConnection() {
  try {
    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tabs   = ss.getSheets().map(s => s.getName());
    const target = ss.getSheets().find(s => DATE_TAB_PATTERN.test(s.getName().trim())) || ss.getSheets()[0];
    const pricing  = getPricingFromSheet();
    const payments = getPaymentMethodsFromSheet();

    Logger.log('✓ Connected: ' + ss.getName());
    Logger.log('✓ Tabs: ' + tabs.join(', '));
    Logger.log('✓ Pricing options: ' + Object.keys(pricing).length);
    Logger.log('✓ Payment methods: ' + payments.join(', '));

    SpreadsheetApp.getUi().alert(
      '✓ Connection test passed!\n\n' +
      'Sheet: ' + ss.getName() + '\n' +
      'Tabs found: ' + tabs.length + '\n' +
      'Next check-in row on "' + target.getName() + '": ' + getNextCheckInRow_(target) + '\n' +
      'Pricing options: ' + Object.keys(pricing).length + '\n' +
      'Payment methods: ' + payments.join(', ')
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('✗ Connection test failed:\n\n' + e.message);
  }
}

// Upserts a student record in the StudentEmails tab (creates tab on first use).
function updateStudentRecord(name, email) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(STUDENT_EMAILS_TAB);

  if (!sheet) {
    sheet = ss.insertSheet(STUDENT_EMAILS_TAB);
    sheet.appendRow(['Full Name', 'Email', 'Last Visit']);
    ss.moveActiveSheet(ss.getNumSheets());
  }

  const today   = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy');
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      const eName  = String(data[i][SE.NAME  - 1]).toLowerCase();
      const eEmail = String(data[i][SE.EMAIL - 1]).toLowerCase();
      if ((email && eEmail === email.toLowerCase()) || eName === name.toLowerCase()) {
        const r = i + 2;
        if (name)  sheet.getRange(r, SE.NAME).setValue(name);
        if (email) sheet.getRange(r, SE.EMAIL).setValue(email);
        sheet.getRange(r, SE.LAST_VISIT).setValue(today);
        return;
      }
    }
  }
  sheet.appendRow([name, email, today]);
}

// Returns all monthly pass records with classes remaining > 0.
function getMonthlyPasses() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MONTHLIES_TAB);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data   = sheet.getRange(2, 1, lastRow - 1, M.CHECKIN_4).getValues();
  const passes = [];
  const tz     = ss.getSpreadsheetTimeZone();

  const fmtDate = d => {
    if (!d || d === '') return '';
    try { return Utilities.formatDate(new Date(d), tz, 'M/d/yy'); } catch (_) { return String(d); }
  };

  for (let i = 0; i < data.length; i++) {
    const row  = data[i];
    const name = String(row[M.NAME - 1] || '').trim();
    if (!name) continue;

    const left = Number(row[M.CLASSES_LEFT - 1]);
    if (isNaN(left) || left <= 0) continue;

    passes.push({
      rowIndex:     i + 2,
      name,
      email:        String(row[M.EMAIL        - 1] || ''),
      classOption:  String(row[M.CLASS_OPTION - 1] || ''),
      purchaseDate: fmtDate(row[M.PURCHASE_DATE - 1]),
      classesLeft:  left,
      checkins: [row[M.CHECKIN_1-1], row[M.CHECKIN_2-1], row[M.CHECKIN_3-1], row[M.CHECKIN_4-1]]
        .filter(d => d !== '' && d !== null && d !== undefined)
        .map(fmtDate)
    });
  }
  return passes;
}

// Checks in a monthly pass holder — decrements count and logs the date.
// Does NOT write to the date tab; monthly passes are tracked separately.
function checkInMonthlyStudent(rowIndex, tabName) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getOrCreateMonthliesTab_(ss);
    const tz    = ss.getSpreadsheetTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');

    const rowData     = sheet.getRange(rowIndex, 1, 1, M.CHECKIN_4).getValues()[0];
    const classesLeft = Number(rowData[M.CLASSES_LEFT - 1]);
    if (isNaN(classesLeft) || classesLeft <= 0) {
      return { success: false, error: 'No classes remaining on this pass.' };
    }

    const newCount = classesLeft - 1;
    sheet.getRange(rowIndex, M.CLASSES_LEFT).setValue(newCount);

    for (const col of [M.CHECKIN_1, M.CHECKIN_2, M.CHECKIN_3, M.CHECKIN_4]) {
      if (!rowData[col - 1]) { sheet.getRange(rowIndex, col).setValue(today); break; }
    }

    return { success: true, classesLeft: newCount, name: rowData[M.NAME - 1] };
  } catch (e) {
    Logger.log('checkInMonthlyStudent error: ' + e.message);
    return { success: false, error: e.message };
  }
}
