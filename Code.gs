/**
 * Service-to-Sales Bridge Dashboard
 * Union Park Buick GMC
 *
 * Apps Script for managing communication between Sales and Service departments
 *
 * Authorized Users:
 * - Brian Callahan (callahanscollectables@gmail.com) - Sales Manager
 * - Dan Testa (dtesta620@gmail.com) - Service Manager
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const CONFIG = {
  // Authorized users
  BRIAN_EMAIL: 'callahanscollectables@gmail.com',
  DAN_EMAIL: 'dtesta620@gmail.com',

  // Name mapping
  NAMES: {
    'callahanscollectables@gmail.com': 'Brian Callahan',
    'dtesta620@gmail.com': 'Dan Testa'
  },

  // Sheet names
  SHEETS: {
    DASHBOARD: 'Dashboard',
    DEALER_TRADE: 'Dealer Trade Re-PDIs',
    ACCESSORY_INSTALLS: 'Customer Accessory Installs',
    NEW_CAR_PARTS: 'New Car Parts Installation',
    SERVICE_APPRAISALS: 'Service Drive Appraisals',
    ARCHIVE: 'Completed Archive'
  },

  // Colors (GM Blue professional palette)
  COLORS: {
    PRIMARY_BLUE: '#003366',
    SECONDARY_BLUE: '#0066CC',
    LIGHT_BLUE: '#E6F2FF',
    URGENT_RED: '#FF4444',
    WARNING_ORANGE: '#FF9900',
    WARNING_YELLOW: '#FFEB3B',
    SUCCESS_GREEN: '#4CAF50',
    COMPLETED_GRAY: '#CCCCCC',
    HOT_RED: '#FF0000',
    WARM_YELLOW: '#FFC107',
    COLD_BLUE: '#2196F3',
    WHITE: '#FFFFFF',
    HEADER_BG: '#003366',
    ALT_ROW: '#F5F9FF'
  },

  // Archive after X days
  ARCHIVE_DAYS: 7,

  // Priority thresholds (hours)
  URGENT_HOURS: 24,
  HIGH_PRIORITY_HOURS: 48
};

// =============================================================================
// MENU & INITIALIZATION
// =============================================================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Bridge Dashboard')
    .addSubMenu(ui.createMenu('Add New Entry')
      .addItem('Dealer Trade Re-PDI', 'showDealerTradeForm')
      .addItem('Customer Accessory Install', 'showAccessoryForm')
      .addItem('New Car Parts Install', 'showPartsForm')
      .addItem('Service Drive Appraisal', 'showAppraisalForm'))
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Run Archive Now', 'archiveCompletedItems')
    .addSeparator()
    .addItem('Setup Sheet Structure', 'setupAllSheets')
    .addItem('Setup Triggers', 'setupTriggers')
    .addItem('Remove All Locks', 'removeAllProtections')
    .addToUi();
}

/**
 * Initial setup - run this first!
 */
function initialSetup() {
  setupAllSheets();
  setupTriggers();
  applyConditionalFormatting();
  refreshDashboard();

  SpreadsheetApp.getUi().alert(
    'Setup Complete!',
    'The Service-to-Sales Bridge Dashboard has been configured.\n\n' +
    'Email notifications are now active.\n\n' +
    'Share this sheet with Dan (dtesta620@gmail.com) using the Share button!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================================================
// SHEET STRUCTURE SETUP
// =============================================================================

/**
 * Sets up all sheets with proper structure
 */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create sheets in order
  setupDashboardSheet(ss);
  setupDealerTradeSheet(ss);
  setupAccessoryInstallsSheet(ss);
  setupNewCarPartsSheet(ss);
  setupServiceAppraisalsSheet(ss);
  setupArchiveSheet(ss);

  // Delete default Sheet1 if it exists
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1) {
    ss.deleteSheet(sheet1);
  }
}

/**
 * Setup Dashboard sheet
 */
function setupDashboardSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DASHBOARD, 0);
  } else {
    sheet.clear();
  }

  // Set column widths
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);

  // Title
  sheet.getRange('A1').setValue('SERVICE-TO-SALES BRIDGE DASHBOARD')
    .setFontSize(24)
    .setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.PRIMARY_BLUE);
  sheet.getRange('A1:E1').merge();

  sheet.getRange('A2').setValue('Union Park Buick GMC')
    .setFontSize(14)
    .setFontColor(CONFIG.COLORS.SECONDARY_BLUE);
  sheet.getRange('A2:E2').merge();

  sheet.getRange('A3').setValue('Last Updated: ' + new Date().toLocaleString())
    .setFontSize(10)
    .setFontColor('#666666');
  sheet.getRange('A3:E3').merge();

  // Priority Section Header
  sheet.getRange('A5').setValue('PRIORITY: SOLD UNITS DELIVERING WITHIN 48 HOURS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.URGENT_RED)
    .setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('A5:E5').merge();

  // Priority section placeholder
  sheet.getRange('A6').setValue('Stock #').setFontWeight('bold').setBackground(CONFIG.COLORS.LIGHT_BLUE);
  sheet.getRange('B6').setValue('Vehicle').setFontWeight('bold').setBackground(CONFIG.COLORS.LIGHT_BLUE);
  sheet.getRange('C6').setValue('Customer').setFontWeight('bold').setBackground(CONFIG.COLORS.LIGHT_BLUE);
  sheet.getRange('D6').setValue('Delivery').setFontWeight('bold').setBackground(CONFIG.COLORS.LIGHT_BLUE);
  sheet.getRange('E6').setValue('Status').setFontWeight('bold').setBackground(CONFIG.COLORS.LIGHT_BLUE);

  // Summary Section
  sheet.getRange('A12').setValue('SUMMARY COUNTS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY_BLUE)
    .setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('A12:E12').merge();

  // Headers for summary
  sheet.getRange('A13').setValue('Category').setFontWeight('bold').setBackground(CONFIG.COLORS.LIGHT_BLUE);
  sheet.getRange('B13').setValue('Pending').setFontWeight('bold').setBackground(CONFIG.COLORS.WARNING_YELLOW);
  sheet.getRange('C13').setValue('Scheduled').setFontWeight('bold').setBackground(CONFIG.COLORS.SECONDARY_BLUE).setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('D13').setValue('Completed Today').setFontWeight('bold').setBackground(CONFIG.COLORS.SUCCESS_GREEN).setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('E13').setValue('').setBackground(CONFIG.COLORS.LIGHT_BLUE);

  // Category rows
  const categories = [
    'Dealer Trade Re-PDIs',
    'Accessory Installs',
    'New Car Parts Needed',
    'Service Drive Appraisals (Hot/Warm/Cold Open)'
  ];

  for (let i = 0; i < categories.length; i++) {
    sheet.getRange(14 + i, 1).setValue(categories[i]);
    sheet.getRange(14 + i, 2).setValue('0');
    sheet.getRange(14 + i, 3).setValue('0');
    sheet.getRange(14 + i, 4).setValue('0');
  }

  // Quick Links Section
  sheet.getRange('A20').setValue('QUICK LINKS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.PRIMARY_BLUE)
    .setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('A20:E20').merge();

  sheet.getRange('A21').setValue('Click sheet tabs below to navigate to each section');
  sheet.getRange('A21:E21').merge();

  // Freeze header
  sheet.setFrozenRows(4);

  // Set tab color
  sheet.setTabColor(CONFIG.COLORS.PRIMARY_BLUE);
}

/**
 * Setup Dealer Trade Re-PDIs sheet
 */
function setupDealerTradeSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DEALER_TRADE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DEALER_TRADE);
  } else {
    sheet.clear();
  }

  const headers = [
    'Date Submitted',
    'Submitted By',
    'Stock Number',
    'VIN (Last 8)',
    'Year/Make/Model',
    'SOLD?',
    'Customer Name',
    'Delivery Date',
    'Delivery Time',
    'Priority Flag',
    'Status',
    'Assigned To',
    'Completion Date',
    'Notes'
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 130);  // Date Submitted
  sheet.setColumnWidth(2, 180);  // Submitted By
  sheet.setColumnWidth(3, 100);  // Stock Number
  sheet.setColumnWidth(4, 100);  // VIN
  sheet.setColumnWidth(5, 180);  // Year/Make/Model
  sheet.setColumnWidth(6, 70);   // SOLD?
  sheet.setColumnWidth(7, 150);  // Customer Name
  sheet.setColumnWidth(8, 120);  // Delivery Date
  sheet.setColumnWidth(9, 100);  // Delivery Time
  sheet.setColumnWidth(10, 100); // Priority Flag
  sheet.setColumnWidth(11, 120); // Status
  sheet.setColumnWidth(12, 150); // Assigned To
  sheet.setColumnWidth(13, 130); // Completion Date
  sheet.setColumnWidth(14, 250); // Notes

  // Data validation - SOLD dropdown
  const soldRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(soldRule);

  // Data validation - Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Scheduled', 'In Progress', 'Completed'], true)
    .build();
  sheet.getRange('K2:K1000').setDataValidation(statusRule);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Apply alternating colors
  applyAlternatingColors(sheet, headers.length);

  // Set tab color
  sheet.setTabColor(CONFIG.COLORS.SECONDARY_BLUE);
}

/**
 * Setup Customer Accessory Installs sheet
 */
function setupAccessoryInstallsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ACCESSORY_INSTALLS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ACCESSORY_INSTALLS);
  } else {
    sheet.clear();
  }

  const headers = [
    'Date Submitted',
    'Submitted By',
    'Customer Name',
    'Customer Phone',
    'Customer Email',
    'Vehicle (Y/M/M)',
    'Stock # or VIN',
    'Part Number(s)',
    'Part Description',
    'Part Ordered?',
    'Part Received?',
    'Requires Tech Install?',
    'Status',
    'Appointment Date/Time',
    'Notes'
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 130);  // Date Submitted
  sheet.setColumnWidth(2, 180);  // Submitted By
  sheet.setColumnWidth(3, 150);  // Customer Name
  sheet.setColumnWidth(4, 120);  // Customer Phone
  sheet.setColumnWidth(5, 180);  // Customer Email
  sheet.setColumnWidth(6, 150);  // Vehicle
  sheet.setColumnWidth(7, 120);  // Stock/VIN
  sheet.setColumnWidth(8, 120);  // Part Number
  sheet.setColumnWidth(9, 200);  // Part Description
  sheet.setColumnWidth(10, 100); // Part Ordered
  sheet.setColumnWidth(11, 100); // Part Received
  sheet.setColumnWidth(12, 130); // Requires Install
  sheet.setColumnWidth(13, 150); // Status
  sheet.setColumnWidth(14, 150); // Appointment
  sheet.setColumnWidth(15, 250); // Notes

  // Data validation - Yes/No dropdowns
  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .build();
  sheet.getRange('J2:J1000').setDataValidation(yesNoRule);
  sheet.getRange('K2:K1000').setDataValidation(yesNoRule);
  sheet.getRange('L2:L1000').setDataValidation(yesNoRule);

  // Data validation - Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Part Ordered', 'Part Received', 'Appointment Scheduled', 'Completed'], true)
    .build();
  sheet.getRange('M2:M1000').setDataValidation(statusRule);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Apply alternating colors
  applyAlternatingColors(sheet, headers.length);

  // Set tab color
  sheet.setTabColor(CONFIG.COLORS.SUCCESS_GREEN);
}

/**
 * Setup New Car Parts Installation sheet
 */
function setupNewCarPartsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.NEW_CAR_PARTS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.NEW_CAR_PARTS);
  } else {
    sheet.clear();
  }

  const headers = [
    'Date Submitted',
    'Submitted By',
    'Stock Number',
    'VIN',
    'Year/Make/Model',
    'Part Number(s)',
    'Part Description',
    'SOLD?',
    'Customer Name',
    'Delivery Date',
    'Priority Flag',
    'Status',
    'Notes'
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 130);  // Date Submitted
  sheet.setColumnWidth(2, 180);  // Submitted By
  sheet.setColumnWidth(3, 100);  // Stock Number
  sheet.setColumnWidth(4, 150);  // VIN
  sheet.setColumnWidth(5, 180);  // Year/Make/Model
  sheet.setColumnWidth(6, 120);  // Part Numbers
  sheet.setColumnWidth(7, 200);  // Part Description
  sheet.setColumnWidth(8, 70);   // SOLD?
  sheet.setColumnWidth(9, 150);  // Customer Name
  sheet.setColumnWidth(10, 120); // Delivery Date
  sheet.setColumnWidth(11, 100); // Priority Flag
  sheet.setColumnWidth(12, 120); // Status
  sheet.setColumnWidth(13, 250); // Notes

  // Data validation - SOLD dropdown
  const soldRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .build();
  sheet.getRange('H2:H1000').setDataValidation(soldRule);

  // Data validation - Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Parts Received', 'Scheduled', 'Completed'], true)
    .build();
  sheet.getRange('L2:L1000').setDataValidation(statusRule);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Apply alternating colors
  applyAlternatingColors(sheet, headers.length);

  // Set tab color
  sheet.setTabColor(CONFIG.COLORS.WARNING_ORANGE);
}

/**
 * Setup Service Drive Appraisals sheet
 */
function setupServiceAppraisalsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SERVICE_APPRAISALS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SERVICE_APPRAISALS);
  } else {
    sheet.clear();
  }

  const headers = [
    'Date Submitted',
    'Submitted By',
    'Customer Name',
    'Customer Phone',
    'Customer Email',
    'Vehicle (Y/M/M)',
    'Mileage',
    'Service Being Performed',
    'Heat Level',
    'Reason for Heat Level',
    'Status',
    'Assigned Salesperson',
    'Follow-up Date',
    'Notes'
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 130);  // Date Submitted
  sheet.setColumnWidth(2, 180);  // Submitted By
  sheet.setColumnWidth(3, 150);  // Customer Name
  sheet.setColumnWidth(4, 120);  // Customer Phone
  sheet.setColumnWidth(5, 180);  // Customer Email
  sheet.setColumnWidth(6, 150);  // Vehicle
  sheet.setColumnWidth(7, 80);   // Mileage
  sheet.setColumnWidth(8, 200);  // Service
  sheet.setColumnWidth(9, 100);  // Heat Level
  sheet.setColumnWidth(10, 250); // Reason
  sheet.setColumnWidth(11, 130); // Status
  sheet.setColumnWidth(12, 150); // Assigned
  sheet.setColumnWidth(13, 120); // Follow-up
  sheet.setColumnWidth(14, 250); // Notes

  // Data validation - Heat Level dropdown
  const heatRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Hot', 'Warm', 'Cold'], true)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(heatRule);

  // Data validation - Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New Lead', 'Contacted', 'Appointment Set', 'Sold', 'Lost'], true)
    .build();
  sheet.getRange('K2:K1000').setDataValidation(statusRule);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Apply alternating colors
  applyAlternatingColors(sheet, headers.length);

  // Set tab color
  sheet.setTabColor(CONFIG.COLORS.HOT_RED);
}

/**
 * Setup Archive sheet
 */
function setupArchiveSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ARCHIVE);
  } else {
    sheet.clear();
  }

  const headers = [
    'Archive Date',
    'Original Sheet',
    'Date Submitted',
    'Submitted By',
    'Stock #/Customer',
    'Description',
    'Status',
    'Completion Date',
    'All Data (JSON)'
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(CONFIG.COLORS.COMPLETED_GRAY)
    .setFontColor(CONFIG.COLORS.PRIMARY_BLUE)
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 250);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 400);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set tab color
  sheet.setTabColor(CONFIG.COLORS.COMPLETED_GRAY);
}

/**
 * Apply alternating row colors to a sheet
 */
function applyAlternatingColors(sheet, numCols) {
  const range = sheet.getRange(2, 1, 998, numCols);
  const banding = range.getBandings();

  // Remove existing banding
  banding.forEach(b => b.remove());

  // Apply new banding
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

// =============================================================================
// CONDITIONAL FORMATTING
// =============================================================================

/**
 * Apply all conditional formatting rules
 */
function applyConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Dealer Trade Re-PDIs formatting
  applyDealerTradeFormatting(ss);

  // New Car Parts formatting
  applyNewCarPartsFormatting(ss);

  // Service Appraisals formatting
  applyServiceAppraisalsFormatting(ss);
}

/**
 * Apply formatting to Dealer Trade sheet
 */
function applyDealerTradeFormatting(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DEALER_TRADE);
  if (!sheet) return;

  // Clear existing rules
  sheet.clearConditionalFormatRules();

  const rules = [];
  const dataRange = sheet.getRange('A2:N1000');

  // SOLD = Yes - highlight entire row yellow
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Yes"')
    .setBackground('#FFF9C4')
    .setRanges([dataRange])
    .build());

  // Priority = URGENT - red background
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J2="URGENT"')
    .setBackground(CONFIG.COLORS.URGENT_RED)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setRanges([sheet.getRange('J2:J1000')])
    .build());

  // Status = Completed - gray out
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="Completed"')
    .setBackground(CONFIG.COLORS.COMPLETED_GRAY)
    .setRanges([dataRange])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * Apply formatting to New Car Parts sheet
 */
function applyNewCarPartsFormatting(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEW_CAR_PARTS);
  if (!sheet) return;

  sheet.clearConditionalFormatRules();

  const rules = [];
  const dataRange = sheet.getRange('A2:M1000');

  // SOLD = Yes - highlight
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2="Yes"')
    .setBackground('#FFF9C4')
    .setRanges([dataRange])
    .build());

  // Priority = URGENT
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="URGENT"')
    .setBackground(CONFIG.COLORS.URGENT_RED)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setRanges([sheet.getRange('K2:K1000')])
    .build());

  // Status = Completed
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$L2="Completed"')
    .setBackground(CONFIG.COLORS.COMPLETED_GRAY)
    .setRanges([dataRange])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * Apply formatting to Service Appraisals sheet
 */
function applyServiceAppraisalsFormatting(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SERVICE_APPRAISALS);
  if (!sheet) return;

  sheet.clearConditionalFormatRules();

  const rules = [];

  // Heat Level = Hot - red background
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Hot')
    .setBackground(CONFIG.COLORS.HOT_RED)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setRanges([sheet.getRange('I2:I1000')])
    .build());

  // Heat Level = Warm - yellow background
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Warm')
    .setBackground(CONFIG.COLORS.WARM_YELLOW)
    .setRanges([sheet.getRange('I2:I1000')])
    .build());

  // Heat Level = Cold - blue background
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cold')
    .setBackground(CONFIG.COLORS.COLD_BLUE)
    .setFontColor(CONFIG.COLORS.WHITE)
    .setRanges([sheet.getRange('I2:I1000')])
    .build());

  // Status = Sold or Lost - gray out
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($K2="Sold",$K2="Lost")')
    .setBackground(CONFIG.COLORS.COMPLETED_GRAY)
    .setRanges([sheet.getRange('A2:N1000')])
    .build());

  sheet.setConditionalFormatRules(rules);
}

// =============================================================================
// AUTO-POPULATION & TRIGGERS
// =============================================================================

/**
 * Setup all triggers
 */
function setupTriggers() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // onEdit trigger for auto-population and notifications
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  // Daily trigger for archiving (runs at 2 AM)
  ScriptApp.newTrigger('archiveCompletedItems')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  // Hourly trigger for dashboard refresh
  ScriptApp.newTrigger('refreshDashboard')
    .timeBased()
    .everyHours(1)
    .create();
}

/**
 * Main edit handler - called on every edit
 */
function onEditHandler(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // Skip header row and non-data sheets
  if (row === 1 || sheetName === CONFIG.SHEETS.DASHBOARD || sheetName === CONFIG.SHEETS.ARCHIVE) {
    return;
  }

  // Handle new row entry (auto-populate timestamp and user)
  if (col === 3 && row > 1) { // Assuming column 3 is first user-entered data
    handleNewEntry(sheet, row, sheetName);
  }

  // Handle status changes
  if (isStatusColumn(sheetName, col)) {
    handleStatusChange(sheet, row, sheetName, e.value, e.oldValue);
  }

  // Calculate priority flag
  calculatePriorityFlag(sheet, row, sheetName);

  // Refresh dashboard
  refreshDashboard();
}

/**
 * Check if the edited column is a status column
 */
function isStatusColumn(sheetName, col) {
  const statusCols = {
    [CONFIG.SHEETS.DEALER_TRADE]: 11,
    [CONFIG.SHEETS.ACCESSORY_INSTALLS]: 13,
    [CONFIG.SHEETS.NEW_CAR_PARTS]: 12,
    [CONFIG.SHEETS.SERVICE_APPRAISALS]: 11
  };
  return statusCols[sheetName] === col;
}

/**
 * Handle new entry - auto-populate timestamp and user
 */
function handleNewEntry(sheet, row, sheetName) {
  const dateCol = 1;
  const userCol = 2;

  // Check if timestamp already exists
  if (sheet.getRange(row, dateCol).getValue() === '') {
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();

    sheet.getRange(row, dateCol).setValue(timestamp);
    sheet.getRange(row, userCol).setValue(userEmail);

    // Send notification for new entry
    sendNewEntryNotification(sheet, row, sheetName, userEmail);
  }
}

/**
 * Calculate and set priority flag
 */
function calculatePriorityFlag(sheet, row, sheetName) {
  let soldCol, deliveryCol, priorityCol;

  switch (sheetName) {
    case CONFIG.SHEETS.DEALER_TRADE:
      soldCol = 6;
      deliveryCol = 8;
      priorityCol = 10;
      break;
    case CONFIG.SHEETS.NEW_CAR_PARTS:
      soldCol = 8;
      deliveryCol = 10;
      priorityCol = 11;
      break;
    default:
      return;
  }

  const isSold = sheet.getRange(row, soldCol).getValue();
  const deliveryDate = sheet.getRange(row, deliveryCol).getValue();

  if (isSold === 'Yes' && deliveryDate) {
    const now = new Date();
    const delivery = new Date(deliveryDate);
    const hoursUntilDelivery = (delivery - now) / (1000 * 60 * 60);

    let priority = 'Normal';
    if (hoursUntilDelivery <= CONFIG.URGENT_HOURS) {
      priority = 'URGENT';
    } else if (hoursUntilDelivery <= CONFIG.HIGH_PRIORITY_HOURS) {
      priority = 'HIGH';
    }

    sheet.getRange(row, priorityCol).setValue(priority);
  } else {
    sheet.getRange(row, priorityCol).setValue('Normal');
  }
}

/**
 * Handle status change
 */
function handleStatusChange(sheet, row, sheetName, newStatus, oldStatus) {
  // If status changed to Completed, set completion date
  if (newStatus === 'Completed' || newStatus === 'Sold' || newStatus === 'Lost') {
    const completionCol = getCompletionDateColumn(sheetName);
    if (completionCol && sheet.getRange(row, completionCol).getValue() === '') {
      sheet.getRange(row, completionCol).setValue(new Date());
    }

    // Send completion notification
    sendCompletionNotification(sheet, row, sheetName);
  }
}

/**
 * Get completion date column for each sheet
 */
function getCompletionDateColumn(sheetName) {
  const cols = {
    [CONFIG.SHEETS.DEALER_TRADE]: 13,
    [CONFIG.SHEETS.ACCESSORY_INSTALLS]: null, // No completion date column
    [CONFIG.SHEETS.NEW_CAR_PARTS]: null,
    [CONFIG.SHEETS.SERVICE_APPRAISALS]: null
  };
  return cols[sheetName];
}

// =============================================================================
// EMAIL NOTIFICATIONS
// =============================================================================

/**
 * Send notification for new entry
 */
function sendNewEntryNotification(sheet, row, sheetName, submitterEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss.getUrl();

  let recipient, subject, body;
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  switch (sheetName) {
    case CONFIG.SHEETS.DEALER_TRADE:
      recipient = CONFIG.DAN_EMAIL;
      subject = 'New Dealer Trade Re-PDI Request';
      body = formatDealerTradeEmail(rowData, url);
      break;

    case CONFIG.SHEETS.ACCESSORY_INSTALLS:
      // Only notify if requires tech install
      if (rowData[11] === 'Yes') { // Requires Tech Install column
        recipient = CONFIG.DAN_EMAIL;
        subject = 'New Accessory Install - Tech Work Needed';
        body = formatAccessoryInstallEmail(rowData, url);
      } else {
        return; // Don't send email
      }
      break;

    case CONFIG.SHEETS.NEW_CAR_PARTS:
      recipient = CONFIG.DAN_EMAIL;
      subject = 'New Car Parts Installation Request';
      body = formatNewCarPartsEmail(rowData, url);
      break;

    case CONFIG.SHEETS.SERVICE_APPRAISALS:
      recipient = CONFIG.BRIAN_EMAIL;
      subject = getAppraisalSubject(rowData[8]); // Heat Level
      body = formatAppraisalEmail(rowData, url);
      break;

    default:
      return;
  }

  if (recipient && subject && body) {
    sendEmail(recipient, subject, body);
  }
}

/**
 * Send completion notification to original submitter
 */
function sendCompletionNotification(sheet, row, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss.getUrl();

  const submitterEmail = sheet.getRange(row, 2).getValue(); // Submitted By column
  if (!submitterEmail) return;

  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  let subject, body;

  switch (sheetName) {
    case CONFIG.SHEETS.DEALER_TRADE:
      subject = 'Dealer Trade Re-PDI Completed - Stock #' + rowData[2];
      body = `
        <h2>Dealer Trade Re-PDI Completed</h2>
        <p>The following item has been marked as completed:</p>
        <ul>
          <li><strong>Stock Number:</strong> ${rowData[2]}</li>
          <li><strong>Vehicle:</strong> ${rowData[4]}</li>
          <li><strong>Customer:</strong> ${rowData[6] || 'N/A'}</li>
        </ul>
        <p><a href="${url}">View Dashboard</a></p>
      `;
      break;

    case CONFIG.SHEETS.ACCESSORY_INSTALLS:
      subject = 'Accessory Install Completed - ' + rowData[2];
      body = `
        <h2>Accessory Install Completed</h2>
        <p>The following accessory install has been completed:</p>
        <ul>
          <li><strong>Customer:</strong> ${rowData[2]}</li>
          <li><strong>Part:</strong> ${rowData[8]}</li>
        </ul>
        <p><a href="${url}">View Dashboard</a></p>
      `;
      break;

    case CONFIG.SHEETS.NEW_CAR_PARTS:
      subject = 'New Car Parts Installation Completed - Stock #' + rowData[2];
      body = `
        <h2>Parts Installation Completed</h2>
        <p>The following parts installation has been completed:</p>
        <ul>
          <li><strong>Stock Number:</strong> ${rowData[2]}</li>
          <li><strong>Vehicle:</strong> ${rowData[4]}</li>
          <li><strong>Part:</strong> ${rowData[6]}</li>
        </ul>
        <p><a href="${url}">View Dashboard</a></p>
      `;
      break;

    case CONFIG.SHEETS.SERVICE_APPRAISALS:
      const status = rowData[10];
      subject = `Service Drive Lead ${status} - ${rowData[2]}`;
      body = `
        <h2>Service Drive Lead Update</h2>
        <p>A service drive appraisal lead has been marked as <strong>${status}</strong>:</p>
        <ul>
          <li><strong>Customer:</strong> ${rowData[2]}</li>
          <li><strong>Vehicle:</strong> ${rowData[5]}</li>
          <li><strong>Heat Level:</strong> ${rowData[8]}</li>
        </ul>
        <p><a href="${url}">View Dashboard</a></p>
      `;
      break;

    default:
      return;
  }

  if (subject && body) {
    sendEmail(submitterEmail, subject, body);
  }
}

/**
 * Format Dealer Trade email
 */
function formatDealerTradeEmail(data, url) {
  const priority = data[9];
  const priorityStyle = priority === 'URGENT' ? 'background-color: #FF4444; color: white; padding: 5px;' : '';

  return `
    <div style="font-family: Arial, sans-serif; max-width: 600px;">
      <h2 style="color: #003366;">New Dealer Trade Re-PDI Request</h2>
      ${priority === 'URGENT' ? '<p style="' + priorityStyle + '"><strong>URGENT - DELIVERY WITHIN 24 HOURS</strong></p>' : ''}
      <table style="border-collapse: collapse; width: 100%;">
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Stock Number:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[2]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>VIN:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[3]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Vehicle:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[4]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>SOLD:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[5]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Customer:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[6] || 'N/A'}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Delivery Date:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[7] || 'N/A'}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Delivery Time:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[8] || 'N/A'}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Priority:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[9]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Notes:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[13] || ''}</td></tr>
      </table>
      <p style="margin-top: 20px;"><a href="${url}" style="background-color: #003366; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Open Dashboard</a></p>
      <p style="color: #666; font-size: 12px;">This is an automated notification from the Service-to-Sales Bridge Dashboard.</p>
    </div>
  `;
}

/**
 * Format Accessory Install email
 */
function formatAccessoryInstallEmail(data, url) {
  return `
    <div style="font-family: Arial, sans-serif; max-width: 600px;">
      <h2 style="color: #003366;">New Accessory Install - Tech Work Required</h2>
      <table style="border-collapse: collapse; width: 100%;">
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Customer:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[2]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Phone:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[3]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Email:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[4]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Vehicle:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[5]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Stock/VIN:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[6]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Part Number(s):</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[7]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Part Description:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[8]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Part Received:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[10]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Notes:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[14] || ''}</td></tr>
      </table>
      <p style="margin-top: 20px;"><a href="${url}" style="background-color: #003366; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Open Dashboard</a></p>
    </div>
  `;
}

/**
 * Format New Car Parts email
 */
function formatNewCarPartsEmail(data, url) {
  const priority = data[10];
  const priorityStyle = priority === 'URGENT' ? 'background-color: #FF4444; color: white; padding: 5px;' : '';

  return `
    <div style="font-family: Arial, sans-serif; max-width: 600px;">
      <h2 style="color: #003366;">New Car Parts Installation Request</h2>
      ${priority === 'URGENT' ? '<p style="' + priorityStyle + '"><strong>URGENT - DELIVERY WITHIN 24 HOURS</strong></p>' : ''}
      <table style="border-collapse: collapse; width: 100%;">
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Stock Number:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[2]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>VIN:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[3]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Vehicle:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[4]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Part Number(s):</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[5]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Part Description:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[6]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>SOLD:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[7]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Customer:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[8] || 'N/A'}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Delivery Date:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[9] || 'N/A'}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Notes:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[12] || ''}</td></tr>
      </table>
      <p style="margin-top: 20px;"><a href="${url}" style="background-color: #003366; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Open Dashboard</a></p>
    </div>
  `;
}

/**
 * Get subject line based on heat level
 */
function getAppraisalSubject(heatLevel) {
  switch (heatLevel) {
    case 'Hot':
      return 'HOT SERVICE DRIVE LEAD - Immediate Attention Needed';
    case 'Warm':
      return 'Warm Service Drive Lead - Follow Up Soon';
    case 'Cold':
      return 'Service Drive Lead - When Available';
    default:
      return 'New Service Drive Appraisal Opportunity';
  }
}

/**
 * Format Appraisal email
 */
function formatAppraisalEmail(data, url) {
  const heatLevel = data[8];
  let heatStyle, heatIcon;

  switch (heatLevel) {
    case 'Hot':
      heatStyle = 'background-color: #FF0000; color: white;';
      heatIcon = 'HOT';
      break;
    case 'Warm':
      heatStyle = 'background-color: #FFC107; color: black;';
      heatIcon = 'WARM';
      break;
    case 'Cold':
      heatStyle = 'background-color: #2196F3; color: white;';
      heatIcon = 'COLD';
      break;
    default:
      heatStyle = '';
      heatIcon = '';
  }

  return `
    <div style="font-family: Arial, sans-serif; max-width: 600px;">
      <h2 style="color: #003366;">Service Drive Appraisal Opportunity</h2>
      <p style="${heatStyle} padding: 10px; font-size: 18px; text-align: center;"><strong>${heatIcon} LEAD</strong></p>
      <table style="border-collapse: collapse; width: 100%;">
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Customer:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[2]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Phone:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[3]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Email:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[4]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Their Vehicle:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[5]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Mileage:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[6]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Service Being Performed:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[7]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Reason:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[9]}</td></tr>
        <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Notes:</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${data[13] || ''}</td></tr>
      </table>
      <p style="margin-top: 20px;"><a href="${url}" style="background-color: #003366; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Open Dashboard</a></p>
    </div>
  `;
}

/**
 * Send email helper function
 */
function sendEmail(to, subject, htmlBody) {
  try {
    MailApp.sendEmail({
      to: to,
      subject: '[Union Park] ' + subject,
      htmlBody: htmlBody
    });
    Logger.log('Email sent to: ' + to + ' | Subject: ' + subject);
  } catch (error) {
    Logger.log('Error sending email: ' + error.message);
  }
}

// =============================================================================
// DASHBOARD REFRESH
// =============================================================================

/**
 * Refresh all dashboard counts and priority items
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);

  if (!dashboard) return;

  // Update timestamp
  dashboard.getRange('A3').setValue('Last Updated: ' + new Date().toLocaleString());

  // Get counts for each category
  const dealerTradeCounts = getCounts(ss, CONFIG.SHEETS.DEALER_TRADE, 11);
  const accessoryCounts = getCounts(ss, CONFIG.SHEETS.ACCESSORY_INSTALLS, 13);
  const newCarPartsCounts = getCounts(ss, CONFIG.SHEETS.NEW_CAR_PARTS, 12);
  const appraisalCounts = getAppraisalCounts(ss);

  // Update summary section
  dashboard.getRange('B14').setValue(dealerTradeCounts.pending);
  dashboard.getRange('C14').setValue(dealerTradeCounts.scheduled);
  dashboard.getRange('D14').setValue(dealerTradeCounts.completedToday);

  dashboard.getRange('B15').setValue(accessoryCounts.pending);
  dashboard.getRange('C15').setValue(accessoryCounts.scheduled);
  dashboard.getRange('D15').setValue(accessoryCounts.completedToday);

  dashboard.getRange('B16').setValue(newCarPartsCounts.pending);
  dashboard.getRange('C16').setValue(newCarPartsCounts.scheduled);
  dashboard.getRange('D16').setValue(newCarPartsCounts.completedToday);

  // For appraisals: Hot / Warm / Cold (Open)
  dashboard.getRange('B17').setValue(appraisalCounts.hot + ' Hot');
  dashboard.getRange('C17').setValue(appraisalCounts.warm + ' Warm');
  dashboard.getRange('D17').setValue(appraisalCounts.cold + ' Cold');

  // Update priority section with sold units delivering within 48 hours
  updatePrioritySection(ss, dashboard);
}

/**
 * Get counts for a sheet
 */
function getCounts(ss, sheetName, statusCol) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    return { pending: 0, scheduled: 0, completedToday: 0 };
  }

  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let pending = 0, scheduled = 0, completedToday = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusCol - 1];
    const dateSubmitted = new Date(data[i][0]);

    if (status === 'Pending' || status === 'Part Ordered') {
      pending++;
    } else if (status === 'Scheduled' || status === 'In Progress' || status === 'Appointment Scheduled' || status === 'Parts Received') {
      scheduled++;
    } else if (status === 'Completed') {
      // Check if completed today
      const completionDate = data[i][getCompletionDateColumn(sheetName) - 1];
      if (completionDate) {
        const compDate = new Date(completionDate);
        compDate.setHours(0, 0, 0, 0);
        if (compDate.getTime() === today.getTime()) {
          completedToday++;
        }
      }
    }
  }

  return { pending, scheduled, completedToday };
}

/**
 * Get appraisal counts by heat level
 */
function getAppraisalCounts(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SERVICE_APPRAISALS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { hot: 0, warm: 0, cold: 0 };
  }

  const data = sheet.getDataRange().getValues();
  let hot = 0, warm = 0, cold = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][10]; // Status column (K)
    const heatLevel = data[i][8]; // Heat Level column (I)

    // Only count open leads
    if (status !== 'Sold' && status !== 'Lost') {
      if (heatLevel === 'Hot') hot++;
      else if (heatLevel === 'Warm') warm++;
      else if (heatLevel === 'Cold') cold++;
    }
  }

  return { hot, warm, cold };
}

/**
 * Update priority section with sold units delivering within 48 hours
 */
function updatePrioritySection(ss, dashboard) {
  // Clear existing priority data (rows 7-11)
  dashboard.getRange('A7:E11').clearContent();

  const priorityItems = [];
  const now = new Date();
  const cutoff = new Date(now.getTime() + (48 * 60 * 60 * 1000));

  // Check Dealer Trade sheet
  const dealerSheet = ss.getSheetByName(CONFIG.SHEETS.DEALER_TRADE);
  if (dealerSheet && dealerSheet.getLastRow() > 1) {
    const data = dealerSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] === 'Yes' && data[i][10] !== 'Completed') { // SOLD = Yes and not completed
        const deliveryDate = new Date(data[i][7]);
        if (deliveryDate <= cutoff && deliveryDate >= now) {
          priorityItems.push({
            stock: data[i][2],
            vehicle: data[i][4],
            customer: data[i][6],
            delivery: deliveryDate,
            status: data[i][10],
            type: 'Re-PDI'
          });
        }
      }
    }
  }

  // Check New Car Parts sheet
  const partsSheet = ss.getSheetByName(CONFIG.SHEETS.NEW_CAR_PARTS);
  if (partsSheet && partsSheet.getLastRow() > 1) {
    const data = partsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === 'Yes' && data[i][11] !== 'Completed') { // SOLD = Yes and not completed
        const deliveryDate = new Date(data[i][9]);
        if (deliveryDate <= cutoff && deliveryDate >= now) {
          priorityItems.push({
            stock: data[i][2],
            vehicle: data[i][4],
            customer: data[i][8],
            delivery: deliveryDate,
            status: data[i][11],
            type: 'Parts'
          });
        }
      }
    }
  }

  // Sort by delivery date
  priorityItems.sort((a, b) => a.delivery - b.delivery);

  // Display up to 5 priority items
  const displayItems = priorityItems.slice(0, 5);
  for (let i = 0; i < displayItems.length; i++) {
    const item = displayItems[i];
    const row = 7 + i;
    dashboard.getRange(row, 1).setValue(item.stock);
    dashboard.getRange(row, 2).setValue(item.vehicle);
    dashboard.getRange(row, 3).setValue(item.customer || 'N/A');
    dashboard.getRange(row, 4).setValue(Utilities.formatDate(item.delivery, Session.getScriptTimeZone(), 'MM/dd h:mm a'));
    dashboard.getRange(row, 5).setValue(item.status + ' (' + item.type + ')');

    // Color based on urgency
    const hoursUntil = (item.delivery - now) / (1000 * 60 * 60);
    if (hoursUntil <= 24) {
      dashboard.getRange(row, 1, 1, 5).setBackground(CONFIG.COLORS.URGENT_RED).setFontColor(CONFIG.COLORS.WHITE);
    } else {
      dashboard.getRange(row, 1, 1, 5).setBackground(CONFIG.COLORS.WARNING_ORANGE);
    }
  }

  if (displayItems.length === 0) {
    dashboard.getRange('A7').setValue('No urgent deliveries in the next 48 hours');
    dashboard.getRange('A7:E7').merge().setFontStyle('italic').setFontColor('#666666');
  }
}

// =============================================================================
// ARCHIVE AUTOMATION
// =============================================================================

/**
 * Archive completed items older than 7 days
 */
function archiveCompletedItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE);

  if (!archiveSheet) return;

  const sheetsToCheck = [
    { name: CONFIG.SHEETS.DEALER_TRADE, statusCol: 11, completedStatus: 'Completed' },
    { name: CONFIG.SHEETS.ACCESSORY_INSTALLS, statusCol: 13, completedStatus: 'Completed' },
    { name: CONFIG.SHEETS.NEW_CAR_PARTS, statusCol: 12, completedStatus: 'Completed' },
    { name: CONFIG.SHEETS.SERVICE_APPRAISALS, statusCol: 11, completedStatus: ['Sold', 'Lost'] }
  ];

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - CONFIG.ARCHIVE_DAYS);

  sheetsToCheck.forEach(config => {
    archiveFromSheet(ss, archiveSheet, config, cutoffDate);
  });
}

/**
 * Archive completed items from a specific sheet
 */
function archiveFromSheet(ss, archiveSheet, config, cutoffDate) {
  const sheet = ss.getSheetByName(config.name);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const status = data[i][config.statusCol - 1];
    const dateSubmitted = new Date(data[i][0]);

    const isCompleted = Array.isArray(config.completedStatus)
      ? config.completedStatus.includes(status)
      : status === config.completedStatus;

    if (isCompleted && dateSubmitted < cutoffDate) {
      // Add to archive
      const archiveRow = [
        new Date(), // Archive Date
        config.name, // Original Sheet
        data[i][0], // Date Submitted
        data[i][1], // Submitted By
        data[i][2], // Stock/Customer
        data[i][4] || data[i][5] || '', // Description
        status, // Status
        data[i][getCompletionDateColumn(config.name) - 1] || '', // Completion Date
        JSON.stringify(data[i]) // Full data as JSON
      ];

      archiveSheet.appendRow(archiveRow);
      rowsToDelete.push(i + 1); // 1-indexed row number
    }
  }

  // Delete archived rows (in reverse order to maintain indices)
  rowsToDelete.forEach(row => {
    sheet.deleteRow(row);
  });

  if (rowsToDelete.length > 0) {
    Logger.log(`Archived ${rowsToDelete.length} items from ${config.name}`);
  }
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Get user's display name from email
 */
function getUserName(email) {
  return CONFIG.NAMES[email] || email;
}

/**
 * Remove all sheet protections (run this if sheets are locked)
 */
function removeAllProtections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(p => p.remove());

    const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    rangeProtections.forEach(p => p.remove());
  });

  SpreadsheetApp.getUi().alert('All locks have been removed. Sheets are now fully editable.');
}

/**
 * Manual test function for email notifications
 */
function testEmailNotification() {
  const testEmail = Session.getActiveUser().getEmail();
  sendEmail(
    testEmail,
    'Test Email from Bridge Dashboard',
    '<h2>Test Email</h2><p>This is a test email from the Service-to-Sales Bridge Dashboard.</p>'
  );
  SpreadsheetApp.getUi().alert('Test email sent to: ' + testEmail);
}

// =============================================================================
// QUICK ENTRY FORMS
// =============================================================================

/**
 * Show Dealer Trade Re-PDI form
 */
function showDealerTradeForm() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      .form-group { margin-bottom: 12px; }
      label { display: block; font-weight: bold; margin-bottom: 4px; color: #003366; }
      input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
      .row { display: flex; gap: 10px; }
      .row .form-group { flex: 1; }
      button { background: #003366; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-top: 10px; }
      button:hover { background: #004488; }
      h2 { color: #003366; margin-top: 0; }
      .sold-section { background: #FFF9C4; padding: 10px; border-radius: 4px; margin-top: 10px; display: none; }
      .urgent-box { background: #FF4444; color: white; padding: 8px; border-radius: 4px; margin-bottom: 12px; }
      .urgent-box label { color: white; display: inline; margin-left: 8px; }
      .urgent-box input { width: auto; }
    </style>
    <h2>Add Dealer Trade Re-PDI</h2>
    <form id="form">
      <div class="urgent-box">
        <input type="checkbox" id="urgent">
        <label for="urgent">URGENT / 911 - Needs immediate attention!</label>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Stock Number *</label>
          <input type="text" id="stock" required>
        </div>
        <div class="form-group">
          <label>VIN (Last 8)</label>
          <input type="text" id="vin" maxlength="17">
        </div>
      </div>
      <div class="form-group">
        <label>Year/Make/Model *</label>
        <input type="text" id="vehicle" placeholder="2024 GMC Sierra 1500" required>
      </div>
      <div class="form-group">
        <label>Is it SOLD?</label>
        <select id="sold" onchange="toggleSold()">
          <option value="No">No</option>
          <option value="Yes">Yes</option>
        </select>
      </div>
      <div class="sold-section" id="soldSection">
        <div class="form-group">
          <label>Customer Name</label>
          <input type="text" id="customer">
        </div>
        <div class="row">
          <div class="form-group">
            <label>Delivery Date</label>
            <input type="date" id="deliveryDate">
          </div>
          <div class="form-group">
            <label>Delivery Time</label>
            <input type="time" id="deliveryTime">
          </div>
        </div>
      </div>
      <div class="form-group">
        <label>Notes</label>
        <input type="text" id="notes" placeholder="Any special instructions...">
      </div>
      <button type="submit">Add Entry</button>
    </form>
    <script>
      function toggleSold() {
        document.getElementById('soldSection').style.display =
          document.getElementById('sold').value === 'Yes' ? 'block' : 'none';
      }
      document.getElementById('form').addEventListener('submit', function(e) {
        e.preventDefault();
        const data = {
          urgent: document.getElementById('urgent').checked,
          stock: document.getElementById('stock').value,
          vin: document.getElementById('vin').value,
          vehicle: document.getElementById('vehicle').value,
          sold: document.getElementById('sold').value,
          customer: document.getElementById('customer').value,
          deliveryDate: document.getElementById('deliveryDate').value,
          deliveryTime: document.getElementById('deliveryTime').value,
          notes: document.getElementById('notes').value
        };
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).addDealerTradeEntry(data);
      });
    </script>
  `)
  .setWidth(450)
  .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Dealer Trade Re-PDI');
}

/**
 * Add dealer trade entry from form
 */
function addDealerTradeEntry(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DEALER_TRADE);

    if (!sheet) {
      throw new Error('Sheet "Dealer Trade Re-PDIs" not found. Run Setup Sheet Structure first.');
    }

    const userEmail = Session.getActiveUser().getEmail();
    const userName = getUserName(userEmail);
    const now = new Date();

    // Calculate priority - URGENT checkbox overrides everything
    let priority = 'Normal';
    if (data.urgent) {
      priority = 'URGENT';
    } else if (data.sold === 'Yes' && data.deliveryDate) {
      const delivery = new Date(data.deliveryDate);
      const hoursUntil = (delivery - now) / (1000 * 60 * 60);
      if (hoursUntil <= 24) priority = 'URGENT';
      else if (hoursUntil <= 48) priority = 'HIGH';
    }

    const row = [
      now,                    // Date Submitted
      userName,               // Submitted By (NAME not email)
      data.stock,             // Stock Number
      data.vin || '',         // VIN
      data.vehicle,           // Year/Make/Model
      data.sold,              // SOLD?
      data.customer || '',    // Customer Name
      data.deliveryDate || '',// Delivery Date
      data.deliveryTime || '',// Delivery Time
      priority,               // Priority Flag
      'Pending',              // Status
      '',                     // Assigned To
      '',                     // Completion Date
      data.notes || ''        // Notes
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush(); // Force save

    // Send notification (don't let this break the save)
    try {
      sendNewEntryNotification(sheet, sheet.getLastRow(), CONFIG.SHEETS.DEALER_TRADE, userEmail);
    } catch (e) {
      Logger.log('Notification error: ' + e.message);
    }

    refreshDashboard();
    SpreadsheetApp.getActive().toast('Dealer Trade Re-PDI added!', 'Success', 3);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    throw error;
  }
}

/**
 * Show Customer Accessory Install form
 */
function showAccessoryForm() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      .form-group { margin-bottom: 12px; }
      label { display: block; font-weight: bold; margin-bottom: 4px; color: #003366; }
      input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
      .row { display: flex; gap: 10px; }
      .row .form-group { flex: 1; }
      button { background: #003366; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-top: 10px; }
      button:hover { background: #004488; }
      h2 { color: #003366; margin-top: 0; }
      .urgent-box { background: #FF4444; color: white; padding: 8px; border-radius: 4px; margin-bottom: 12px; }
      .urgent-box label { color: white; display: inline; margin-left: 8px; }
      .urgent-box input { width: auto; }
    </style>
    <h2>Add Customer Accessory Install</h2>
    <form id="form">
      <div class="urgent-box">
        <input type="checkbox" id="urgent">
        <label for="urgent">URGENT / 911 - Needs immediate attention!</label>
      </div>
      <div class="form-group">
        <label>Customer Name *</label>
        <input type="text" id="customer" required>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Phone</label>
          <input type="tel" id="phone">
        </div>
        <div class="form-group">
          <label>Email</label>
          <input type="email" id="email">
        </div>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Vehicle *</label>
          <input type="text" id="vehicle" placeholder="2024 GMC Acadia" required>
        </div>
        <div class="form-group">
          <label>Stock # or VIN</label>
          <input type="text" id="stock">
        </div>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Part Number(s)</label>
          <input type="text" id="partNum">
        </div>
        <div class="form-group">
          <label>Part Description *</label>
          <input type="text" id="partDesc" required>
        </div>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Part Ordered?</label>
          <select id="ordered"><option value="No">No</option><option value="Yes">Yes</option></select>
        </div>
        <div class="form-group">
          <label>Part Received?</label>
          <select id="received"><option value="No">No</option><option value="Yes">Yes</option></select>
        </div>
        <div class="form-group">
          <label>Needs Tech Install?</label>
          <select id="techInstall"><option value="Yes">Yes</option><option value="No">No</option></select>
        </div>
      </div>
      <div class="form-group">
        <label>Notes</label>
        <input type="text" id="notes">
      </div>
      <button type="submit">Add Entry</button>
    </form>
    <script>
      document.getElementById('form').addEventListener('submit', function(e) {
        e.preventDefault();
        const data = {
          urgent: document.getElementById('urgent').checked,
          customer: document.getElementById('customer').value,
          phone: document.getElementById('phone').value,
          email: document.getElementById('email').value,
          vehicle: document.getElementById('vehicle').value,
          stock: document.getElementById('stock').value,
          partNum: document.getElementById('partNum').value,
          partDesc: document.getElementById('partDesc').value,
          ordered: document.getElementById('ordered').value,
          received: document.getElementById('received').value,
          techInstall: document.getElementById('techInstall').value,
          notes: document.getElementById('notes').value
        };
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).addAccessoryEntry(data);
      });
    </script>
  `)
  .setWidth(500)
  .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Customer Accessory Install');
}

/**
 * Add accessory entry from form
 */
function addAccessoryEntry(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ACCESSORY_INSTALLS);

    if (!sheet) {
      throw new Error('Sheet "Customer Accessory Installs" not found. Run Setup Sheet Structure first.');
    }

    const userEmail = Session.getActiveUser().getEmail();
    const userName = getUserName(userEmail);
    const now = new Date();

    // Add URGENT to notes if checked
    let notes = data.notes || '';
    if (data.urgent) {
      notes = 'URGENT/911 - ' + notes;
    }

    const row = [
      now,                    // Date Submitted
      userName,               // Submitted By (NAME not email)
      data.customer,          // Customer Name
      data.phone || '',       // Phone
      data.email || '',       // Email
      data.vehicle,           // Vehicle
      data.stock || '',       // Stock/VIN
      data.partNum || '',     // Part Number
      data.partDesc,          // Part Description
      data.ordered,           // Part Ordered?
      data.received,          // Part Received?
      data.techInstall,       // Requires Tech?
      'Part Ordered',         // Status
      '',                     // Appointment
      notes                   // Notes
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush(); // Force save

    // Send notification if tech install needed OR urgent
    if (data.techInstall === 'Yes' || data.urgent) {
      try {
        sendNewEntryNotification(sheet, sheet.getLastRow(), CONFIG.SHEETS.ACCESSORY_INSTALLS, userEmail);
      } catch (e) {
        Logger.log('Notification error: ' + e.message);
      }
    }
    refreshDashboard();

    SpreadsheetApp.getActive().toast('Accessory Install added!', 'Success', 3);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    throw error;
  }
}

/**
 * Show New Car Parts form
 */
function showPartsForm() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      .form-group { margin-bottom: 12px; }
      label { display: block; font-weight: bold; margin-bottom: 4px; color: #003366; }
      input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
      .row { display: flex; gap: 10px; }
      .row .form-group { flex: 1; }
      button { background: #003366; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-top: 10px; }
      button:hover { background: #004488; }
      h2 { color: #003366; margin-top: 0; }
      .sold-section { background: #FFF9C4; padding: 10px; border-radius: 4px; margin-top: 10px; display: none; }
      .urgent-box { background: #FF4444; color: white; padding: 8px; border-radius: 4px; margin-bottom: 12px; }
      .urgent-box label { color: white; display: inline; margin-left: 8px; }
      .urgent-box input { width: auto; }
    </style>
    <h2>Add New Car Parts Install</h2>
    <form id="form">
      <div class="urgent-box">
        <input type="checkbox" id="urgent">
        <label for="urgent">URGENT / 911 - Needs immediate attention!</label>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Stock Number *</label>
          <input type="text" id="stock" required>
        </div>
        <div class="form-group">
          <label>VIN</label>
          <input type="text" id="vin">
        </div>
      </div>
      <div class="form-group">
        <label>Year/Make/Model *</label>
        <input type="text" id="vehicle" placeholder="2024 Buick Enclave" required>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Part Number(s)</label>
          <input type="text" id="partNum">
        </div>
        <div class="form-group">
          <label>Part Description *</label>
          <input type="text" id="partDesc" required>
        </div>
      </div>
      <div class="form-group">
        <label>Is it SOLD?</label>
        <select id="sold" onchange="toggleSold()">
          <option value="No">No</option>
          <option value="Yes">Yes</option>
        </select>
      </div>
      <div class="sold-section" id="soldSection">
        <div class="form-group">
          <label>Customer Name</label>
          <input type="text" id="customer">
        </div>
        <div class="form-group">
          <label>Delivery Date</label>
          <input type="date" id="deliveryDate">
        </div>
      </div>
      <div class="form-group">
        <label>Notes</label>
        <input type="text" id="notes">
      </div>
      <button type="submit">Add Entry</button>
    </form>
    <script>
      function toggleSold() {
        document.getElementById('soldSection').style.display =
          document.getElementById('sold').value === 'Yes' ? 'block' : 'none';
      }
      document.getElementById('form').addEventListener('submit', function(e) {
        e.preventDefault();
        const data = {
          urgent: document.getElementById('urgent').checked,
          stock: document.getElementById('stock').value,
          vin: document.getElementById('vin').value,
          vehicle: document.getElementById('vehicle').value,
          partNum: document.getElementById('partNum').value,
          partDesc: document.getElementById('partDesc').value,
          sold: document.getElementById('sold').value,
          customer: document.getElementById('customer').value,
          deliveryDate: document.getElementById('deliveryDate').value,
          notes: document.getElementById('notes').value
        };
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).addPartsEntry(data);
      });
    </script>
  `)
  .setWidth(450)
  .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Car Parts Install');
}

/**
 * Add parts entry from form
 */
function addPartsEntry(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.NEW_CAR_PARTS);

    if (!sheet) {
      throw new Error('Sheet "New Car Parts Installation" not found. Run Setup Sheet Structure first.');
    }

    const userEmail = Session.getActiveUser().getEmail();
    const userName = getUserName(userEmail);
    const now = new Date();

    // Calculate priority - URGENT checkbox overrides everything
    let priority = 'Normal';
    if (data.urgent) {
      priority = 'URGENT';
    } else if (data.sold === 'Yes' && data.deliveryDate) {
      const delivery = new Date(data.deliveryDate);
      const hoursUntil = (delivery - now) / (1000 * 60 * 60);
      if (hoursUntil <= 24) priority = 'URGENT';
      else if (hoursUntil <= 48) priority = 'HIGH';
    }

    const row = [
      now,                    // Date Submitted
      userName,               // Submitted By (NAME not email)
      data.stock,             // Stock Number
      data.vin || '',         // VIN
      data.vehicle,           // Year/Make/Model
      data.partNum || '',     // Part Number
      data.partDesc,          // Part Description
      data.sold,              // SOLD?
      data.customer || '',    // Customer Name
      data.deliveryDate || '',// Delivery Date
      priority,               // Priority Flag
      'Pending',              // Status
      data.notes || ''        // Notes
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush(); // Force save

    try {
      sendNewEntryNotification(sheet, sheet.getLastRow(), CONFIG.SHEETS.NEW_CAR_PARTS, userEmail);
    } catch (e) {
      Logger.log('Notification error: ' + e.message);
    }
    refreshDashboard();

    SpreadsheetApp.getActive().toast('Parts Install added!', 'Success', 3);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    throw error;
  }
}

/**
 * Show Service Drive Appraisal form
 */
function showAppraisalForm() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      .form-group { margin-bottom: 12px; }
      label { display: block; font-weight: bold; margin-bottom: 4px; color: #003366; }
      input, select, textarea { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
      .row { display: flex; gap: 10px; }
      .row .form-group { flex: 1; }
      button { background: #003366; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-top: 10px; }
      button:hover { background: #004488; }
      h2 { color: #003366; margin-top: 0; }
      .urgent-box { background: #FF4444; color: white; padding: 8px; border-radius: 4px; margin-bottom: 12px; }
      .urgent-box label { color: white; display: inline; margin-left: 8px; }
      .urgent-box input { width: auto; }
    </style>
    <h2>Add Service Drive Appraisal</h2>
    <form id="form">
      <div class="urgent-box">
        <input type="checkbox" id="urgent">
        <label for="urgent">URGENT / 911 - Customer waiting or hot lead!</label>
      </div>
      <div class="form-group">
        <label>Customer Name *</label>
        <input type="text" id="customer" required>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Phone *</label>
          <input type="tel" id="phone" required>
        </div>
        <div class="form-group">
          <label>Email</label>
          <input type="email" id="email">
        </div>
      </div>
      <div class="row">
        <div class="form-group">
          <label>Their Vehicle *</label>
          <input type="text" id="vehicle" placeholder="2019 Ford F-150" required>
        </div>
        <div class="form-group">
          <label>Mileage</label>
          <input type="text" id="mileage" placeholder="65,000">
        </div>
      </div>
      <div class="form-group">
        <label>Service Being Performed</label>
        <input type="text" id="service" placeholder="Oil change, brake inspection, etc.">
      </div>
      <div class="row">
        <div class="form-group">
          <label>Heat Level *</label>
          <select id="heat" required>
            <option value="Hot">HOT - Strong interest</option>
            <option value="Warm">WARM - Some interest</option>
            <option value="Cold">COLD - Worth following up</option>
          </select>
        </div>
      </div>
      <div class="form-group">
        <label>Why this heat level? *</label>
        <input type="text" id="reason" placeholder="Complained about payments, asked about new trucks..." required>
      </div>
      <div class="form-group">
        <label>Notes</label>
        <input type="text" id="notes">
      </div>
      <button type="submit">Add Entry</button>
    </form>
    <script>
      document.getElementById('form').addEventListener('submit', function(e) {
        e.preventDefault();
        const data = {
          urgent: document.getElementById('urgent').checked,
          customer: document.getElementById('customer').value,
          phone: document.getElementById('phone').value,
          email: document.getElementById('email').value,
          vehicle: document.getElementById('vehicle').value,
          mileage: document.getElementById('mileage').value,
          service: document.getElementById('service').value,
          heat: document.getElementById('heat').value,
          reason: document.getElementById('reason').value,
          notes: document.getElementById('notes').value
        };
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).addAppraisalEntry(data);
      });
    </script>
  `)
  .setWidth(450)
  .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Service Drive Appraisal');
}

/**
 * Add appraisal entry from form
 */
function addAppraisalEntry(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SERVICE_APPRAISALS);

    if (!sheet) {
      throw new Error('Sheet "Service Drive Appraisals" not found. Run Setup Sheet Structure first.');
    }

    const userEmail = Session.getActiveUser().getEmail();
    const userName = getUserName(userEmail);
    const now = new Date();

    // If URGENT, force heat level to Hot and prepend to notes
    let heatLevel = data.heat;
    let notes = data.notes || '';
    if (data.urgent) {
      heatLevel = 'Hot';
      notes = 'URGENT/911 - ' + notes;
    }

    const row = [
      now,                    // Date Submitted
      userName,               // Submitted By (NAME not email)
      data.customer,          // Customer Name
      data.phone,             // Phone
      data.email || '',       // Email
      data.vehicle,           // Vehicle
      data.mileage || '',     // Mileage
      data.service || '',     // Service Being Performed
      heatLevel,              // Heat Level (Hot if URGENT)
      data.reason,            // Reason
      'New Lead',             // Status
      '',                     // Assigned Salesperson
      '',                     // Follow-up Date
      notes                   // Notes (with URGENT prefix if checked)
    ];

    sheet.appendRow(row);
    SpreadsheetApp.flush(); // Force save

    try {
      sendNewEntryNotification(sheet, sheet.getLastRow(), CONFIG.SHEETS.SERVICE_APPRAISALS, userEmail);
    } catch (e) {
      Logger.log('Notification error: ' + e.message);
    }
    refreshDashboard();

    SpreadsheetApp.getActive().toast('Service Drive Appraisal added!', 'Success', 3);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    throw error;
  }
}
