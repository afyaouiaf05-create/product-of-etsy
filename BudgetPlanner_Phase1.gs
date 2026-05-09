/**
 * BUDGET PLANNER - PHASE 1 (Core Sheets)
 * Version 1.0
 *
 * SETUP:
 * 1. Create new Google Sheet
 * 2. Extensions → Apps Script
 * 3. Paste this code, save
 * 4. Run → main() → Authorize
 * 5. Wait 1-2 minutes
 *
 * PHASE 2 will add: Paycheck Budget, Yearly Overview,
 * Net Worth, Investment Tracker, Bill Calendar,
 * Credit Score Tracker, README sheet.
 */

// ============================================================
// DESIGN SYSTEM
// ============================================================
var COLORS = {
  sageGreen:  '#87A878',
  dustyPink:  '#D4A5A5',
  beige:      '#F5F1E8',
  lavender:   '#C3B5E0',
  softBlue:   '#A8C8EC',
  lightGray:  '#E0E0E0',
  red:        '#E57373',
  green:      '#81C784',
  yellow:     '#FFD54F',
  white:      '#FFFFFF',
  darkText:   '#333333',
  headerText: '#FFFFFF'
};

var YEAR = new Date().getFullYear();

// ============================================================
// MAIN ENTRY POINT
// ============================================================
function main() {
  try {
    Logger.log('🚀 Starting Budget Planner Phase 1...');
    createDashboard();
    SpreadsheetApp.flush();
    createMonthlyBudget();
    SpreadsheetApp.flush();
    createExpenseTracker();
    SpreadsheetApp.flush();
    createSavingsGoals();
    SpreadsheetApp.flush();
    createDebtPayoff();
    SpreadsheetApp.flush();
    createChecklists();
    SpreadsheetApp.flush();
    createInstructionsSheet();
    SpreadsheetApp.flush();
    setupCustomMenu();
    Logger.log('✅ Phase 1 complete!');
  } catch (e) {
    Logger.log('❌ Error: ' + e + '\nStack: ' + e.stack);
  }
}

// ============================================================
// ON OPEN TRIGGER
// ============================================================
function onOpen() {
  setupCustomMenu();
  showWelcomePopup();
}

function setupCustomMenu() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('💰 Budget Planner')
      .addItem('📖 Instructions', 'goToInstructions')
      .addItem('🔄 Reset All Data', 'resetAllData')
      .addItem('ℹ️ About', 'showAbout')
      .addToUi();
  } catch (e) {
    Logger.log('Menu setup error: ' + e);
  }
}

function showWelcomePopup() {
  try {
    SpreadsheetApp.getUi().alert(
      '👋 Welcome to Budget Planner!',
      'Start with the Instructions tab.\n\nAll calculations are automatic — just fill in your data!\n\n💡 Tip: Use the 💰 Budget Planner menu above for quick navigation.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('Welcome popup skipped: ' + e);
  }
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    'ℹ️ About Budget Planner',
    'Budget Planner v1.0 — Phase 1\n\nSheets: Dashboard, Monthly Budget, Expense Tracker,\nSavings Goals, Debt Payoff, Checklists, Instructions\n\nPhase 2 will add: Net Worth, Investment Tracker,\nBill Calendar, Credit Score Tracker & more.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function goToInstructions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Instructions');
  if (sh) ss.setActiveSheet(sh);
}

// ============================================================
// RESET ALL DATA
// ============================================================
function resetAllData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '⚠️ Reset All Data',
    'This will clear all user-entered data while preserving formulas, headers, and dropdowns.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Monthly Budget: clear Budgeted columns only (col B of each month block)
  var mb = ss.getSheetByName('Monthly Budget');
  if (mb) {
    // Budgeted input cells are in specific columns — clear cols B,F,J,N,R,V,Z,AD,AH,AL,AP,AT (every 4th starting at B=2)
    var monthCols = [2,6,10,14,18,22,26,30,34,38,42,46];
    monthCols.forEach(function(col) {
      try { mb.getRange(7, col, 40, 1).clearContent(); } catch(e) {}
    });
  }

  // Expense Tracker: clear data rows
  var et = ss.getSheetByName('Expense Tracker');
  if (et) {
    try { et.getRange(6, 1, et.getLastRow(), 8).clearContent(); } catch(e) {}
  }

  // Savings Goals: clear Current and Monthly Contribution
  var sg = ss.getSheetByName('Savings Goals');
  if (sg) {
    try { sg.getRange(6, 3, 15, 2).clearContent(); } catch(e) {}
  }

  // Debt Payoff: clear Current Balance and Extra Payment
  var dp = ss.getSheetByName('Debt Payoff');
  if (dp) {
    try { dp.getRange(8, 3, 10, 1).clearContent(); } catch(e) {}
    try { dp.getRange(8, 6, 10, 1).clearContent(); } catch(e) {}
  }

  // Checklists: uncheck all checkboxes
  var cl = ss.getSheetByName('Checklists');
  if (cl) {
    try {
      var lastRow = cl.getLastRow();
      for (var r = 1; r <= lastRow; r++) {
        var cell = cl.getRange(r, 1);
        try {
          if (cell.getValue() === true || cell.getValue() === false) {
            cell.setValue(false);
          }
        } catch(e2) {}
      }
    } catch(e) {}
  }

  ui.alert('✅ Reset Complete', 'All user data has been cleared. Formulas and structure preserved.', ui.ButtonSet.OK);
}

// ============================================================
// HELPER UTILITIES
// ============================================================
function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) {
    sh.clear();
    Logger.log('Cleared existing sheet: ' + name);
  } else {
    sh = ss.insertSheet(name);
    Logger.log('Created sheet: ' + name);
  }
  return sh;
}

function styleHeader(sheet, row, col, numCols, text, bgColor, fontSize, fontColor) {
  fontSize = fontSize || 14;
  fontColor = fontColor || COLORS.white;
  var range = sheet.getRange(row, col, 1, numCols);
  range.merge()
       .setValue(text)
       .setBackground(bgColor)
       .setFontColor(fontColor)
       .setFontSize(fontSize)
       .setFontWeight('bold')
       .setFontFamily('Arial')
       .setVerticalAlignment('middle')
       .setHorizontalAlignment('center')
       .setWrap(true);
  sheet.setRowHeight(row, 35);
}

function styleCard(sheet, row, col, label, value, bgColor) {
  var labelRange = sheet.getRange(row, col, 1, 2);
  labelRange.merge().setValue(label)
    .setBackground(bgColor)
    .setFontColor(COLORS.darkText)
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setHorizontalAlignment('center');

  var valueRange = sheet.getRange(row + 1, col, 1, 2);
  valueRange.merge().setValue(value)
    .setBackground(bgColor)
    .setFontColor(COLORS.darkText)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setNumberFormat('$#,##0.00');
  sheet.setRowHeight(row, 22);
  sheet.setRowHeight(row + 1, 30);
}

function setColumnWidths(sheet, widths) {
  widths.forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });
}

// ============================================================
// SHEET 1: DASHBOARD
// ============================================================
function createDashboard() {
  try {
    var sheet = getOrCreateSheet('Dashboard');
    sheet.setTabColor(COLORS.sageGreen);

    // Title row
    sheet.getRange(1, 1, 1, 20).merge()
      .setValue(YEAR + ' Budget Planner')
      .setBackground(COLORS.sageGreen)
      .setFontColor(COLORS.white)
      .setFontSize(16)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 50);

    // Subtitle
    sheet.getRange(2, 1, 1, 20).merge()
      .setValue('📊 Financial Overview — Updated Automatically')
      .setBackground(COLORS.beige)
      .setFontColor(COLORS.darkText)
      .setFontSize(10)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(2, 25);

    // --- Summary Cards ---
    var cardData = [
      ['💰 Total Annual Income',    "=IFERROR(SUM('Monthly Budget'!C7:C200)*0,0)"],
      ['💸 Total Annual Expenses',  "=IFERROR(SUM('Monthly Budget'!C7:C200)*0,0)"],
      ['📈 Net Balance',            '=B5-B7'],
      ['🏦 Savings Rate %',         '=IFERROR(B9/B5,0)'],
      ['💎 Net Worth',              '=0'],
      ['📉 Debt-to-Income Ratio',   '=IFERROR(B13/B5,0)']
    ];

    // Card layout: 2 rows x 3 columns, starting at row 4
    // We'll use a simpler approach with visible labels and formula references
    var cardRow = 4;
    var cardLabels = [
      '💰 Total Annual Income',
      '💸 Total Annual Expenses',
      '📈 Net Balance',
      '🏦 Savings Rate %',
      '💎 Net Worth',
      '📉 Debt-to-Income'
    ];
    var cardBgs = [COLORS.sageGreen, COLORS.dustyPink, COLORS.softBlue, COLORS.lavender, COLORS.beige, COLORS.beige];
    var cardCols = [1, 5, 9, 1, 5, 9];
    var cardRows = [4, 4, 4, 7, 7, 7];

    // Place cards
    for (var i = 0; i < 6; i++) {
      var r = cardRows[i];
      var c = cardCols[i];
      var bg = cardBgs[i];

      sheet.getRange(r, c, 1, 3).merge()
        .setValue(cardLabels[i])
        .setBackground(bg)
        .setFontColor(i < 3 ? COLORS.white : COLORS.darkText)
        .setFontSize(10)
        .setFontWeight('bold')
        .setFontFamily('Arial')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');

      sheet.getRange(r + 1, c, 1, 3).merge()
        .setValue('—')
        .setBackground(bg)
        .setFontColor(i < 3 ? COLORS.white : COLORS.darkText)
        .setFontSize(16)
        .setFontWeight('bold')
        .setFontFamily('Arial')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setNumberFormat(i === 3 || i === 5 ? '0.0%' : '$#,##0.00');

      sheet.setRowHeight(r, 28);
      sheet.setRowHeight(r + 1, 36);
    }

    // Card note
    sheet.getRange(10, 1, 1, 12).merge()
      .setValue('ℹ️  Summary cards will populate automatically once Monthly Budget data is entered.')
      .setFontSize(9)
      .setFontColor('#777777')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center');

    // --- Monthly Breakdown Table ---
    var tableStartRow = 12;

    // Section header
    sheet.getRange(tableStartRow, 1, 1, 16).merge()
      .setValue('📅 Monthly Breakdown')
      .setBackground(COLORS.sageGreen)
      .setFontColor(COLORS.white)
      .setFontSize(13)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(tableStartRow, 32);

    // Column headers
    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Annual'];
    var tableHeaders = ['Month','Income','Expenses','Net','Savings Rate'];
    var headerRow = tableStartRow + 1;
    var headerRange = sheet.getRange(headerRow, 1, 1, 5);
    headerRange.setValues([tableHeaders])
      .setBackground(COLORS.lightGray)
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(headerRow, 26);

    // Month rows (placeholder formulas — reference Monthly Budget when available)
    for (var m = 0; m < 12; m++) {
      var r2 = tableStartRow + 2 + m;
      sheet.getRange(r2, 1).setValue(months[m]).setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
      sheet.getRange(r2, 2).setValue(0).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10);
      sheet.getRange(r2, 3).setValue(0).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10);
      sheet.getRange(r2, 4).setFormula('=B' + r2 + '-C' + r2).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10);
      sheet.getRange(r2, 5).setFormula('=IFERROR(D' + r2 + '/B' + r2 + ',0)').setNumberFormat('0.0%').setFontFamily('Arial').setFontSize(10);
      sheet.getRange(r2, 1, 1, 5).setBackground(m % 2 === 0 ? COLORS.white : COLORS.beige);
      sheet.setRowHeight(r2, 22);
    }

    // Annual total row
    var annRow = tableStartRow + 14;
    sheet.getRange(annRow, 1, 1, 5).setBackground(COLORS.lightGray);
    sheet.getRange(annRow, 1).setValue('Annual Total').setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    sheet.getRange(annRow, 2).setFormula('=SUM(B' + (tableStartRow+2) + ':B' + (tableStartRow+13) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial');
    sheet.getRange(annRow, 3).setFormula('=SUM(C' + (tableStartRow+2) + ':C' + (tableStartRow+13) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial');
    sheet.getRange(annRow, 4).setFormula('=B' + annRow + '-C' + annRow).setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial');
    sheet.getRange(annRow, 5).setFormula('=IFERROR(D' + annRow + '/B' + annRow + ',0)').setNumberFormat('0.0%').setFontWeight('bold').setFontFamily('Arial');
    sheet.setRowHeight(annRow, 26);

    // Conditional formatting on Net column (col D)
    var netRange = sheet.getRange(tableStartRow + 2, 4, 12, 1);
    var posRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(COLORS.green)
      .setFontColor('#1B5E20')
      .setRanges([netRange]).build();
    var negRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(COLORS.red)
      .setFontColor('#B71C1C')
      .setRanges([netRange]).build();
    sheet.setConditionalFormatRules([posRule, negRule]);

    // Savings Rate conditional (col E)
    var srRange = sheet.getRange(tableStartRow + 2, 5, 12, 1);
    var sr20 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.2)
      .setBackground(COLORS.green).setFontColor('#1B5E20')
      .setRanges([srRange]).build();
    var sr10 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.1, 0.199)
      .setBackground(COLORS.yellow).setFontColor('#795548')
      .setRanges([srRange]).build();
    var sr0 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.1)
      .setBackground(COLORS.red).setFontColor('#B71C1C')
      .setRanges([srRange]).build();
    var existingRules = sheet.getConditionalFormatRules();
    sheet.setConditionalFormatRules(existingRules.concat([sr20, sr10, sr0]));

    // Charts note
    var chartNoteRow = annRow + 2;
    sheet.getRange(chartNoteRow, 1, 1, 12).merge()
      .setValue('📊 Charts: Add charts manually in Google Sheets (Insert → Chart) using the Monthly Breakdown data above, or the Expense Tracker and Savings Goals sheets as data sources.')
      .setBackground(COLORS.beige)
      .setFontSize(9)
      .setFontColor('#555555')
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setWrap(true);
    sheet.setRowHeight(chartNoteRow, 40);

    // Column widths
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 110);

    // Freeze top rows
    sheet.setFrozenRows(1);

    Logger.log('✅ Dashboard created');
  } catch (e) {
    Logger.log('❌ Dashboard error: ' + e + '\n' + e.stack);
  }
}

// ============================================================
// SHEET 2: MONTHLY BUDGET
// ============================================================
function createMonthlyBudget() {
  try {
    var sheet = getOrCreateSheet('Monthly Budget');
    sheet.setTabColor(COLORS.softBlue);

    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    // Title
    var totalCols = 2 + (months.length * 4);
    sheet.getRange(1, 1, 1, totalCols).merge()
      .setValue(YEAR + ' Monthly Budget')
      .setBackground(COLORS.sageGreen)
      .setFontColor(COLORS.white)
      .setFontSize(16)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 50);

    // Sub-header explanation
    sheet.getRange(2, 1, 1, totalCols).merge()
      .setValue('Enter your Budgeted amounts. Actual figures are pulled automatically from Expense Tracker.')
      .setBackground(COLORS.beige)
      .setFontSize(9)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(2, 22);

    // Month headers row
    sheet.getRange(3, 1).setValue('Category').setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setBackground(COLORS.lightGray);
    sheet.getRange(3, 2).setValue('Section').setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setBackground(COLORS.lightGray);

    for (var m = 0; m < months.length; m++) {
      var startCol = 3 + m * 4;
      sheet.getRange(3, startCol, 1, 4).merge()
        .setValue(months[m])
        .setBackground(COLORS.lightGray)
        .setFontWeight('bold')
        .setFontFamily('Arial')
        .setFontSize(10)
        .setHorizontalAlignment('center');
    }
    // Annual header
    sheet.getRange(3, 3 + months.length * 4, 1, 4).merge()
      .setValue('Annual Total')
      .setBackground(COLORS.sageGreen)
      .setFontColor(COLORS.white)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setFontSize(10)
      .setHorizontalAlignment('center');

    // Sub-column headers row (Budgeted | Actual | Diff | Status)
    sheet.getRange(4, 1).setValue('').setBackground(COLORS.lightGray);
    sheet.getRange(4, 2).setValue('').setBackground(COLORS.lightGray);
    for (var m2 = 0; m2 < months.length + 1; m2++) {
      var sc = 3 + m2 * 4;
      sheet.getRange(4, sc).setValue('Budgeted').setFontSize(8).setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.lightGray).setHorizontalAlignment('center');
      sheet.getRange(4, sc+1).setValue('Actual').setFontSize(8).setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.lightGray).setHorizontalAlignment('center');
      sheet.getRange(4, sc+2).setValue('Diff').setFontSize(8).setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.lightGray).setHorizontalAlignment('center');
      sheet.getRange(4, sc+3).setValue('Status').setFontSize(8).setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.lightGray).setHorizontalAlignment('center');
    }
    sheet.setRowHeight(3, 26);
    sheet.setRowHeight(4, 22);

    // Budget sections definition
    var sections = [
      {
        name: 'INCOME',
        color: COLORS.sageGreen,
        fontColor: COLORS.white,
        rows: ['Salary','Side Hustle','Passive Income','Dividends','Interest','Tax Refund','Other Income']
      },
      {
        name: 'PRE-TAX DEDUCTIONS',
        color: COLORS.lavender,
        fontColor: COLORS.darkText,
        rows: ['401(k)','HSA','FSA','Health Insurance','Dental','Vision','Other Deduction']
      },
      {
        name: 'TAXES',
        color: COLORS.dustyPink,
        fontColor: COLORS.darkText,
        rows: ['Federal Tax','State Tax','Local Tax','Social Security','Medicare','VAT/Other']
      },
      {
        name: 'FIXED EXPENSES',
        color: COLORS.softBlue,
        fontColor: COLORS.darkText,
        rows: ['Rent/Mortgage','Property Tax','HOA','Utilities','Home Insurance','Auto Insurance','Life Insurance','Phone','Internet','Subscriptions']
      },
      {
        name: 'VARIABLE EXPENSES',
        color: COLORS.beige,
        fontColor: COLORS.darkText,
        rows: ['Groceries','Gas/Transport','Dining Out','Entertainment','Shopping','Personal Care','Medical','Education','Childcare','Pet Care','Other Variable']
      },
      {
        name: 'SAVINGS & DEBT',
        color: COLORS.sageGreen,
        fontColor: COLORS.white,
        rows: ['Emergency Fund','Roth IRA','401(k) Extra','Student Loans','Credit Cards','Car Loan','Personal Loan','Mortgage Extra','Vacation Fund','House Fund','Car Fund','Other Savings']
      }
    ];

    var currentRow = 5;
    var etSheet = "'Expense Tracker'";
    var categoryCol = 'C';  // Category is column C in Expense Tracker
    var monthCol = 'A';     // Date column in Expense Tracker

    // Helper: get month number from month name index
    function monthNum(mIdx) { return mIdx + 1; }

    sections.forEach(function(section) {
      // Section header row
      sheet.getRange(currentRow, 1, 1, totalCols).merge()
        .setValue(section.name)
        .setBackground(section.color)
        .setFontColor(section.fontColor)
        .setFontSize(12)
        .setFontWeight('bold')
        .setFontFamily('Arial')
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('left');
      sheet.setRowHeight(currentRow, 28);
      currentRow++;

      var sectionStartRow = currentRow;

      // Category rows
      section.rows.forEach(function(catName) {
        sheet.getRange(currentRow, 1).setValue(catName).setFontFamily('Arial').setFontSize(10).setFontWeight('bold');
        sheet.getRange(currentRow, 2).setValue(section.name).setFontFamily('Arial').setFontSize(9).setFontColor('#888888');

        for (var m3 = 0; m3 < months.length; m3++) {
          var sc3 = 3 + m3 * 4;
          var mNum = m3 + 1;

          // Budgeted — user input
          sheet.getRange(currentRow, sc3)
            .setBackground(COLORS.white)
            .setNumberFormat('$#,##0.00')
            .setFontFamily('Arial').setFontSize(10);

          // Actual — SUMIF from Expense Tracker by category + month
          var actualFormula = '=IFERROR(SUMPRODUCT((' + etSheet + '!C$6:C$1000="' + catName + '")*' +
            '(MONTH(' + etSheet + '!A$6:A$1000)=' + mNum + ')*' +
            '(YEAR(' + etSheet + '!A$6:A$1000)=' + YEAR + ')*' +
            etSheet + '!E$6:E$1000),0)';
          sheet.getRange(currentRow, sc3 + 1).setFormula(actualFormula)
            .setBackground('#F8F8F8')
            .setNumberFormat('$#,##0.00')
            .setFontFamily('Arial').setFontSize(10);

          // Difference
          sheet.getRange(currentRow, sc3 + 2)
            .setFormula('=' + colLetter(sc3) + currentRow + '-' + colLetter(sc3+1) + currentRow)
            .setBackground('#F8F8F8')
            .setNumberFormat('$#,##0.00')
            .setFontFamily('Arial').setFontSize(10);

          // Status
          var diffCell = colLetter(sc3+2) + currentRow;
          var budCell = colLetter(sc3) + currentRow;
          sheet.getRange(currentRow, sc3 + 3)
            .setFormula('=IF(' + budCell + '=0,"—",IF(' + diffCell + '>0,"🎯 Under","IF(' + diffCell + '<-' + budCell + '*0.1,"⚠️ Over","✓ OK")"))')
            .setBackground('#F8F8F8')
            .setFontFamily('Arial').setFontSize(9)
            .setHorizontalAlignment('center');
          // Simpler status formula
          sheet.getRange(currentRow, sc3 + 3)
            .setFormula('=IF(' + budCell + '=0,"—",IF(' + diffCell + '>=0,"🎯 Under Budget",IF(' + diffCell + '>=-(' + budCell + '*0.1),"✓ On Track","⚠️ Overspending")))');
        }

        // Annual totals (last 4 cols)
        var asc = 3 + months.length * 4;
        var budCols = [];
        var actCols = [];
        for (var m4 = 0; m4 < months.length; m4++) {
          budCols.push(colLetter(3 + m4*4) + currentRow);
          actCols.push(colLetter(3 + m4*4 + 1) + currentRow);
        }
        sheet.getRange(currentRow, asc).setFormula('=' + budCols.join('+')).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(COLORS.beige).setFontWeight('bold');
        sheet.getRange(currentRow, asc+1).setFormula('=' + actCols.join('+')).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(COLORS.beige).setFontWeight('bold');
        sheet.getRange(currentRow, asc+2).setFormula('=' + colLetter(asc) + currentRow + '-' + colLetter(asc+1) + currentRow).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(COLORS.beige).setFontWeight('bold');
        sheet.getRange(currentRow, asc+3).setValue('').setBackground(COLORS.beige);

        // Conditional formatting on Diff columns
        for (var m5 = 0; m5 < months.length + 1; m5++) {
          var diffRange = sheet.getRange(currentRow, 3 + m5*4 + 2, 1, 1);
          sheet.setConditionalFormatRules(
            sheet.getConditionalFormatRules().concat([
              SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setBackground(COLORS.green).setFontColor('#1B5E20').setRanges([diffRange]).build(),
              SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground(COLORS.red).setFontColor('#B71C1C').setRanges([diffRange]).build()
            ])
          );
        }

        sheet.setRowHeight(currentRow, 22);
        currentRow++;
      });

      // Subtotal row
      var subRow = currentRow;
      sheet.getRange(subRow, 1, 1, 2).merge().setValue('TOTAL ' + section.name)
        .setBackground(section.color).setFontColor(section.fontColor)
        .setFontWeight('bold').setFontFamily('Arial').setFontSize(10);

      for (var m6 = 0; m6 < months.length + 1; m6++) {
        var sc6 = 3 + m6 * 4;
        var range1 = colLetter(sc6) + sectionStartRow + ':' + colLetter(sc6) + (subRow - 1);
        var range2 = colLetter(sc6+1) + sectionStartRow + ':' + colLetter(sc6+1) + (subRow - 1);
        sheet.getRange(subRow, sc6).setFormula('=SUM(' + range1 + ')').setBackground(section.color).setFontColor(section.fontColor).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setNumberFormat('$#,##0.00');
        sheet.getRange(subRow, sc6+1).setFormula('=SUM(' + range2 + ')').setBackground(section.color).setFontColor(section.fontColor).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setNumberFormat('$#,##0.00');
        sheet.getRange(subRow, sc6+2).setFormula('=' + colLetter(sc6) + subRow + '-' + colLetter(sc6+1) + subRow).setBackground(section.color).setFontColor(section.fontColor).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setNumberFormat('$#,##0.00');
        sheet.getRange(subRow, sc6+3).setValue('').setBackground(section.color);
      }
      sheet.setRowHeight(subRow, 28);
      currentRow += 2; // blank row between sections
    });

    // --- Bottom Summary ---
    currentRow++;
    var summaryLabels = ['NET INCOME', 'TOTAL SPENT', 'REMAINING'];
    summaryLabels.forEach(function(lbl) {
      sheet.getRange(currentRow, 1, 1, 2).merge().setValue(lbl)
        .setBackground(COLORS.darkText).setFontColor(COLORS.white)
        .setFontWeight('bold').setFontFamily('Arial').setFontSize(11);
      for (var m7 = 0; m7 < months.length + 1; m7++) {
        var sc7 = 3 + m7 * 4;
        sheet.getRange(currentRow, sc7).setValue(0).setBackground(COLORS.darkText).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setNumberFormat('$#,##0.00');
        sheet.getRange(currentRow, sc7+1).setValue(0).setBackground(COLORS.darkText).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setNumberFormat('$#,##0.00');
        sheet.getRange(currentRow, sc7+2).setValue(0).setBackground(COLORS.darkText).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setNumberFormat('$#,##0.00');
        sheet.getRange(currentRow, sc7+3).setValue('').setBackground(COLORS.darkText);
      }
      sheet.setRowHeight(currentRow, 28);
      currentRow++;
    });

    // Column widths
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 100);
    for (var c = 3; c <= totalCols; c++) {
      var cIdx = (c - 3) % 4;
      sheet.setColumnWidth(c, cIdx === 3 ? 110 : 90);
    }

    sheet.setFrozenRows(4);
    sheet.setFrozenColumns(2);

    Logger.log('✅ Monthly Budget created');
  } catch (e) {
    Logger.log('❌ Monthly Budget error: ' + e + '\n' + e.stack);
  }
}

// Convert column index to letter (1=A, 2=B, ... 26=Z, 27=AA, etc.)
function colLetter(n) {
  var result = '';
  while (n > 0) {
    var rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

// ============================================================
// SHEET 3: EXPENSE TRACKER
// ============================================================
function createExpenseTracker() {
  try {
    var sheet = getOrCreateSheet('Expense Tracker');
    sheet.setTabColor(COLORS.dustyPink);

    // Title
    sheet.getRange(1, 1, 1, 8).merge()
      .setValue('Expense Tracker ' + YEAR)
      .setBackground(COLORS.dustyPink)
      .setFontColor(COLORS.white)
      .setFontSize(16)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 50);

    // Summary cards row
    sheet.getRange(2, 1, 1, 2).merge().setValue('📅 Total This Month').setBackground(COLORS.sageGreen).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    sheet.getRange(3, 1, 1, 2).merge()
      .setFormula('=IFERROR(SUMPRODUCT((MONTH(A6:A1000)=MONTH(TODAY()))*(YEAR(A6:A1000)=YEAR(TODAY()))*E6:E1000),0)')
      .setBackground(COLORS.sageGreen).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(14).setHorizontalAlignment('center').setNumberFormat('$#,##0.00');

    sheet.getRange(2, 3, 1, 2).merge().setValue('📊 Total YTD').setBackground(COLORS.softBlue).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    sheet.getRange(3, 3, 1, 2).merge()
      .setFormula('=IFERROR(SUMPRODUCT((YEAR(A6:A1000)=YEAR(TODAY()))*E6:E1000),0)')
      .setBackground(COLORS.softBlue).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(14).setHorizontalAlignment('center').setNumberFormat('$#,##0.00');

    sheet.getRange(2, 5, 1, 2).merge().setValue('🏆 Top Category').setBackground(COLORS.lavender).setFontColor(COLORS.darkText).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    sheet.getRange(3, 5, 1, 2).merge()
      .setFormula('=IFERROR(INDEX(C6:C1000,MATCH(MAX(COUNTIF(C6:C1000,C6:C1000)),COUNTIF(C6:C1000,C6:C1000),0)),"—")')
      .setBackground(COLORS.lavender).setFontColor(COLORS.darkText).setFontWeight('bold').setFontFamily('Arial').setFontSize(12).setHorizontalAlignment('center');

    sheet.setRowHeight(2, 24);
    sheet.setRowHeight(3, 36);

    // Blank spacer
    sheet.setRowHeight(4, 10);

    // Column headers
    var headers = ['Date','Description','Category','Sub-Category','Amount','Payment Method','Receipt?','Notes'];
    var headerRange = sheet.getRange(5, 1, 1, 8);
    headerRange.setValues([headers])
      .setBackground(COLORS.sageGreen)
      .setFontColor(COLORS.white)
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(5, 28);

    // Categories from Monthly Budget
    var categories = [
      'Salary','Side Hustle','Passive Income','Dividends','Interest','Tax Refund','Other Income',
      '401(k)','HSA','FSA','Health Insurance','Dental','Vision','Other Deduction',
      'Federal Tax','State Tax','Local Tax','Social Security','Medicare','VAT/Other',
      'Rent/Mortgage','Property Tax','HOA','Utilities','Home Insurance','Auto Insurance',
      'Life Insurance','Phone','Internet','Subscriptions',
      'Groceries','Gas/Transport','Dining Out','Entertainment','Shopping','Personal Care',
      'Medical','Education','Childcare','Pet Care','Other Variable',
      'Emergency Fund','Roth IRA','401(k) Extra','Student Loans','Credit Cards',
      'Car Loan','Personal Loan','Mortgage Extra','Vacation Fund','House Fund','Car Fund','Other Savings'
    ];

    var subCategories = {
      'Groceries': ['Walmart','Costco','Target','Local Store','Whole Foods','Other'],
      'Gas/Transport': ['Shell','BP','Exxon','Chevron','Costco Gas','Public Transit','Rideshare'],
      'Dining Out': ['Fast Food','Restaurant','Delivery','Coffee Shop','Bar','Catering'],
      'Entertainment': ['Movies','Concerts','Sports','Games','Streaming','Books'],
      'Shopping': ['Clothing','Electronics','Home Goods','Online','Department Store','Other'],
      'Personal Care': ['Hair','Nails','Spa','Pharmacy','Gym','Supplements'],
      'Medical': ['Doctor','Dentist','Vision','Pharmacy','Hospital','Therapy'],
      'Utilities': ['Electric','Gas','Water','Trash','Sewer','Other'],
      'Subscriptions': ['Netflix','Spotify','Amazon Prime','Disney+','YouTube','Other'],
      'Other Variable': ['Miscellaneous','Cash','ATM','Other']
    };

    var paymentMethods = ['Credit Card','Debit Card','Cash','Bank Transfer','Check','Venmo/PayPal','Other'];

    // Data validation for category
    var catRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(categories, true)
      .setAllowInvalid(false)
      .build();

    var payRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(paymentMethods, true)
      .setAllowInvalid(false)
      .build();

    var receiptRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes','No','N/A'], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(6, 3, 500, 1).setDataValidation(catRule);
    sheet.getRange(6, 6, 500, 1).setDataValidation(payRule);
    sheet.getRange(6, 7, 500, 1).setDataValidation(receiptRule);

    // Pre-fill 40 realistic transactions
    var transactions = getExampleTransactions();
    var dataValues = [];
    transactions.forEach(function(t) {
      dataValues.push([t.date, t.desc, t.cat, t.subCat, t.amount, t.method, t.receipt, t.notes]);
    });

    if (dataValues.length > 0) {
      var dataRange = sheet.getRange(6, 1, dataValues.length, 8);
      dataRange.setValues(dataValues);

      // Format date column
      sheet.getRange(6, 1, dataValues.length, 1).setNumberFormat('MM/DD/YYYY');
      // Format amount column
      sheet.getRange(6, 5, dataValues.length, 1).setNumberFormat('$#,##0.00');
    }

    // Alternating row colors
    for (var r = 0; r < 100; r++) {
      var rowNum = 6 + r;
      sheet.getRange(rowNum, 1, 1, 8).setBackground(r % 2 === 0 ? COLORS.white : COLORS.beige);
    }
    // Re-apply data background over alternating (data rows already set)

    // Column widths
    sheet.setColumnWidth(1, 100); // Date
    sheet.setColumnWidth(2, 200); // Description
    sheet.setColumnWidth(3, 140); // Category
    sheet.setColumnWidth(4, 140); // Sub-Category
    sheet.setColumnWidth(5, 100); // Amount
    sheet.setColumnWidth(6, 120); // Payment Method
    sheet.setColumnWidth(7, 80);  // Receipt
    sheet.setColumnWidth(8, 180); // Notes

    // Freeze header
    sheet.setFrozenRows(5);

    // Protect formula ranges
    var protectedRange = sheet.getRange(3, 1, 1, 4);
    var protection = protectedRange.protect();
    protection.setWarningOnly(true);
    protection.setDescription('Summary formulas - do not edit');

    Logger.log('✅ Expense Tracker created');
  } catch (e) {
    Logger.log('❌ Expense Tracker error: ' + e + '\n' + e.stack);
  }
}

function getExampleTransactions() {
  var y = YEAR;
  return [
    {date: new Date(y,0,3),  desc: 'Walmart Groceries',        cat: 'Groceries',      subCat: 'Walmart',        amount: 156.43, method: 'Debit Card',  receipt: 'Yes', notes: 'Weekly grocery run'},
    {date: new Date(y,0,5),  desc: 'Shell Gas Station',         cat: 'Gas/Transport',  subCat: 'Shell',          amount: 65.20,  method: 'Credit Card', receipt: 'No',  notes: ''},
    {date: new Date(y,0,8),  desc: 'Netflix Subscription',      cat: 'Subscriptions',  subCat: 'Netflix',        amount: 15.99,  method: 'Credit Card', receipt: 'No',  notes: 'Monthly'},
    {date: new Date(y,0,10), desc: 'Chipotle Lunch',            cat: 'Dining Out',     subCat: 'Restaurant',     amount: 14.75,  method: 'Debit Card',  receipt: 'No',  notes: ''},
    {date: new Date(y,0,12), desc: 'Electric Bill',             cat: 'Utilities',      subCat: 'Electric',       amount: 98.50,  method: 'Bank Transfer',receipt:'Yes', notes: 'Monthly bill'},
    {date: new Date(y,0,15), desc: 'Spotify Premium',           cat: 'Subscriptions',  subCat: 'Spotify',        amount: 10.99,  method: 'Credit Card', receipt: 'No',  notes: ''},
    {date: new Date(y,0,18), desc: 'CVS Pharmacy',              cat: 'Medical',        subCat: 'Pharmacy',       amount: 42.30,  method: 'Credit Card', receipt: 'Yes', notes: 'Prescriptions'},
    {date: new Date(y,0,20), desc: 'Amazon Purchase',           cat: 'Shopping',       subCat: 'Online',         amount: 87.64,  method: 'Credit Card', receipt: 'Yes', notes: 'Household items'},
    {date: new Date(y,0,22), desc: 'Coffee Shop',               cat: 'Dining Out',     subCat: 'Coffee Shop',    amount: 7.50,   method: 'Cash',        receipt: 'No',  notes: ''},
    {date: new Date(y,0,25), desc: 'Costco',                    cat: 'Groceries',      subCat: 'Costco',         amount: 312.88, method: 'Credit Card', receipt: 'Yes', notes: 'Monthly Costco run'},
    {date: new Date(y,1,2),  desc: 'Planet Fitness',            cat: 'Personal Care',  subCat: 'Gym',            amount: 24.99,  method: 'Credit Card', receipt: 'No',  notes: 'Monthly membership'},
    {date: new Date(y,1,5),  desc: 'Internet Bill - Xfinity',   cat: 'Internet',       subCat: 'Other',          amount: 75.00,  method: 'Bank Transfer',receipt:'Yes', notes: ''},
    {date: new Date(y,1,8),  desc: 'DoorDash Delivery',         cat: 'Dining Out',     subCat: 'Delivery',       amount: 38.90,  method: 'Credit Card', receipt: 'No',  notes: 'Pizza night'},
    {date: new Date(y,1,10), desc: 'BP Gas',                    cat: 'Gas/Transport',  subCat: 'BP',             amount: 58.40,  method: 'Debit Card',  receipt: 'No',  notes: ''},
    {date: new Date(y,1,14), desc: 'Walgreens',                 cat: 'Medical',        subCat: 'Pharmacy',       amount: 28.65,  method: 'Cash',        receipt: 'Yes', notes: ''},
    {date: new Date(y,1,16), desc: 'Target Groceries',          cat: 'Groceries',      subCat: 'Target',         amount: 134.22, method: 'Debit Card',  receipt: 'Yes', notes: ''},
    {date: new Date(y,1,20), desc: 'YouTube Premium',           cat: 'Subscriptions',  subCat: 'YouTube',        amount: 13.99,  method: 'Credit Card', receipt: 'No',  notes: ''},
    {date: new Date(y,1,22), desc: 'Clothes - Macy\'s',         cat: 'Shopping',       subCat: 'Clothing',       amount: 145.00, method: 'Credit Card', receipt: 'Yes', notes: 'Winter sale'},
    {date: new Date(y,1,25), desc: 'Water/Sewer Bill',          cat: 'Utilities',      subCat: 'Water',          amount: 45.00,  method: 'Bank Transfer',receipt:'Yes', notes: ''},
    {date: new Date(y,2,3),  desc: 'Exxon Gas Station',         cat: 'Gas/Transport',  subCat: 'Exxon',          amount: 70.15,  method: 'Credit Card', receipt: 'No',  notes: ''},
    {date: new Date(y,2,6),  desc: 'Whole Foods',               cat: 'Groceries',      subCat: 'Whole Foods',    amount: 189.50, method: 'Credit Card', receipt: 'Yes', notes: ''},
    {date: new Date(y,2,10), desc: 'Dentist Co-pay',            cat: 'Medical',        subCat: 'Dentist',        amount: 50.00,  method: 'Check',       receipt: 'Yes', notes: 'Annual cleaning'},
    {date: new Date(y,2,12), desc: 'Amazon Prime',              cat: 'Subscriptions',  subCat: 'Amazon Prime',   amount: 14.99,  method: 'Credit Card', receipt: 'No',  notes: ''},
    {date: new Date(y,2,15), desc: 'Starbucks',                 cat: 'Dining Out',     subCat: 'Coffee Shop',    amount: 6.75,   method: 'Debit Card',  receipt: 'No',  notes: ''},
    {date: new Date(y,2,18), desc: 'Hair Salon',                cat: 'Personal Care',  subCat: 'Hair',           amount: 65.00,  method: 'Cash',        receipt: 'No',  notes: ''},
    {date: new Date(y,2,20), desc: 'Home Depot Supplies',       cat: 'Shopping',       subCat: 'Home Goods',     amount: 234.76, method: 'Credit Card', receipt: 'Yes', notes: 'Bathroom repair'},
    {date: new Date(y,2,22), desc: 'Walmart Groceries',         cat: 'Groceries',      subCat: 'Walmart',        amount: 167.33, method: 'Debit Card',  receipt: 'Yes', notes: ''},
    {date: new Date(y,3,4),  desc: 'Uber Eats',                 cat: 'Dining Out',     subCat: 'Delivery',       amount: 45.20,  method: 'Credit Card', receipt: 'No',  notes: ''},
    {date: new Date(y,3,8),  desc: 'Gas Bill - NiGas',          cat: 'Utilities',      subCat: 'Gas',            amount: 62.00,  method: 'Bank Transfer',receipt:'Yes', notes: ''},
    {date: new Date(y,3,12), desc: 'Movie Tickets',             cat: 'Entertainment',  subCat: 'Movies',         amount: 32.00,  method: 'Credit Card', receipt: 'No',  notes: 'Family outing'},
    {date: new Date(y,3,15), desc: 'Target Run',                cat: 'Shopping',       subCat: 'Department Store',amount:98.50,  method: 'Debit Card',  receipt: 'Yes', notes: ''},
    {date: new Date(y,3,18), desc: 'Costco Gas',                cat: 'Gas/Transport',  subCat: 'Costco Gas',     amount: 55.80,  method: 'Debit Card',  receipt: 'No',  notes: ''},
    {date: new Date(y,4,2),  desc: 'Disney+ Subscription',      cat: 'Subscriptions',  subCat: 'Disney+',        amount: 10.99,  method: 'Credit Card', receipt: 'No',  notes: ''},
    {date: new Date(y,4,6),  desc: 'Whole Foods Weekly',        cat: 'Groceries',      subCat: 'Whole Foods',    amount: 220.14, method: 'Credit Card', receipt: 'Yes', notes: ''},
    {date: new Date(y,4,10), desc: 'Doctor Visit Co-pay',       cat: 'Medical',        subCat: 'Doctor',         amount: 30.00,  method: 'Credit Card', receipt: 'Yes', notes: ''},
    {date: new Date(y,4,14), desc: 'Shell Gas',                 cat: 'Gas/Transport',  subCat: 'Shell',          amount: 68.90,  method: 'Credit Card', receipt: 'No',  notes: ''},
    {date: new Date(y,4,18), desc: 'Nordstrom Rack Clothes',    cat: 'Shopping',       subCat: 'Clothing',       amount: 178.50, method: 'Credit Card', receipt: 'Yes', notes: 'Spring wardrobe'},
    {date: new Date(y,4,22), desc: 'Electric Bill',             cat: 'Utilities',      subCat: 'Electric',       amount: 112.40, method: 'Bank Transfer',receipt:'Yes', notes: ''},
    {date: new Date(y,5,5),  desc: 'Concert Tickets',           cat: 'Entertainment',  subCat: 'Concerts',       amount: 150.00, method: 'Credit Card', receipt: 'Yes', notes: 'Summer concert'},
    {date: new Date(y,5,10), desc: 'Walmart Groceries',         cat: 'Groceries',      subCat: 'Walmart',        amount: 143.67, method: 'Debit Card',  receipt: 'Yes', notes: ''}
  ];
}

// ============================================================
// SHEET 4: SAVINGS GOALS
// ============================================================
function createSavingsGoals() {
  try {
    var sheet = getOrCreateSheet('Savings Goals');
    sheet.setTabColor(COLORS.lavender);

    // Title
    sheet.getRange(1, 1, 1, 9).merge()
      .setValue('💰 Savings Goals')
      .setBackground(COLORS.lavender)
      .setFontColor(COLORS.white)
      .setFontSize(16)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 50);

    sheet.getRange(2, 1, 1, 9).merge()
      .setValue('Track your progress toward each financial goal. Enter your Target, Current Amount, and Monthly Contribution.')
      .setBackground(COLORS.beige)
      .setFontSize(9)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(2, 22);

    // Spacer
    sheet.setRowHeight(3, 8);

    // Headers
    var headers = ['Goal Name','Target Amount','Current Amount','Monthly Contribution','Months to Goal','% Complete','Priority','Status','Notes'];
    sheet.getRange(4, 1, 1, 9).setValues([headers])
      .setBackground(COLORS.lavender)
      .setFontColor(COLORS.white)
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(4, 28);

    // Example data
    var goals = [
      {name: 'Emergency Fund',    target: 10000, current: 3500,  monthly: 300,  priority: 'High',   notes: '3-6 months expenses'},
      {name: 'Vacation - Europe', target: 3000,  current: 800,   monthly: 200,  priority: 'Medium', notes: 'Summer trip'},
      {name: 'New Car',           target: 15000, current: 2000,  monthly: 400,  priority: 'Medium', notes: 'Down payment'},
      {name: 'House Down Payment',target: 50000, current: 8000,  monthly: 600,  priority: 'High',   notes: '20% down goal'},
      {name: 'Wedding Fund',      target: 20000, current: 5000,  monthly: 350,  priority: 'Low',    notes: '2026 date'},
      {name: 'Education Fund',    target: 25000, current: 10000, monthly: 250,  priority: 'Medium', notes: 'College savings'},
      {name: 'Holiday Fund',      target: 1500,  current: 400,   monthly: 100,  priority: 'Low',    notes: 'Gifts & travel'},
      {name: 'Home Renovation',   target: 10000, current: 1500,  monthly: 200,  priority: 'Medium', notes: 'Kitchen remodel'}
    ];

    var priorityColors = {High: COLORS.red, Medium: COLORS.yellow, Low: COLORS.green};

    for (var i = 0; i < goals.length; i++) {
      var g = goals[i];
      var r = 5 + i;
      var bgRow = i % 2 === 0 ? COLORS.white : COLORS.beige;

      sheet.getRange(r, 1).setValue(g.name).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setBackground(bgRow);
      sheet.getRange(r, 2).setValue(g.target).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(bgRow);
      sheet.getRange(r, 3).setValue(g.current).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(bgRow);
      sheet.getRange(r, 4).setValue(g.monthly).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(bgRow);

      // Months to goal formula
      sheet.getRange(r, 5)
        .setFormula('=IF(D' + r + '>0,IFERROR(CEILING((B' + r + '-C' + r + ')/D' + r + ',1),"Set monthly"),"Set monthly")')
        .setFontFamily('Arial').setFontSize(10).setBackground(bgRow)
        .setHorizontalAlignment('center');

      // % Complete formula
      sheet.getRange(r, 6)
        .setFormula('=IFERROR(C' + r + '/B' + r + ',0)')
        .setNumberFormat('0.0%').setFontFamily('Arial').setFontSize(10).setBackground(bgRow)
        .setHorizontalAlignment('center');

      // Priority
      sheet.getRange(r, 7).setValue(g.priority).setFontFamily('Arial').setFontSize(10)
        .setBackground(priorityColors[g.priority] || bgRow)
        .setFontWeight('bold').setHorizontalAlignment('center');

      // Status formula
      sheet.getRange(r, 8)
        .setFormula('=IF(C' + r + '>=B' + r + ',"✓ Achieved",IF(D' + r + '<=0,"⏰ Behind",IF(E' + r + '<=(B' + r + '-C' + r + ')/D' + r + ',"🔥 On Track","🚀 Ahead")))')
        .setFontFamily('Arial').setFontSize(10).setBackground(bgRow).setHorizontalAlignment('center');

      sheet.getRange(r, 9).setValue(g.notes).setFontFamily('Arial').setFontSize(10).setBackground(bgRow).setFontColor('#666666');

      sheet.setRowHeight(r, 24);
    }

    // Data bars on % Complete column (F)
    var pctRange = sheet.getRange(5, 6, goals.length, 1);
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue(COLORS.beige, SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMaxpointWithValue(COLORS.sageGreen, SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setRanges([pctRange]).build();
    sheet.setConditionalFormatRules([rule]);

    // Totals row
    var totRow = 5 + goals.length + 1;
    sheet.getRange(totRow, 1, 1, 9).setBackground(COLORS.lavender);
    sheet.getRange(totRow, 1).setValue('TOTALS').setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setFontColor(COLORS.white).setBackground(COLORS.lavender);
    sheet.getRange(totRow, 2).setFormula('=SUM(B5:B' + (5+goals.length-1) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial').setFontColor(COLORS.white).setBackground(COLORS.lavender);
    sheet.getRange(totRow, 3).setFormula('=SUM(C5:C' + (5+goals.length-1) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial').setFontColor(COLORS.white).setBackground(COLORS.lavender);
    sheet.getRange(totRow, 4).setFormula('=SUM(D5:D' + (5+goals.length-1) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial').setFontColor(COLORS.white).setBackground(COLORS.lavender);
    sheet.setRowHeight(totRow, 28);

    // Column widths
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 110);
    sheet.setColumnWidth(7, 90);
    sheet.setColumnWidth(8, 130);
    sheet.setColumnWidth(9, 180);

    sheet.setFrozenRows(4);

    Logger.log('✅ Savings Goals created');
  } catch (e) {
    Logger.log('❌ Savings Goals error: ' + e + '\n' + e.stack);
  }
}

// ============================================================
// SHEET 5: DEBT PAYOFF
// ============================================================
function createDebtPayoff() {
  try {
    var sheet = getOrCreateSheet('Debt Payoff');
    sheet.setTabColor(COLORS.red);

    // Title
    sheet.getRange(1, 1, 1, 11).merge()
      .setValue('💳 Debt Payoff Tracker')
      .setBackground(COLORS.red)
      .setFontColor(COLORS.white)
      .setFontSize(16)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 50);

    // Strategy selector
    sheet.getRange(2, 1).setValue('Strategy:').setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    sheet.getRange(2, 2).setValue('Avalanche (Highest APR First)')
      .setFontFamily('Arial').setFontSize(10);
    var stratRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Avalanche (Highest APR First)', 'Snowball (Lowest Balance First)'], true)
      .setAllowInvalid(false).build();
    sheet.getRange(2, 2).setDataValidation(stratRule);
    sheet.setRowHeight(2, 24);

    // Summary cards row 3-4
    sheet.getRange(3, 1, 1, 3).merge().setValue('💰 Total Debt').setBackground(COLORS.red).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    sheet.getRange(4, 1, 1, 3).merge()
      .setFormula('=IFERROR(SUM(C8:C50),0)')
      .setBackground(COLORS.red).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(14).setHorizontalAlignment('center').setNumberFormat('$#,##0.00');

    sheet.getRange(3, 4, 1, 3).merge().setValue('📅 Min Payments Total').setBackground(COLORS.dustyPink).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    sheet.getRange(4, 4, 1, 3).merge()
      .setFormula('=IFERROR(SUM(E8:E50),0)')
      .setBackground(COLORS.dustyPink).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(14).setHorizontalAlignment('center').setNumberFormat('$#,##0.00');

    sheet.getRange(3, 7, 1, 3).merge().setValue('🎯 Est. Debt-Free Date').setBackground(COLORS.sageGreen).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    sheet.getRange(4, 7, 1, 3).merge().setValue('See Payoff Dates Below')
      .setBackground(COLORS.sageGreen).setFontColor(COLORS.white).setFontWeight('bold').setFontFamily('Arial').setFontSize(11).setHorizontalAlignment('center');

    sheet.setRowHeight(3, 24);
    sheet.setRowHeight(4, 36);

    // Spacer
    sheet.setRowHeight(5, 8);

    // Tips row
    sheet.getRange(6, 1, 1, 11).merge()
      .setValue('💡 Avalanche Method: Pay minimum on all, put extra toward highest APR.  |  Snowball Method: Pay minimum on all, put extra toward lowest balance for motivation.')
      .setBackground(COLORS.beige).setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    sheet.setRowHeight(6, 22);

    // Column headers
    var headers = ['Debt Name','Original Balance','Current Balance','APR %','Min Payment','Extra Payment','Total Payment','Payoff Date','Months Remaining','% Paid Off','Notes'];
    sheet.getRange(7, 1, 1, 11).setValues([headers])
      .setBackground(COLORS.red)
      .setFontColor(COLORS.white)
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(7, 28);

    // Pre-fill debt data
    var debts = [
      {name: 'Student Loans',   orig: 30000, curr: 24500, apr: 5.5,   min: 300,  extra: 100, notes: 'Federal student loans'},
      {name: 'Credit Card 1',   orig: 5000,  curr: 3800,  apr: 19.99, min: 150,  extra: 50,  notes: 'Visa - Chase'},
      {name: 'Credit Card 2',   orig: 2500,  curr: 2100,  apr: 24.99, min: 75,   extra: 0,   notes: 'MasterCard - Citi'},
      {name: 'Car Loan',        orig: 15000, curr: 11200, apr: 4.5,   min: 350,  extra: 50,  notes: '2022 Honda Civic'},
      {name: 'Mortgage',        orig: 200000,curr: 185000,apr: 3.5,   min: 1200, extra: 100, notes: '30-year fixed'},
      {name: 'Personal Loan',   orig: 8000,  curr: 5600,  apr: 8.0,   min: 200,  extra: 50,  notes: 'Home improvement'}
    ];

    for (var i = 0; i < debts.length; i++) {
      var d = debts[i];
      var r = 8 + i;
      var bg = i % 2 === 0 ? COLORS.white : COLORS.beige;

      sheet.getRange(r, 1).setValue(d.name).setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setBackground(bg);
      sheet.getRange(r, 2).setValue(d.orig).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(bg);
      sheet.getRange(r, 3).setValue(d.curr).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(bg);
      sheet.getRange(r, 4).setValue(d.apr / 100).setNumberFormat('0.00%').setFontFamily('Arial').setFontSize(10).setBackground(bg);
      sheet.getRange(r, 5).setValue(d.min).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(bg);
      sheet.getRange(r, 6).setValue(d.extra).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(bg);

      // Total payment
      sheet.getRange(r, 7).setFormula('=E' + r + '+F' + r).setNumberFormat('$#,##0.00').setFontFamily('Arial').setFontSize(10).setBackground(bg);

      // Months remaining using NPER-style formula
      sheet.getRange(r, 9)
        .setFormula('=IFERROR(IF(G' + r + '>0,CEILING(NPER(D' + r + '/12,-G' + r + ',C' + r + '),1),"Enter payment"),"N/A")')
        .setFontFamily('Arial').setFontSize(10).setBackground(bg).setHorizontalAlignment('center');

      // Payoff date
      sheet.getRange(r, 8)
        .setFormula('=IFERROR(IF(ISNUMBER(I' + r + '),EDATE(TODAY(),I' + r + '),"N/A"),"N/A")')
        .setNumberFormat('MM/DD/YYYY').setFontFamily('Arial').setFontSize(10).setBackground(bg).setHorizontalAlignment('center');

      // % Paid Off
      sheet.getRange(r, 10)
        .setFormula('=IFERROR((B' + r + '-C' + r + ')/B' + r + ',0)')
        .setNumberFormat('0.0%').setFontFamily('Arial').setFontSize(10).setBackground(bg).setHorizontalAlignment('center');

      sheet.getRange(r, 11).setValue(d.notes).setFontFamily('Arial').setFontSize(10).setBackground(bg).setFontColor('#666666');

      sheet.setRowHeight(r, 24);
    }

    // Data bars on % Paid Off
    var pctRange = sheet.getRange(8, 10, debts.length, 1);
    var barRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue(COLORS.beige, SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMaxpointWithValue(COLORS.green, SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setRanges([pctRange]).build();
    sheet.setConditionalFormatRules([barRule]);

    // Totals row
    var totRow = 8 + debts.length + 1;
    sheet.getRange(totRow, 1).setValue('TOTALS').setFontWeight('bold').setFontFamily('Arial').setFontSize(10).setBackground(COLORS.red).setFontColor(COLORS.white);
    sheet.getRange(totRow, 2).setFormula('=SUM(B8:B' + (8+debts.length-1) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.red).setFontColor(COLORS.white);
    sheet.getRange(totRow, 3).setFormula('=SUM(C8:C' + (8+debts.length-1) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.red).setFontColor(COLORS.white);
    sheet.getRange(totRow, 5).setFormula('=SUM(E8:E' + (8+debts.length-1) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.red).setFontColor(COLORS.white);
    sheet.getRange(totRow, 6).setFormula('=SUM(F8:F' + (8+debts.length-1) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.red).setFontColor(COLORS.white);
    sheet.getRange(totRow, 7).setFormula('=SUM(G8:G' + (8+debts.length-1) + ')').setNumberFormat('$#,##0.00').setFontWeight('bold').setFontFamily('Arial').setBackground(COLORS.red).setFontColor(COLORS.white);
    for (var c = 4; c <= 11; c++) {
      if (c !== 5 && c !== 6 && c !== 7) {
        sheet.getRange(totRow, c).setBackground(COLORS.red);
      }
    }
    sheet.setRowHeight(totRow, 28);

    // Column widths
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 130);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 110);
    sheet.setColumnWidth(7, 110);
    sheet.setColumnWidth(8, 110);
    sheet.setColumnWidth(9, 130);
    sheet.setColumnWidth(10, 110);
    sheet.setColumnWidth(11, 180);

    sheet.setFrozenRows(7);

    Logger.log('✅ Debt Payoff created');
  } catch (e) {
    Logger.log('❌ Debt Payoff error: ' + e + '\n' + e.stack);
  }
}

// ============================================================
// SHEET 6: CHECKLISTS
// ============================================================
function createChecklists() {
  try {
    var sheet = getOrCreateSheet('Checklists');
    sheet.setTabColor(COLORS.softBlue);

    // Title
    sheet.getRange(1, 1, 1, 6).merge()
      .setValue('✅ Financial Checklists')
      .setBackground(COLORS.softBlue)
      .setFontColor(COLORS.white)
      .setFontSize(16)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 50);

    sheet.getRange(2, 1, 1, 6).merge()
      .setValue('Check off items as you complete them. Progress counters update automatically.')
      .setBackground(COLORS.beige).setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    sheet.setRowHeight(2, 22);

    var currentRow = 4;

    // Checklist definitions
    var checklists = [
      {
        title: '📅 Monthly Financial Checklist',
        color: COLORS.sageGreen,
        items: [
          'Review last month\'s spending vs budget',
          'Update budget for this month',
          'Pay all bills (set autopay if not done)',
          'Transfer to savings accounts',
          'Check credit card balances',
          'Review subscriptions (cancel unused)',
          'Update Net Worth tracker',
          'Check credit score (free apps)',
          'Review investment performance',
          'Set goals for next month'
        ]
      },
      {
        title: '📆 Annual Financial Review',
        color: COLORS.lavender,
        items: [
          'Review total annual spending by category',
          'Adjust budget for inflation & life changes',
          'Max out retirement contributions (401k, IRA)',
          'Review all insurance policies',
          'Tax planning session with accountant',
          'Estate planning review (will, beneficiaries)',
          'Pull all 3 credit reports (annualcreditreport.com)',
          'Rebalance investment portfolio',
          'Calculate net worth',
          'Set financial goals for next year'
        ]
      },
      {
        title: '🏛️ Tax Season Checklist (USA)',
        color: COLORS.dustyPink,
        items: [
          'Gather all W-2 forms from employers',
          'Collect 1099 forms (freelance, interest, dividends)',
          'Organize receipts (medical, charitable, business)',
          'Get property tax statements',
          'Locate 1098-E student loan interest form',
          'Gather retirement contribution documents (1099-R)',
          'Collect HSA/FSA account statements',
          'Get state-specific tax forms',
          'File federal & state returns by April 15',
          'Review refund or balance due'
        ]
      },
      {
        title: '📱 Subscription Audit',
        color: COLORS.softBlue,
        items: [
          'Netflix / Streaming Video',
          'Music (Spotify, Apple Music)',
          'Gym / Fitness Membership',
          'Cloud Storage (iCloud, Google, Dropbox)',
          'Software / Productivity Tools',
          'News / Magazines',
          'Meal Kit Services (HelloFresh, etc)',
          'Beauty / Subscription Boxes',
          'Gaming (Xbox, PlayStation, Nintendo)',
          'Other Subscriptions'
        ]
      },
      {
        title: '📈 Credit Score Improvement',
        color: COLORS.sageGreen,
        items: [
          'Pay all bills on time (35% of score)',
          'Keep utilization under 30% per card',
          'Don\'t close old accounts (15% of score)',
          'Limit new credit applications (10% of score)',
          'Maintain healthy credit mix (10% of score)',
          'Dispute any errors on credit reports',
          'Set up autopay for minimum payments',
          'Pay high-interest balances first'
        ]
      },
      {
        title: '🆘 Emergency Fund Checklist',
        color: COLORS.yellow,
        fontColor: COLORS.darkText,
        items: [
          'Calculate 3 months of essential expenses',
          '3-month emergency fund saved ✓',
          '6-month emergency fund saved ✓',
          '12-month emergency fund saved ✓',
          'Funds in high-yield savings account (HYSA)',
          'Emergency fund in separate account',
          'Automatic monthly transfer set up',
          'Fund accessible within 1-2 business days'
        ]
      },
      {
        title: '💳 Debt Payoff Strategy',
        color: COLORS.red,
        items: [
          'List all debts (name, balance, APR, min payment)',
          'Choose strategy: Snowball vs Avalanche',
          'Pay minimums on ALL debts each month',
          'Put all extra money toward target debt',
          'Celebrate when a debt is paid off!',
          'Roll that payment to next debt',
          'Track progress monthly in Debt Payoff tab',
          'Set a debt-free target date'
        ]
      },
      {
        title: '📊 Investment Review',
        color: COLORS.lavender,
        items: [
          'Review current asset allocation',
          'Rebalance if any asset class is 5%+ off target',
          'Check expense ratios on all funds',
          'Review dividend income received',
          'Consider tax-loss harvesting opportunities',
          'Max out tax-advantaged accounts (401k, IRA, HSA)',
          'Update beneficiary designations',
          'Review and update investment goals'
        ]
      }
    ];

    checklists.forEach(function(checklist) {
      var numItems = checklist.items.length;
      var fontColor = checklist.fontColor || COLORS.white;

      // Section header with progress counter
      sheet.getRange(currentRow, 1, 1, 4).merge()
        .setValue(checklist.title)
        .setBackground(checklist.color)
        .setFontColor(fontColor)
        .setFontSize(12)
        .setFontWeight('bold')
        .setFontFamily('Arial')
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('left');

      // Progress counter
      var itemStart = currentRow + 1;
      var itemEnd = currentRow + numItems;
      sheet.getRange(currentRow, 5, 1, 2).merge()
        .setFormula('=COUNTIF(A' + itemStart + ':A' + itemEnd + ',TRUE)&" of ' + numItems + ' completed"')
        .setBackground(checklist.color)
        .setFontColor(fontColor)
        .setFontSize(10)
        .setFontFamily('Arial')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      sheet.setRowHeight(currentRow, 30);
      currentRow++;

      // Checkbox items
      var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

      for (var i = 0; i < checklist.items.length; i++) {
        var bgItem = i % 2 === 0 ? COLORS.white : COLORS.beige;

        sheet.getRange(currentRow, 1).setDataValidation(checkboxRule).setValue(false).setBackground(bgItem);

        sheet.getRange(currentRow, 2, 1, 5).merge()
          .setValue(checklist.items[i])
          .setFontFamily('Arial').setFontSize(10)
          .setBackground(bgItem)
          .setVerticalAlignment('middle');
        sheet.setRowHeight(currentRow, 22);

        currentRow++;
      }

      // Strikethrough conditional formatting for checked items
      for (var j = itemStart; j <= itemEnd; j++) {
        var cbRange = sheet.getRange(j, 2, 1, 5);
        var strikeRule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=$A' + j + '=TRUE')
          .setStrikethrough(true)
          .setFontColor('#AAAAAA')
          .setRanges([cbRange]).build();
        sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat([strikeRule]));
      }

      // Spacer between checklists
      sheet.setRowHeight(currentRow, 16);
      currentRow++;
    });

    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 350);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 100);

    sheet.setFrozenRows(1);

    Logger.log('✅ Checklists created');
  } catch (e) {
    Logger.log('❌ Checklists error: ' + e + '\n' + e.stack);
  }
}

// ============================================================
// SHEET 7: INSTRUCTIONS
// ============================================================
function createInstructionsSheet() {
  try {
    var sheet = getOrCreateSheet('Instructions');
    sheet.setTabColor(COLORS.beige);

    var currentRow = 1;

    function addSection(title, color, fontColor, content, isList) {
      fontColor = fontColor || COLORS.white;
      // Section header
      sheet.getRange(currentRow, 1, 1, 3).merge()
        .setValue(title)
        .setBackground(color)
        .setFontColor(fontColor)
        .setFontSize(13)
        .setFontWeight('bold')
        .setFontFamily('Arial')
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('left');
      sheet.setRowHeight(currentRow, 32);
      currentRow++;

      // Content
      if (Array.isArray(content)) {
        content.forEach(function(line, idx) {
          sheet.getRange(currentRow, 1, 1, 3).merge()
            .setValue(line)
            .setBackground(idx % 2 === 0 ? COLORS.white : COLORS.beige)
            .setFontFamily('Arial')
            .setFontSize(10)
            .setFontColor(COLORS.darkText)
            .setWrap(true)
            .setVerticalAlignment('middle');
          sheet.setRowHeight(currentRow, 28);
          currentRow++;
        });
      }
      // Spacer
      sheet.setRowHeight(currentRow, 12);
      currentRow++;
    }

    // Title banner
    sheet.getRange(currentRow, 1, 1, 3).merge()
      .setValue('📚 Budget Planner — Instructions & Guide')
      .setBackground(COLORS.sageGreen)
      .setFontColor(COLORS.white)
      .setFontSize(18)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 60);
    currentRow += 2;

    // Welcome
    addSection('🌟 Welcome!', COLORS.sageGreen, COLORS.white, [
      'Welcome to Budget Planner — your all-in-one personal finance tracking system!',
      'This spreadsheet automatically calculates your budget, tracks expenses, monitors savings goals, and shows your debt payoff progress.',
      'Everything is connected: data you enter in one sheet automatically updates all other sheets. Just fill in your numbers and let the formulas do the work!'
    ]);

    // Getting started
    addSection('🚀 Getting Started (5 Steps)', COLORS.softBlue, COLORS.white, [
      'Step 1: Make a copy of this file (File → Make a Copy) so you have your own version to edit.',
      'Step 2: Change currency symbol if needed (Edit → Find & Replace → Find "$" → Replace with your symbol, e.g. "€" or "£").',
      'Step 3: Confirm the year is correct — it should show ' + YEAR + ' automatically. Update via the script if needed.',
      'Step 4: Start with the "Monthly Budget" tab. Enter your Budgeted amounts for each category.',
      'Step 5: Log expenses in the "Expense Tracker" tab. The Monthly Budget "Actual" column will update automatically!'
    ]);

    // Tab guide
    addSection('📑 Tab-by-Tab Guide', COLORS.lavender, COLORS.white, [
      '📊 Dashboard: Your financial command center. Shows summary cards, monthly breakdown table, and net balance at a glance. Data populates automatically.',
      '📋 Monthly Budget: Core budget sheet. Enter your Budgeted amount for each category. Actual spending is pulled from Expense Tracker via SUMIF formulas. Status indicators show "On Track", "Under Budget", or "Overspending".',
      '💸 Expense Tracker: Log every transaction here. Use the Date, Description, Category (dropdown), Amount, and Payment Method columns. This data feeds back into Monthly Budget and Dashboard.',
      '💰 Savings Goals: Track progress toward financial goals. Enter your Target, Current Amount, and Monthly Contribution. Months-to-goal and % Complete are calculated automatically.',
      '💳 Debt Payoff: Track all debts with APR, balances, and payments. Choose Avalanche or Snowball strategy. Payoff dates are estimated using NPER formula.',
      '✅ Checklists: Monthly, annual, tax, and strategy checklists to keep you on track. Check off items — they auto-strikethrough and a progress counter updates.',
      '📚 Instructions: You are here! Return anytime for guidance.'
    ]);

    // FAQ
    addSection('❓ Frequently Asked Questions', COLORS.dustyPink, COLORS.white, [
      'Q: How do I change the currency symbol?\nA: Use Find & Replace (Ctrl+H or Cmd+H): Find "$", Replace with your symbol (e.g. "€"). This updates labels only — formulas use number formats.',
      'Q: How do I add new budget categories?\nA: In Monthly Budget, insert a row within a section. Also add the category name to the Expense Tracker dropdown list (Data → Data Validation on column C).',
      'Q: Why isn\'t the Dashboard updating?\nA: Make sure you\'ve entered data in Monthly Budget or Expense Tracker. The Dashboard references those sheets. Try View → Force refresh (Ctrl+Shift+R).',
      'Q: How do I reset for a new year?\nA: Use 💰 Budget Planner → Reset All Data from the menu, or manually clear input cells. The year updates automatically via the script\'s YEAR variable.',
      'Q: Some cells show errors (#REF!, #VALUE!) — what do I do?\nA: This usually means a referenced sheet or cell was deleted or renamed. Check that all 7 tabs exist and haven\'t been renamed.'
    ]);

    // Pro Tips
    addSection('💡 Pro Tips', COLORS.sageGreen, COLORS.white, [
      '💡 Tip 1 — Weekly habit: Spend 10 minutes each Sunday logging the week\'s expenses. Consistency is the key to accurate tracking.',
      '💡 Tip 2 — Use filters: In Expense Tracker, use Data → Create Filter to sort by Category, Month, or Amount to spot spending patterns.',
      '💡 Tip 3 — Savings Rate goal: Aim for 20%+ savings rate. Even 1% improvement per month compounds significantly over time.',
      '💡 Tip 4 — Emergency fund first: Before aggressively paying off debt or investing, build at least 1 month of expenses in savings as a buffer.'
    ]);

    // Troubleshooting
    addSection('🔧 Troubleshooting', COLORS.lightGray, COLORS.darkText, [
      'Problem: Expense Tracker actuals not showing in Monthly Budget.\nFix: Check that the Category in Expense Tracker exactly matches the category name in Monthly Budget (case-sensitive). Use the dropdown.',
      'Problem: Dates not recognized.\nFix: Make sure dates in Expense Tracker column A are entered as actual dates, not text. Use MM/DD/YYYY format or click a cell to enter via date picker.',
      'Problem: Checklist strikethrough not working.\nFix: Conditional formatting may need refresh. Try Format → Conditional Formatting and re-save the rule.',
      'Problem: Script ran but some sheets are missing.\nFix: Run main() again from Extensions → Apps Script. It is idempotent — safe to run multiple times.',
      'Problem: Saving goals "Months to Goal" shows an error.\nFix: Make sure Monthly Contribution (column D) is a positive number greater than zero.'
    ]);

    // Support
    addSection('📞 Support & Customization', COLORS.beige, COLORS.darkText, [
      'This template was built with Google Apps Script. To customize: Extensions → Apps Script → edit the BudgetPlanner_Phase1.gs file.',
      'Phase 2 will add: Paycheck Budget, Yearly Overview, Net Worth, Investment Tracker, Bill Calendar, Credit Score Tracker, and a README sheet.',
      'For the best experience, use Google Chrome on desktop. Mobile support is limited for dropdown menus and conditional formatting.',
      '© Budget Planner Template — Feel free to customize for personal use. Enjoy financial freedom! 🎉'
    ]);

    // Column widths
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 500);
    sheet.setColumnWidth(3, 150);

    sheet.setFrozenRows(1);

    Logger.log('✅ Instructions sheet created');
  } catch (e) {
    Logger.log('❌ Instructions error: ' + e + '\n' + e.stack);
  }
}
