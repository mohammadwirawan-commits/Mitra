/**
 * Mitra Commission Intake — monthly cleanup tool (Google Apps Script)
 * ---------------------------------------------------------------------------
 * Adds an "Intake Tools ▸ Clear last month's entries" menu to the shared intake
 * Google Sheet. Clearing wipes the pasted data (everything below the header row)
 * in the stakeholder-filled tabs, while PRESERVING:
 *   - the header row (row 1) and all cell formatting / colors / data-validation,
 *   - the reference "do-not-edit" tabs (3. Region Split, SSB),
 *   - the README tab.
 *
 * INSTALL (one time):
 *   1. Open the intake Google Sheet.
 *   2. Extensions ▸ Apps Script.
 *   3. Delete any placeholder code, paste ALL of this file, click Save.
 *   4. Reload the Sheet. A new "Intake Tools" menu appears next to Help.
 *   5. First click runs an authorization prompt (approve for your account).
 *
 * USE (each month, after the workbook has been generated from last month's data):
 *   Intake Tools ▸ Clear last month's entries ▸ confirm.
 *
 * SAFETY: the action asks for confirmation and lists exactly which tabs it will
 * clear. It cannot be undone from the button, but Google Sheets keeps full history
 * under File ▸ Version history ▸ See version history if you ever need to restore.
 */

// Tabs that must NEVER be cleared (matched by exact name — edit if you rename them).
var PRESERVE_TABS = ['README', '3. Region Split', 'SSB'];

// Number of header rows to keep at the top of each cleared tab.
var HEADER_ROWS = 1;


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Intake Tools')
    .addItem("Clear last month's entries", 'clearLastMonthEntries')
    .addToUi();
}


function clearLastMonthEntries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var toClear = ss.getSheets().filter(function (sh) {
    return PRESERVE_TABS.indexOf(sh.getName()) === -1;
  });
  var names = toClear.map(function (sh) { return sh.getName(); });

  var resp = ui.alert(
    "Clear last month's entries?",
    'This clears the pasted data (below the header row) in these ' + names.length + ' tab(s):\n\n' +
      names.join(', ') + '\n\n' +
      'Preserved (untouched): ' + PRESERVE_TABS.join(', ') + '\n\n' +
      'Headers, formatting, and the reference tabs are kept.\n' +
      '(Restore later via File ▸ Version history if needed.)\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (resp !== ui.Button.YES) {
    ui.alert('Cancelled — nothing was changed.');
    return;
  }

  var tabsCleared = 0, rowsCleared = 0;
  for (var i = 0; i < toClear.length; i++) {
    var sh = toClear[i];
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow > HEADER_ROWS && lastCol > 0) {
      var n = lastRow - HEADER_ROWS;
      // clearContent() removes values only — keeps formatting, colors and validation.
      sh.getRange(HEADER_ROWS + 1, 1, n, lastCol).clearContent();
      rowsCleared += n;
      tabsCleared++;
    }
  }

  ui.alert(
    'Done.',
    'Cleared ' + rowsCleared + ' row(s) across ' + tabsCleared + ' tab(s).\n' +
      'Preserved: header rows, formatting, and ' + PRESERVE_TABS.join(' / ') + '.',
    ui.ButtonSet.OK);
}
