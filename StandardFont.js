// ------------------------------------------------------------
// STANDARD FONT - Calibri 10, applied to a generous pre-formatted
// range so future writes (setValues only changes cell VALUES,
// never font) automatically display correctly without needing to
// touch every write site in the codebase.
// ------------------------------------------------------------
function applyStandardFont_(sheet) {
  if (!sheet) return;
  const numRows = 20000; // generous multi-year buffer
  const numCols = 20;
  sheet.getRange(1, 1, numRows, numCols)
    .setFontFamily("Calibri")
    .setFontSize(10);
}

/**
 * ONE-TIME MIGRATION - applies Calibri 10 to every sheet this app
 * manages, whether it already has data or not. Run this once after
 * adding applyStandardFont_ to retroactively cover sheets that
 * existed before this change. Safe to re-run anytime.
 */
function applyStandardFontToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [
    SHEET_USERS, SHEET_PICKS, SHEET_RESULTS, SHEET_PARLAY,
    SHEET_TIEBREAK, SHEET_LEADER_OVERRIDE, SHEET_SCHEDULE, SHEET_ODDS
  ];

  const results = [];
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      applyStandardFont_(sheet);
      results.push(name + ": formatted");
    } else {
      results.push(name + ": not found, skipped");
    }
  });

  Logger.log(results.join("\n"));
  return { status: "ok", results: results };
}