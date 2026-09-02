// ------------------------------------------------------------
// CUSTOM SHEETS MENU
//
// Adds a "🏈 Pick 'Em Admin" menu to the Sheets toolbar with
// grouped submenus, so you don't have to hunt through the
// function dropdown in the Apps Script editor for common tasks.
//
// Requires this file to live in the SAME Apps Script project as
// Code.gs and Automation.gs - it calls their functions directly,
// since all .gs files in one project share the same global scope.
//
// The menu appears automatically next time you open the
// spreadsheet (onOpen runs on open). If you don't want to wait,
// just run onOpen once manually from the function dropdown.
// ------------------------------------------------------------

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🏈 Pick 'Em Admin")
    .addSubMenu(
      ui.createMenu("Schedule")
        .addItem("Pull Regular Season...", "menuPullRegularSeason")
        .addItem("Pull Preseason...", "menuPullPreseason")
        .addItem("Remove Preseason...", "menuRemovePreseason")
        .addItem("Check Schedule Row Counts", "menuCheckScheduleRowCounts")
    )
        .addSubMenu(
      ui.createMenu("Results & Odds")
        .addItem("Pull Final Scores...", "menuPullFinalScores")
        .addItem("Pull Odds Snapshot...", "menuPullOddsForWeek")
        .addItem("Pull Current Odds Snapshot (all weeks, regular season)", "menuSnapshotOdds")
        .addSeparator()
        .addItem("Install Results Triggers (Fri/Mon/Tue 5am)", "menuInstallResultsTriggers")
        .addItem("Remove Results Triggers", "menuRemoveResultsTriggers")
        .addItem("Install Odds Triggers (Tue/Wed)", "menuInstallOddsTriggers")
        .addItem("Remove Odds Triggers", "menuRemoveOddsTriggers")
    )
    .addSubMenu(
      ui.createMenu("Parlay / Tie Break")
        .addItem("Recompute All Parlay Picks", "menuRecomputeParlay")
        .addItem("Check Tie Break Vote Status", "menuCheckTieBreakVoteStatus")
        .addItem("Check Tie Break Flags", "menuCheckTieBreakFlags")
    )
    .addSubMenu(
      ui.createMenu("Pick Window")
        .addItem("Open Next Week Today...", "menuOpenNextWeekToday")
        .addItem("Extend Current Week Through Tomorrow", "menuExtendThroughTomorrow")
        .addItem("Extend Current Week to Thursday", "menuExtendToThursday")
        .addItem("Check Pick Window Extension", "menuCheckPickWindowExtension")
    )
    .addSubMenu(
      ui.createMenu("Season Setup")
        .addItem("Reset Season Testing Data...", "menuResetSeasonTestingData")
        .addItem("Run User Sheet Migrations", "menuRunUserMigrations")
    )
    .addToUi();
}

// ------------------------------------------------------------
// Small shared helpers for prompting/alerting
// ------------------------------------------------------------

function promptForYear_(ui, defaultYear) {
  const result = ui.prompt("Which year?", "e.g. " + defaultYear, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  const text = result.getResponseText().trim();
  return text ? Number(text) : defaultYear;
}

function promptForWeek_(ui) {
  const result = ui.prompt("Which week?", "e.g. 9 for regular season, or P1-P4 for preseason", ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  const text = result.getResponseText().trim();
  return text || null;
}

// ------------------------------------------------------------
// SCHEDULE
// ------------------------------------------------------------

function menuPullRegularSeason() {
  const ui = SpreadsheetApp.getUi();
  const year = promptForYear_(ui, 2026);
  if (!year) return;

  const result = pullNFLScheduleUpsert(year);
  ui.alert("Pull Regular Season - " + year,
    "Status: " + result.status +
    "\nUpdated: " + (result.updated || 0) +
    "\nInserted: " + (result.inserted || 0) +
    (result.message ? "\n" + result.message : ""),
    ui.ButtonSet.OK);
}

function menuPullPreseason() {
  const ui = SpreadsheetApp.getUi();
  const year = promptForYear_(ui, 2026);
  if (!year) return;

  const result = pullPreseasonIntoScheduleForTesting(year);
  ui.alert("Pull Preseason - " + year,
    "Status: " + result.status +
    "\nUpdated: " + (result.updated || 0) +
    "\nInserted: " + (result.inserted || 0) +
    (result.message ? "\n" + result.message : ""),
    ui.ButtonSet.OK);
}

function menuRemovePreseason() {
  const ui = SpreadsheetApp.getUi();
  const year = promptForYear_(ui, 2026);
  if (!year) return;

  const confirm = ui.alert("Remove Preseason - " + year,
    "This will delete all PRESEASON-tagged rows for " + year + " from NFL_Schedule. Regular season and bye rows are untouched. Continue?",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const result = removePreseasonRowsFromSchedule(year);
  ui.alert("Removed " + result.removed + " preseason row(s) for " + year + ".");
}

function menuCheckScheduleRowCounts() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NFL_Schedule");
  if (!sheet) {
    ui.alert("NFL_Schedule sheet not found.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colYear = headers.indexOf("SeasonYear");
  const colRowType = headers.indexOf("RowType");

  const counts = {};
  for (let i = 1; i < data.length; i++) {
    const year = data[i][colYear] || "(blank)";
    const rowType = data[i][colRowType];
    if (!counts[year]) counts[year] = { GAME: 0, BYE: 0, PRESEASON: 0, other: 0 };
    if (counts[year][rowType] !== undefined) counts[year][rowType]++;
    else counts[year].other++;
  }

  const lines = Object.keys(counts).sort().map(year => {
    const c = counts[year];
    return year + ": GAME=" + c.GAME + ", BYE=" + c.BYE + ", PRESEASON=" + c.PRESEASON + ", other=" + c.other;
  });

  ui.alert("Schedule Row Counts", lines.join("\n") || "No data found.", ui.ButtonSet.OK);
}

// ------------------------------------------------------------
// RESULTS & ODDS
// ------------------------------------------------------------

function menuPullFinalScores() {
  const ui = SpreadsheetApp.getUi();
  const year = promptForYear_(ui, 2026);
  if (!year) return;
  const week = promptForWeek_(ui);
  if (!week) return;

  const result = pullFinalScores(year, week);
  ui.alert("Pull Final Scores - " + year + " Week " + week,
    "Status: " + result.status +
    "\nUpdated: " + (result.updated || 0) +
    "\nSkipped (incomplete): " + (result.skippedIncomplete || 0) +
    (result.message ? "\n" + result.message : ""),
    ui.ButtonSet.OK);
}

function menuPullOddsForWeek() {
  const ui = SpreadsheetApp.getUi();
  const year = promptForYear_(ui, 2026);
  if (!year) return;
  const week = promptForWeek_(ui);
  if (!week) return;

  const result = snapshotOddsForWeek_(year, week, false, false);
  ui.alert("Pull Odds Snapshot - " + year + " Week " + week,
    "Status: " + result.status +
    "\nRows added: " + (result.rowsAdded || 0) +
    (result.message ? "\n" + result.message : ""),
    ui.ButtonSet.OK);
}
function menuInstallResultsTriggers() {
  const ui = SpreadsheetApp.getUi();
  const result = installResultsTriggers();
  ui.alert("Install Results Triggers",
    "Status: " + result.status +
    "\nTimezone: " + result.timezone +
    "\n\nTriggers set for Fri 5am, Mon 5am, Tue 5am.",
    ui.ButtonSet.OK);
}

function menuRemoveResultsTriggers() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert("Remove Results Triggers",
    "This stops automatic final-score pulls (Fri/Mon/Tue 5am). Continue?",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const result = removeResultsTriggers();
  ui.alert("Removed " + result.removed + " results-pull trigger(s).");
}

function menuInstallOddsTriggers() {
  const ui = SpreadsheetApp.getUi();
  const result = installOddsTriggers();
  ui.alert("Install Odds Triggers",
    "Status: " + result.status +
    "\nTimezone: " + result.timezone +
    "\n\nTriggers set for Tue 12:05am, Tue 6am, Tue 4pm, Wed 6am, Wed 3pm.",
    ui.ButtonSet.OK);
}

function menuRemoveOddsTriggers() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert("Remove Odds Triggers",
    "This stops automatic odds pulls (Tue/Wed). Continue?",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const result = removeOddsTriggers();
  ui.alert("Removed " + result.removed + " odds-pull trigger(s).");
}
function menuSnapshotOdds() {
  const ui = SpreadsheetApp.getUi();
  snapshotNFLCurrentOdds();
  ui.alert("Odds snapshot complete - check the Execution Log for details.");
}

// ------------------------------------------------------------
// PARLAY / TIE BREAK
// ------------------------------------------------------------

function menuRecomputeParlay() {
  const ui = SpreadsheetApp.getUi();
  const result = generateAllParlayPicksAllYearsFast();
  ui.alert("Recompute All Parlay Picks",
    "Min year recomputed: " + result.minYear +
    "\nRows removed & rewritten: " + result.removedInScopeRows +
    "\nFresh rows written: " + result.freshRowsWritten +
    "\nYear/week combinations: " + result.yearWeeksProcessed +
    "\n\n(Years before " + result.minYear + " were never touched.)",
    ui.ButtonSet.OK);
}

function menuCheckTieBreakVoteStatus() {
  const ui = SpreadsheetApp.getUi();
  checkTieBreakVoteStatus();
  ui.alert("Done - check the Execution Log for the full vote breakdown.");
}

function menuCheckTieBreakFlags() {
  const ui = SpreadsheetApp.getUi();
  checkTieBreakFlags();
  ui.alert("Done - check the Execution Log for the flag/credential details.");
}

// ------------------------------------------------------------
// PICK WINDOW
// ------------------------------------------------------------

function menuOpenNextWeekToday() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert("Open Next Week Today",
    "This opens the NEXT week's pick window starting today, ahead of its normal Tuesday open. Continue?",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const result = openNextWeekPickWindowToday();
  ui.alert("Open Next Week Today",
    result.status === "ok"
      ? ("Week " + result.week + " (" + result.year + ") now opens " + result.newOpeningDate + ".")
      : ("Error: " + result.message),
    ui.ButtonSet.OK);
}

function menuExtendThroughTomorrow() {
  const ui = SpreadsheetApp.getUi();
  const result = extendCurrentWeekPickWindowThroughTomorrow();
  ui.alert("Extend Pick Window",
    result.status === "ok"
      ? ("Week " + result.week + " (" + result.year + ") now closes " + result.newClosingDate + ".")
      : ("Error: " + result.message),
    ui.ButtonSet.OK);
}

function menuExtendToThursday() {
  const ui = SpreadsheetApp.getUi();
  const result = extendCurrentWeekPickWindowToThursday();
  ui.alert("Extend Pick Window",
    result.status === "ok"
      ? ("Week " + result.week + " (" + result.year + ") now closes " + result.newClosingDate + ".")
      : ("Error: " + result.message),
    ui.ButtonSet.OK);
}

function menuCheckPickWindowExtension() {
  const ui = SpreadsheetApp.getUi();
  const result = checkPickWindowExtension();
  ui.alert("Pick Window Status",
    "Year/Week: " + result.year + " / " + result.week +
    "\nExtension: " + (result.extensionValue || "(none set)") +
    "\nCurrent window: " + JSON.stringify(result.pickWindow),
    ui.ButtonSet.OK);
}

// ------------------------------------------------------------
// SEASON SETUP
// ------------------------------------------------------------

function menuResetSeasonTestingData() {
  const ui = SpreadsheetApp.getUi();
  const year = promptForYear_(ui, 2026);
  if (!year) return;

  const confirm = ui.alert("Reset Season Testing Data - " + year,
    "This clears Picks, TieBreakPicks, LeaderOverrides, ParlayPicks, and API_NFL_Results for " + year + " ONLY. " +
    "Other years, NFL_Schedule, Users, and API_NFL_OddsHistory are never touched. This cannot be undone. Continue?",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const result = resetSeasonTestingData(year);
  ui.alert("Reset complete for " + year + " - check the Execution Log for the full breakdown.");
}

function menuRunUserMigrations() {
  const ui = SpreadsheetApp.getUi();
  const results = [];

  results.push("Phone: " + addPhoneColumnToUsers().message);
  results.push("SendText: " + addSendTextColumnToUsers().message);
  results.push("Notification prefs: " + addNotificationPreferenceColumns().message);
  results.push("Display names: " + addDisplayNameColumnToUsers().message);

  ui.alert("User Sheet Migrations", results.join("\n"), ui.ButtonSet.OK);
}