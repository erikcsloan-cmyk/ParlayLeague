/**
 * ============================================================
 * NFL PICK 'EM APP - Server-side Code (Google Apps Script)
 * ============================================================
 * Sheets used:
 *   API_NFL_Schedule   (already exists - built by pullNFLRegularSeasonToSheet)
 *   API_NFL_Results    (created automatically by this file - final scores)
 *   Users              (created automatically by this file - list of players)
 *   Picks              (created automatically by this file - who picked what)
 * ============================================================
 */


const SHEET_SCHEDULE = "NFL_Schedule";
const SHEET_RESULTS  = "API_NFL_Results";
const SHEET_USERS    = "Users";
const SHEET_PICKS    = "Picks";
const SHEET_PARLAY   = "ParlayPicks";
const SHEET_ODDS     = "API_NFL_OddsHistory";
const SHEET_TIEBREAK = "TieBreakPicks";
const SHEET_LEADER_OVERRIDE = "LeaderOverrides";
const SHEET_NOTICES = "Notices";
const SHEET_NOTICE_DISMISSALS = "NoticeDismissals";
const SHEET_USER_FEED_STATE = "UserFeedState";
const SHEET_RUNBACK_VOTES = "RunbackVotes";


const PARLAY_USER = "Parlay"; // not a real login user - computed majority vote


// ------------------------------------------------------------
// WEB APP ENTRY POINT
// ------------------------------------------------------------
function doGet(e) {
  // Public guide page - no login/authentication needed, shareable as a
  // direct link. Anything else falls through to the normal app.
  if (e && e.parameter && e.parameter.page === "guide") {
    return HtmlService.createTemplateFromFile("Guide")
      .evaluate()
      .setTitle("Parlay Pick 'Em - League Guide")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }


  ensureSupportSheets();
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Parlay League")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function getTeamRecords_(year) {
  const results = sheetToObjects_(SHEET_RESULTS)
    .filter(r =>
      Number(r.SeasonYear) === Number(year) &&
      String(r.Week).charAt(0) !== "P" &&
      r.WinningTeam
    );


  const records = {};
  function bump(team, field) {
    if (!records[team]) records[team] = { wins: 0, losses: 0, ties: 0 };
    records[team][field]++;
  }


  results.forEach(r => {
    const parts = String(r.Matchup || "").split(" @ ");
    if (parts.length !== 2) return;
    const away = parts[0].trim();
    const home = parts[1].trim();
    if (!away || !home) return;


    if (r.WinningTeam === "TIE") {
      bump(away, "ties");
      bump(home, "ties");
    } else if (r.WinningTeam === away) {
      bump(away, "wins");
      bump(home, "losses");
    } else if (r.WinningTeam === home) {
      bump(home, "wins");
      bump(away, "losses");
    }
  });


  return records;
}


function formatTeamRecord_(rec) {
  if (!rec) return "";
  return rec.ties > 0
    ? rec.wins + "-" + rec.losses + "-" + rec.ties
    : rec.wins + "-" + rec.losses;
}




// ------------------------------------------------------------
// SETUP - creates the Users / Picks / Results sheets if missing
// Safe to call repeatedly - only creates what's missing.
// ------------------------------------------------------------
// Apps Script globals reset on every execution, so this flag means "already
// confirmed the support sheets exist during THIS server call". Several
// functions each call ensureSupportSheets() (getWeekData, then
// computeParlayPicksForWeek, ...), and the real check costs a getSheetByName
// per sheet - once per call is plenty.
let supportSheetsEnsuredThisRun_ = false;

function ensureSupportSheets() {
  if (supportSheetsEnsuredThisRun_) return;
  ensureSupportSheetsImpl_();
  supportSheetsEnsuredThisRun_ = true;
}

function ensureSupportSheetsImpl_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();


  // USERS
  let usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(SHEET_USERS);
    usersSheet.getRange(1, 1, 1, 9)
      .setValues([[
        "UserName", "Active", "SortOrder", "Phone", "DisplayName",
        "SendText", "SendTieInfo", "SendRecapInfo", "SendStillTied"
      ]])
      .setFontWeight("bold");
    usersSheet.getRange(2, 1, 5, 9).setValues([
      ["Player 1", true, 1, "", "", true, true, true, true],
      ["Player 2", true, 2, "", "", true, true, true, true],
      ["Player 3", true, 3, "", "", true, true, true, true],
      ["Player 4", true, 4, "", "", true, true, true, true],
      ["Player 5", true, 5, "", "", true, true, true, true]
    ]);
    applyStandardFont_(usersSheet);
  }


  // PICKS
  let picksSheet = ss.getSheetByName(SHEET_PICKS);
  if (!picksSheet) {
    picksSheet = ss.insertSheet(SHEET_PICKS);
    picksSheet.getRange(1, 1, 1, 8)
      .setValues([[
        "PickID", "SeasonYear", "Week", "EventID",
        "Matchup", "UserName", "PickedTeam", "SubmittedAt"
      ]])
      .setFontWeight("bold");
    applyStandardFont_(picksSheet);
  }


  // RESULTS
  let resultsSheet = ss.getSheetByName(SHEET_RESULTS);
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet(SHEET_RESULTS);
    resultsSheet.getRange(1, 1, 1, 8)
      .setValues([[
        "SeasonYear", "Week", "EventID", "Matchup",
        "AwayScore", "HomeScore", "WinningTeam", "UpdatedAt"
      ]])
      .setFontWeight("bold");
    applyStandardFont_(resultsSheet);
  }


  // PARLAY PICKS (computed majority vote - not entered by a real user)
  let parlaySheet = ss.getSheetByName(SHEET_PARLAY);
  if (!parlaySheet) {
    parlaySheet = ss.insertSheet(SHEET_PARLAY);
    parlaySheet.getRange(1, 1, 1, 8)
      .setValues([[
        "SeasonYear", "Week", "EventID", "Matchup",
        "PickedTeam", "VoteCount", "TotalVotes", "GeneratedAt"
      ]])
      .setFontWeight("bold");
    applyStandardFont_(parlaySheet);
  }


  // TIE BREAK PICKS - a separate re-vote, only used for games where the
  // main Parlay vote comes out tied. Doesn't touch anyone's original pick.
  let tieBreakSheet = ss.getSheetByName(SHEET_TIEBREAK);
  if (!tieBreakSheet) {
    tieBreakSheet = ss.insertSheet(SHEET_TIEBREAK);
    tieBreakSheet.getRange(1, 1, 1, 8)
      .setValues([[
        "TieBreakID", "SeasonYear", "Week", "EventID",
        "Matchup", "UserName", "PickedTeam", "SubmittedAt"
      ]])
      .setFontWeight("bold");
    applyStandardFont_(tieBreakSheet);
  }


  // LEADER OVERRIDES - a fallback for games still tied after a full
  // re-vote round. The current season standings leader can make the
  // final call, which takes top priority over any vote tally.
  let leaderOverrideSheet = ss.getSheetByName(SHEET_LEADER_OVERRIDE);
  if (!leaderOverrideSheet) {
    leaderOverrideSheet = ss.insertSheet(SHEET_LEADER_OVERRIDE);
    leaderOverrideSheet.getRange(1, 1, 1, 6)
      .setValues([[
        "SeasonYear", "Week", "EventID", "Matchup", "PickedTeam", "DecidedBy"
      ]])
      .setFontWeight("bold");
    applyStandardFont_(leaderOverrideSheet);
  }
}


// ------------------------------------------------------------
// ONE-TIME REPAIR - fills in any blank PickID cells using
// EventID + UserName. Independent of backfillPickEventIds, so
// safe to run even if EventID was already fixed manually.
// Safe to re-run; only touches blank PickID cells.
// ------------------------------------------------------------
function backfillMissingPickIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PICKS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: "ok", fixed: 0, skipped: [] };
  }


  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colPickId = headers.indexOf("PickID");
  const colEventId = headers.indexOf("EventID");
  const colUser = headers.indexOf("UserName");


  let fixed = 0;
  const skipped = [];


  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[colPickId] && String(row[colPickId]).trim() !== "") continue;


    const eventId = row[colEventId];
    const userName = row[colUser];


    if (!eventId || eventId === "GAME" || !userName) {
      skipped.push("Row " + (i + 1) + ": missing EventID or UserName");
      continue;
    }


    sheet.getRange(i + 1, colPickId + 1).setValue(eventId + "-" + userName);
    fixed++;
  }


  return { status: "ok", fixed: fixed, skipped: skipped };
}


// ------------------------------------------------------------
// ONE-TIME MIGRATION - run manually if your Users sheet was
// created before the Phone column existed.
// Safe to re-run; does nothing if Phone already exists.
// ------------------------------------------------------------
function addPhoneColumnToUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };


  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("Phone") !== -1) {
    return { status: "ok", message: "Phone column already exists" };
  }


  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue("Phone").setFontWeight("bold");
  return { status: "ok", message: "Phone column added - fill in each user's number" };
}


function addResolvedViaColumnToParlayPicks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PARLAY);
  if (!sheet) return { status: "error", message: "ParlayPicks sheet not found" };


  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("ResolvedVia") !== -1) {
    return { status: "ok", message: "ResolvedVia column already exists" };
  }


  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue("ResolvedVia").setFontWeight("bold");
  return { status: "ok", message: "ResolvedVia column added" };
}


// ------------------------------------------------------------
// ONE-TIME MIGRATION - run manually if your Users sheet was
// created before the SendText column existed. Existing rows are
// defaulted to TRUE so nobody who was already getting texts
// silently stops - opt individual users out afterward as needed.
// Safe to re-run; does nothing if SendText already exists.
// ------------------------------------------------------------
function addSendTextColumnToUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };


  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("SendText") !== -1) {
    return { status: "ok", message: "SendText column already exists" };
  }


  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue("SendText").setFontWeight("bold");


  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const defaults = [];
    for (let i = 0; i < lastRow - 1; i++) defaults.push([true]);
    sheet.getRange(2, newCol, lastRow - 1, 1).setValues(defaults);
  }


  return { status: "ok", message: "SendText column added - existing users defaulted to TRUE" };
}


// ------------------------------------------------------------
// ONE-TIME MIGRATION - run manually to add the three granular
// notification-type columns (SendTieInfo, SendRecapInfo,
// SendStillTied) alongside the existing SendText master switch.
// Existing users default to TRUE on all three, so nobody who was
// already getting texts silently loses any message type - turn
// individual ones off afterward as wanted.
// Safe to re-run; only adds columns that don't already exist.
// ------------------------------------------------------------
function addNotificationPreferenceColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };


  const columnsToAdd = ["SendTieInfo", "SendRecapInfo", "SendStillTied"];
  const added = [];


  columnsToAdd.forEach(colName => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf(colName) !== -1) return; // already exists, skip


    const newCol = headers.length + 1;
    sheet.getRange(1, newCol).setValue(colName).setFontWeight("bold");


    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const defaults = [];
      for (let i = 0; i < lastRow - 1; i++) defaults.push([true]);
      sheet.getRange(2, newCol, lastRow - 1, 1).setValues(defaults);
    }


    added.push(colName);
  });


  return { status: "ok", added: added, message: added.length ? ("Added: " + added.join(", ")) : "All columns already existed" };
}


// ------------------------------------------------------------
// ONE-TIME MIGRATION - adds a DisplayName column so users can
// have a fun nickname shown in the app instead of their login
// UserName. Purely cosmetic - UserName stays the real data key
// used everywhere (Picks, Standings, login matching, etc.), this
// never touches any of that. Existing users default to blank,
// meaning they'll keep showing their UserName until you fill one in.
// Safe to re-run; does nothing if DisplayName already exists.
// ------------------------------------------------------------
function addDisplayNameColumnToUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };


  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("DisplayName") !== -1) {
    return { status: "ok", message: "DisplayName column already exists" };
  }


  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue("DisplayName").setFontWeight("bold");
  return { status: "ok", message: "DisplayName column added - fill in nicknames as you like, blank falls back to UserName" };
}


// ------------------------------------------------------------
// ONE-TIME MIGRATION - adds a ShortName column for compact display
// (e.g. names-row on Parlay Pick). Optional per-person override;
// blank falls back to the first 4 letters of DisplayName/UserName.
// Safe to re-run; does nothing if ShortName already exists.
// ------------------------------------------------------------
function addShortNameColumnToUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };


  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("ShortName") !== -1) {
    return { status: "ok", message: "ShortName column already exists" };
  }


  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue("ShortName").setFontWeight("bold");
  return { status: "ok", message: "ShortName column added - fill in a compact name (e.g. 'Chan'), blank falls back to first 4 letters of DisplayName" };
}
// ------------------------------------------------------------
// ONE-TIME MIGRATION - adds an IsAdmin column to Users (TRUE/FALSE).
// Controls visibility of the in-app Admin tab. Blank/FALSE = not
// an admin. Safe to re-run; does nothing if IsAdmin already exists.
// ------------------------------------------------------------
function addIsAdminColumnToUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };


  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("IsAdmin") !== -1) {
    return { status: "ok", message: "IsAdmin column already exists" };
  }


  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue("IsAdmin").setFontWeight("bold");
  return { status: "ok", message: "IsAdmin column added - set TRUE for admins" };
}


/**
 * Whether a given UserName is flagged as an admin. Used to gate
 * the in-app Admin tab and its backend actions.
 */
function isUserAdmin_(userName) {
  const users = sheetToObjects_(SHEET_USERS);
  const u = users.find(u => u.UserName === userName);
  if (!u) return false;
  return u.IsAdmin === true || u.IsAdmin === "TRUE" || u.IsAdmin === "true";
}
/**
 * UserName -> ShortName lookup for the client. Falls back to the
 * first 4 letters of DisplayName (itself already falling back to
 * UserName) wherever ShortName is blank.
 */
function getUserShortNames_() {
  const users = sheetToObjects_(SHEET_USERS);
  return shortNamesFromRows_(users, displayNamesFromRows_(users));
}

function shortNamesFromRows_(users, displayNames) {
  const map = {};
  users.forEach(u => {
    const explicit = (u.ShortName && String(u.ShortName).trim()) ? String(u.ShortName).trim() : "";
    if (explicit) {
      map[u.UserName] = explicit;
    } else {
      const nick = displayNames[u.UserName] || u.UserName;
      map[u.UserName] = nick.length <= 4 ? nick : nick.slice(0, 4);
    }
  });
  return map;
}


function getUserIsAdmin(userName) {
  return isUserAdmin_(userName);
}


/**
 * UserName -> DisplayName lookup for the client. Falls back to the
 * UserName itself wherever DisplayName is blank, so this is safe
 * to use everywhere without null-checking on the client.
 */
function getUserDisplayNames_() {
  return displayNamesFromRows_(sheetToObjects_(SHEET_USERS));
}

function displayNamesFromRows_(users) {
  const map = {};
  users.forEach(u => {
    const nick = (u.DisplayName && String(u.DisplayName).trim()) ? String(u.DisplayName).trim() : u.UserName;
    map[u.UserName] = nick;
  });
  return map;
}


// ------------------------------------------------------------
// PROFILE TAB - self-service nickname + notification preferences.
// Every logged-in user can manage their own DisplayName and the 4
// SendXxx toggles here. Phone number is deliberately NOT editable
// from this surface - phone is the login credential, so opening
// that up to self-service would let someone lock out or impersonate
// another user. Phone stays admin-only (see the Admin tab).
// ------------------------------------------------------------

/**
 * Everything the Profile tab needs to render for the logged-in user:
 * their current nickname plus the 4 notification toggles. Booleans
 * default to TRUE when blank via boolFieldDefaultTrue_ (Automation.gs)
 * so a user who predates one of these columns isn't silently opted out.
 */
/**
 * Reminder times are meant to be stored as a plain "HH:mm" string,
 * but an unformatted Sheets cell will silently auto-convert a typed
 * or setValue'd "18:30" into an actual time-of-day serial (a Date
 * object) instead of keeping it as text. That's a problem two ways:
 * (1) minutesSinceMidnight_ can't parse a Date as "HH:mm", so the
 * reminder silently never fires, and (2) google.script.run can fail
 * to serialize a raw Date nested inside a returned object, which is
 * what actually broke the whole Profile tab load. This normalizes
 * either shape back to a clean "HH:mm" string, and updateUserProfile
 * below forces the cell to plain-text format so it can't happen again.
 */
function normalizeReminderTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }
  return value ? String(value) : "";
}

function getUserProfile(userName) {
  // Wrapped defensively: if ANYTHING in here throws (a bad cell value,
  // a missing column, whatever), the client gets a real error message
  // back instead of google.script.run silently mangling the response -
  // which is what a raw Date value nested in this object was doing
  // before (see normalizeReminderTime_ above).
  try {
    const users = sheetToObjects_(SHEET_USERS);
    const u = users.find(u => u.UserName === userName);
    if (!u) return { status: "error", message: "User not found: " + userName };

    return {
      status: "ok",
      userName: String(u.UserName || ""),
      displayName: (u.DisplayName && String(u.DisplayName).trim()) ? String(u.DisplayName).trim() : "",
      sendText: boolFieldDefaultTrue_(u.SendText),
      sendRecapInfo: boolFieldDefaultTrue_(u.SendRecapInfo),
      sendTieInfo: boolFieldDefaultTrue_(u.SendTieInfo),
      sendStillTied: boolFieldDefaultTrue_(u.SendStillTied),
      sendRunbackInfo: boolFieldDefaultTrue_(u.SendRunbackInfo),
      sendParlayInfo: boolFieldDefaultTrue_(u.SendParlayInfo),
      // Pick reminders are opt-IN, so unlike the SendXxx toggles above
      // these default to FALSE/blank when the column is missing or empty -
      // nobody gets a reminder they never asked for. Every value here is
      // forced to a plain string/boolean - never a raw Date/object - since
      // that's exactly what broke this response before.
      reminder1: {
        enabled: !!(u.PickReminder1Enabled === true || u.PickReminder1Enabled === "TRUE" || u.PickReminder1Enabled === "true"),
        day: String(u.PickReminder1Day || ""),
        time: String(normalizeReminderTime_(u.PickReminder1Time) || "")
      },
      reminder2: {
        enabled: !!(u.PickReminder2Enabled === true || u.PickReminder2Enabled === "TRUE" || u.PickReminder2Enabled === "true"),
        day: String(u.PickReminder2Day || ""),
        time: String(normalizeReminderTime_(u.PickReminder2Time) || "")
      },
      reminder3: {
        enabled: !!(u.PickReminder3Enabled === true || u.PickReminder3Enabled === "TRUE" || u.PickReminder3Enabled === "true"),
        day: String(u.PickReminder3Day || ""),
        time: String(normalizeReminderTime_(u.PickReminder3Time) || "")
      }
    };
  } catch (err) {
    return { status: "error", message: "getUserProfile crashed: " + err.message };
  }
}

const PICK_REMINDER_DAYS_ = ["Tuesday", "Wednesday", "Thursday"];

/**
 * Validates one reminder slot from the client. Returns an error
 * string, or null if it's fine to save. An unchecked/disabled slot
 * is always valid regardless of what's in day/time - clearing the
 * checkbox is always a safe way to turn one off.
 */
function validateReminderSlot_(slot) {
  slot = slot || {};
  if (!slot.enabled) return null;
  if (PICK_REMINDER_DAYS_.indexOf(slot.day) === -1) {
    return "Pick a day (Tuesday, Wednesday, or Thursday) for that reminder.";
  }
  if (!/^([01]\d|2[0-3]):(00|15|30|45)$/.test(String(slot.time || ""))) {
    return "Pick a time for that reminder (on the hour, or :15/:30/:45).";
  }
  return null;
}

/**
 * Saves the logged-in user's own nickname + notification prefs,
 * including up to 2 self-managed pick reminders. Self-service only
 * touches DisplayName, the SendXxx columns, and the PickReminderN
 * columns - Phone, Active, SortOrder, IsAdmin etc. are never
 * written here.
 * prefs = { sendText, sendRecapInfo, sendTieInfo, sendStillTied,
 *           sendRunbackInfo,
 *           reminder1: {enabled, day, time}, reminder2: {enabled, day, time} }
 * day is one of "Tuesday"/"Wednesday"/"Thursday", time is "HH:mm" (24hr).
 * Reminders are opt-in - an unchecked slot is simply cleared/ignored,
 * never defaulted on.
 */
function updateUserProfile(userName, displayName, prefs) {
  if (!userName) return { status: "error", message: "Not logged in." };
  prefs = prefs || {};

  const reminder1Error = validateReminderSlot_(prefs.reminder1);
  if (reminder1Error) return { status: "error", message: reminder1Error };
  const reminder2Error = validateReminderSlot_(prefs.reminder2);
  if (reminder2Error) return { status: "error", message: reminder2Error };
  const reminder3Error = validateReminderSlot_(prefs.reminder3);
  if (reminder3Error) return { status: "error", message: reminder3Error };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colUser = headers.indexOf("UserName");
  const colDisplayName = headers.indexOf("DisplayName");
  const colSendText = headers.indexOf("SendText");
  const colSendRecapInfo = headers.indexOf("SendRecapInfo");
  const colSendTieInfo = headers.indexOf("SendTieInfo");
  const colSendStillTied = headers.indexOf("SendStillTied");
  const colSendRunbackInfo = headers.indexOf("SendRunbackInfo");
  const colSendParlayInfo = headers.indexOf("SendParlayInfo");
  const colR1Enabled = headers.indexOf("PickReminder1Enabled");
  const colR1Day = headers.indexOf("PickReminder1Day");
  const colR1Time = headers.indexOf("PickReminder1Time");
  const colR2Enabled = headers.indexOf("PickReminder2Enabled");
  const colR2Day = headers.indexOf("PickReminder2Day");
  const colR2Time = headers.indexOf("PickReminder2Time");
  const colR3Enabled = headers.indexOf("PickReminder3Enabled");
  const colR3Day = headers.indexOf("PickReminder3Day");
  const colR3Time = headers.indexOf("PickReminder3Time");

  if (colUser === -1 || colDisplayName === -1 || colSendText === -1 ||
      colSendRecapInfo === -1 || colSendTieInfo === -1 || colSendStillTied === -1) {
    return { status: "error", message: "Users sheet is missing an expected column." };
  }

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colUser]) === String(userName)) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return { status: "error", message: "User not found." };

  const cleanDisplayName = String(displayName || "").trim();
  sheet.getRange(rowIndex, colDisplayName + 1).setValue(cleanDisplayName);
  sheet.getRange(rowIndex, colSendText + 1).setValue(!!prefs.sendText);
  sheet.getRange(rowIndex, colSendRecapInfo + 1).setValue(!!prefs.sendRecapInfo);
  sheet.getRange(rowIndex, colSendTieInfo + 1).setValue(!!prefs.sendTieInfo);
  sheet.getRange(rowIndex, colSendStillTied + 1).setValue(!!prefs.sendStillTied);
  if (colSendRunbackInfo !== -1) {
    sheet.getRange(rowIndex, colSendRunbackInfo + 1).setValue(!!prefs.sendRunbackInfo);
  }
  if (colSendParlayInfo !== -1) {
    sheet.getRange(rowIndex, colSendParlayInfo + 1).setValue(!!prefs.sendParlayInfo);
  }

  const r1 = prefs.reminder1 || {};
  const r2 = prefs.reminder2 || {};
  const r3 = prefs.reminder3 || {};
  if (colR1Enabled !== -1) sheet.getRange(rowIndex, colR1Enabled + 1).setValue(!!r1.enabled);
  if (colR1Day !== -1) sheet.getRange(rowIndex, colR1Day + 1).setValue(r1.enabled ? r1.day : "");
  // Force plain-text format before writing - otherwise Sheets will
  // silently auto-convert "18:30" into a time-of-day serial (a Date)
  // instead of keeping it as text, which breaks both reading it back
  // here and the reminder-matching logic in Automation.gs.
  if (colR1Time !== -1) {
    sheet.getRange(rowIndex, colR1Time + 1).setNumberFormat("@").setValue(r1.enabled ? r1.time : "");
  }
  if (colR2Enabled !== -1) sheet.getRange(rowIndex, colR2Enabled + 1).setValue(!!r2.enabled);
  if (colR2Day !== -1) sheet.getRange(rowIndex, colR2Day + 1).setValue(r2.enabled ? r2.day : "");
  if (colR2Time !== -1) {
    sheet.getRange(rowIndex, colR2Time + 1).setNumberFormat("@").setValue(r2.enabled ? r2.time : "");
  }
  if (colR3Enabled !== -1) sheet.getRange(rowIndex, colR3Enabled + 1).setValue(!!r3.enabled);
  if (colR3Day !== -1) sheet.getRange(rowIndex, colR3Day + 1).setValue(r3.enabled ? r3.day : "");
  if (colR3Time !== -1) {
    sheet.getRange(rowIndex, colR3Time + 1).setNumberFormat("@").setValue(r3.enabled ? r3.time : "");
  }

  return {
    status: "ok",
    displayName: cleanDisplayName,
    sendText: !!prefs.sendText,
    sendRecapInfo: !!prefs.sendRecapInfo,
    sendTieInfo: !!prefs.sendTieInfo,
    sendStillTied: !!prefs.sendStillTied,
    sendRunbackInfo: !!prefs.sendRunbackInfo,
    reminder1: { enabled: !!r1.enabled, day: r1.enabled ? r1.day : "", time: r1.enabled ? r1.time : "" },
    reminder2: { enabled: !!r2.enabled, day: r2.enabled ? r2.day : "", time: r2.enabled ? r2.time : "" },
    reminder3: { enabled: !!r3.enabled, day: r3.enabled ? r3.day : "", time: r3.enabled ? r3.time : "" }
  };
}


// ------------------------------------------------------------
// ONE-TIME MIGRATION - adds the PickReminder columns (3 slots x
// enabled/day/time). Unlike the SendXxx notification columns, these
// default to FALSE/blank for every existing user - pick reminders
// are opt-in, so nobody starts out subscribed to one they never
// asked for. Safe to re-run; only adds columns that don't exist -
// so running this again after slot 3 was added only adds the 3
// new columns, leaving slots 1 and 2 (and everyone's existing
// settings in them) untouched.
// ------------------------------------------------------------
function addPickReminderColumnsToUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };

  const columnsToAdd = [
    "PickReminder1Enabled", "PickReminder1Day", "PickReminder1Time",
    "PickReminder2Enabled", "PickReminder2Day", "PickReminder2Time",
    "PickReminder3Enabled", "PickReminder3Day", "PickReminder3Time"
  ];
  const added = [];

  columnsToAdd.forEach(colName => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf(colName) !== -1) return; // already exists, skip

    const newCol = headers.length + 1;
    sheet.getRange(1, newCol).setValue(colName).setFontWeight("bold");
    // The two Time columns must stay plain text - otherwise Sheets
    // auto-converts a typed/written "18:30" into a time serial (a
    // Date), which breaks both reading it back and the reminder
    // matching logic. Whole-column format so it holds for every row,
    // not just ones that exist yet.
    if (colName.indexOf("Time") !== -1) {
      sheet.getRange(2, newCol, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("@");
    }
    // Deliberately no default-fill here - blank/FALSE is correct for
    // an opt-in feature, unlike addNotificationPreferenceColumns().
    added.push(colName);
  });

  return { status: "ok", added: added, message: added.length ? ("Added: " + added.join(", ")) : "All columns already existed" };
}


// ------------------------------------------------------------
// ONE-TIME REPAIR - run manually if picks were entered with
// "GAME" (or blank) in the EventID column instead of the real
// ESPN event ID. Matches each pick back to the schedule by
// SeasonYear + Week + Matchup and fixes EventID + PickID in place.
// Safe to re-run; only touches rows that still look broken.
// ------------------------------------------------------------
// ------------------------------------------------------------
// ONE-TIME REPAIR - fixes rows scrambled by an old bug where new
// picks were appended as a fixed-position array that didn't match
// this sheet's actual column order. Symptom: EventID is blank and
// Matchup holds a raw numeric event ID instead of "AWAY @ HOME" text.
// Looks up the real matchup text from the schedule using that number
// as the EventID. Safe to re-run; only touches rows matching that pattern.
// ------------------------------------------------------------
function repairScrambledPicks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PICKS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: "ok", fixed: 0, stillBroken: [] };
  }


  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colEventId = headers.indexOf("EventID");
  const colMatchup = headers.indexOf("Matchup");


  const schedule = sheetToObjects_(SHEET_SCHEDULE).filter(r => r.RowType === "GAME");
  const matchupByEventId = {};
  schedule.forEach(g => matchupByEventId[String(g.EventID)] = g.Matchup);


  let fixed = 0;
  const stillBroken = [];


  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const eventIdVal = row[colEventId];
    const matchupVal = row[colMatchup];
    const eventIdBlank = eventIdVal === "" || eventIdVal === null || eventIdVal === undefined;
    const matchupLooksNumeric = /^\d+$/.test(String(matchupVal).trim());


    if (!eventIdBlank || !matchupLooksNumeric) continue;


    const realEventId = String(matchupVal).trim();
    const realMatchup = matchupByEventId[realEventId];


    sheet.getRange(i + 1, colEventId + 1).setValue(realEventId);


    if (realMatchup) {
      sheet.getRange(i + 1, colMatchup + 1).setValue(realMatchup);
      fixed++;
    } else {
      stillBroken.push("Row " + (i + 1) + ": EventID " + realEventId + " not found in schedule - matchup text left as-is");
    }
  }


  return { status: "ok", fixed: fixed, stillBroken: stillBroken };
}


function backfillPickEventIds() {
  const picksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PICKS);
  if (!picksSheet || picksSheet.getLastRow() < 2) {
    return { status: "ok", fixed: 0, notFound: [] };
  }


  const data = picksSheet.getDataRange().getValues();
  const headers = data[0];
  const colPickId = headers.indexOf("PickID");
  const colYear = headers.indexOf("SeasonYear");
  const colWeek = headers.indexOf("Week");
  const colEventId = headers.indexOf("EventID");
  const colMatchup = headers.indexOf("Matchup");
  const colUser = headers.indexOf("UserName");


  const schedule = sheetToObjects_(SHEET_SCHEDULE).filter(r => r.RowType === "GAME");
  const lookup = {};
  schedule.forEach(g => {
    const key = g.SeasonYear + "|" + g.Week + "|" + String(g.Matchup).trim();
    lookup[key] = g.EventID;
  });


  let fixed = 0;
  const notFound = [];


  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentEventId = row[colEventId];


    // Only touch rows that still have a bogus EventID
    if (currentEventId && currentEventId !== "GAME" && currentEventId !== "") continue;


    const key = row[colYear] + "|" + row[colWeek] + "|" + String(row[colMatchup]).trim();
    const realEventId = lookup[key];


    if (realEventId) {
      picksSheet.getRange(i + 1, colEventId + 1).setValue(realEventId);
      if (colPickId !== -1) {
        picksSheet.getRange(i + 1, colPickId + 1).setValue(realEventId + "-" + row[colUser]);
      }
      fixed++;
    } else {
      notFound.push("Row " + (i + 1) + ": " + row[colMatchup] + " (Week " + row[colWeek] + ", " + row[colYear] + ")");
    }
  }


  return { status: "ok", fixed: fixed, notFound: notFound };
}


// ------------------------------------------------------------
// ONE-TIME CLEANUP - removes any old rows in Picks where
// UserName is "Parlay". Parlay is now computed automatically
// in the ParlayPicks sheet, not entered as a manual pick.
// Safe to re-run; does nothing once no Parlay rows remain.
// ------------------------------------------------------------
function cleanupManualParlayPicks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PICKS);
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: "ok", removed: 0 };
  }


  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colUser = headers.indexOf("UserName");


  // Walk bottom-up so row deletion doesn't shift indices we still need to check
  let removed = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][colUser] === PARLAY_USER) {
      sheet.deleteRow(i + 1);
      removed++;
    }
  }


  return { status: "ok", removed: removed };
}


// ------------------------------------------------------------
// PARLAY - computes the majority pick among real users for
// every game in a given week and writes it to ParlayPicks.
// A tie at the top vote count means no consensus (PickedTeam
// left blank). Called automatically by getWeekData, so it stays
// in sync any time the matchups screen is viewed - no manual run needed.
// ------------------------------------------------------------
function computeParlayPicksForWeek(year, week) {
  ensureSupportSheets();


  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r =>
      Number(r.SeasonYear) === Number(year) &&
      String(r.Week) === String(week) &&
      (r.RowType === "GAME" || r.RowType === "PRESEASON")
    );


  if (!games.length) return { status: "ok", gamesProcessed: 0 };


  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PARLAY);


  // FAST PATH (the common case): work out what ParlayPicks SHOULD say,
  // compare it to what it says now, and stop right there if nothing
  // changed - no lock, no writes. This runs every time anyone opens the
  // Games or Parlay screen, and used to rewrite every game's row (one
  // sheet write per game, under a script-wide lock) even when no pick
  // had changed, which was the biggest source of slow loads.
  let diff = diffParlayRows_(sheet, buildParlayRowsForWeek_(year, week, games));
  if (!diff.toUpdate.length && !diff.toAppend.length) {
    return { status: "ok", gamesProcessed: games.length, rowsWritten: 0 };
  }


  // SLOW PATH: something actually changed. Take the lock, then recompute
  // from scratch INSIDE it so the result reflects the latest picks even if
  // another submission landed in between, and so two near-simultaneous
  // submissions can never both conclude "no row exists yet" and both append
  // duplicate rows for the same game.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // wait up to 30s for another execution to finish


  try {
    diff = diffParlayRows_(sheet, buildParlayRowsForWeek_(year, week, games));
    const now = new Date();
    const GENERATED_AT_COL = 7; // 0-based index of GeneratedAt in the row array


    if (diff.toUpdate.length || diff.toAppend.length) {
      const width = (diff.toUpdate[0] || { row: diff.toAppend[0] }).row.length;
      if (width > sheet.getMaxColumns()) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
      }
    }


    diff.toUpdate.forEach(u => {
      u.row[GENERATED_AT_COL] = now;
      sheet.getRange(u.sheetRow, 1, 1, u.row.length).setValues([u.row]);
    });


    if (diff.toAppend.length) {
      diff.toAppend.forEach(r => { r[GENERATED_AT_COL] = now; });
      const startRow = sheet.getLastRow() + 1;
      const endRow = startRow + diff.toAppend.length - 1;
      if (endRow > sheet.getMaxRows()) {
        sheet.insertRowsAfter(sheet.getMaxRows(), endRow - sheet.getMaxRows());
      }
      // One batched write for all new rows instead of one appendRow each
      sheet.getRange(startRow, 1, diff.toAppend.length, diff.toAppend[0].length).setValues(diff.toAppend);
    }


    return {
      status: "ok",
      gamesProcessed: games.length,
      rowsWritten: diff.toUpdate.length + diff.toAppend.length
    };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Pure computation half of computeParlayPicksForWeek: reads picks, tie-break
 * votes and leader overrides, and returns one row array per game in the exact
 * column order ParlayPicks is written in (SeasonYear, Week, EventID, Matchup,
 * PickedTeam, VoteCount, TotalVotes, GeneratedAt, ResolvedVia). GeneratedAt is
 * left blank here and stamped only when a row is actually written. Writes
 * nothing.
 */
function buildParlayRowsForWeek_(year, week, games) {
  const activeUsers = getUsers();


  const picks = sheetToObjects_(SHEET_PICKS)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      String(p.Week) === String(week) &&
      p.UserName !== PARLAY_USER &&
      p.PickedTeam &&
      p.PickedTeam !== "SKIP"
    );


  const votesByEvent = {};
  picks.forEach(p => {
    if (!votesByEvent[p.EventID]) votesByEvent[p.EventID] = {};
    votesByEvent[p.EventID][p.PickedTeam] = (votesByEvent[p.EventID][p.PickedTeam] || 0) + 1;
  });


  const tieBreakVotes = sheetToObjects_(SHEET_TIEBREAK)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      String(p.Week) === String(week) &&
      p.PickedTeam
    );


  const tieBreakVotesByEvent = {};
  const tieBreakVotersByEvent = {};
  tieBreakVotes.forEach(p => {
    if (!tieBreakVotesByEvent[p.EventID]) tieBreakVotesByEvent[p.EventID] = {};
    tieBreakVotesByEvent[p.EventID][p.PickedTeam] = (tieBreakVotesByEvent[p.EventID][p.PickedTeam] || 0) + 1;


    if (!tieBreakVotersByEvent[p.EventID]) tieBreakVotersByEvent[p.EventID] = new Set();
    tieBreakVotersByEvent[p.EventID].add(p.UserName);
  });


  const leaderOverrides = sheetToObjects_(SHEET_LEADER_OVERRIDE)
    .filter(p => Number(p.SeasonYear) === Number(year) && String(p.Week) === String(week) && p.PickedTeam);
  const leaderOverrideByEvent = {};
  leaderOverrides.forEach(p => leaderOverrideByEvent[p.EventID] = p.PickedTeam);


  const resolveVoteMap_ = voteMap => {
    const teams = Object.keys(voteMap);
    let top = "";
    let count = 0;
    teams.forEach(t => {
      if (voteMap[t] > count) {
        count = voteMap[t];
        top = t;
      }
    });
    const topTeams = teams.filter(t => voteMap[t] === count);
    return { pickedTeam: topTeams.length > 1 ? "" : top, voteCount: count, tied: topTeams.length > 1 };
  };


  return games.map(g => {
    const voteMap = votesByEvent[g.EventID] || {};
    const teams = Object.keys(voteMap);


    let pickedTeam = "";
    let voteCount = 0;
    let resolvedVia = "";
    const totalVotes = teams.reduce((sum, t) => sum + voteMap[t], 0);


    if (teams.length) {
      const main = resolveVoteMap_(voteMap);
      pickedTeam = main.pickedTeam;
      voteCount = main.voteCount;


      if (main.tied && leaderOverrideByEvent[g.EventID]) {
        pickedTeam = leaderOverrideByEvent[g.EventID];
        resolvedVia = "leader";


      } else if (main.tied) {
        const votersSet = tieBreakVotersByEvent[g.EventID] || new Set();
        const everyoneVoted = activeUsers.length > 0 && activeUsers.every(u => votersSet.has(u));


        if (everyoneVoted) {
          const tieBreak = resolveVoteMap_(tieBreakVotesByEvent[g.EventID]);
          if (tieBreak.pickedTeam) {
            pickedTeam = tieBreak.pickedTeam;
            resolvedVia = "tiebreak";
          }
        }
      }
    }


    return [year, week, g.EventID, g.Matchup, pickedTeam, voteCount, totalVotes, "", resolvedVia];
  });
}


/**
 * Compares the rows computeParlayPicksForWeek WANTS in ParlayPicks against
 * what's there now. Returns { toUpdate: [{sheetRow, row}], toAppend: [row] },
 * both empty when nothing differs. GeneratedAt (a timestamp that always
 * differs) is deliberately excluded from the comparison. Compares by column
 * position because the writer is positional too.
 */
function diffParlayRows_(sheet, desiredRows) {
  const existing = sheet.getDataRange().getValues();
  let colEventId = existing.length ? existing[0].indexOf("EventID") : -1;
  if (colEventId === -1) colEventId = 2;


  const sheetRowByEvent = {};
  for (let i = 1; i < existing.length; i++) {
    sheetRowByEvent[existing[i][colEventId]] = i + 1;
  }


  const norm = v => (v === undefined || v === null) ? "" : String(v);
  const COMPARE_COLS = [0, 1, 2, 3, 4, 5, 6, 8]; // everything except GeneratedAt (7)


  const toUpdate = [];
  const toAppend = [];
  desiredRows.forEach(row => {
    const sheetRow = sheetRowByEvent[row[2]];
    if (!sheetRow) {
      toAppend.push(row);
      return;
    }
    const cur = existing[sheetRow - 1];
    const unchanged = COMPARE_COLS.every(c => norm(cur[c]) === norm(row[c]));
    if (!unchanged) toUpdate.push({ sheetRow: sheetRow, row: row });
  });


  return { toUpdate: toUpdate, toAppend: toAppend };
}


// ------------------------------------------------------------
// ONE-TIME REPAIR - wipes the ParlayPicks sheet and rewrites
// the correct header row. Use this if the sheet was created
// manually (e.g. copied from Picks) with the wrong columns,
// since computeParlayPicksForWeek writes by column position.
// ------------------------------------------------------------
function resetParlayPicksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PARLAY);


  if (sheet) {
    ss.deleteSheet(sheet);
  }


  sheet = ss.insertSheet(SHEET_PARLAY);
  sheet.getRange(1, 1, 1, 8)
    .setValues([[
      "SeasonYear", "Week", "EventID", "Matchup",
      "PickedTeam", "VoteCount", "TotalVotes", "GeneratedAt"
    ]])
    .setFontWeight("bold");
  applyStandardFont_(sheet);


  return { status: "ok", message: "ParlayPicks sheet reset with correct headers" };
}


// ------------------------------------------------------------
// BULK GENERATE - recomputes Parlay's majority pick for every
// week of every season currently in API_NFL_Schedule, using
// whatever picks already exist in the Picks sheet. No arguments,
// so it can be run directly from the Apps Script dropdown.
// ------------------------------------------------------------
function generateAllParlayPicksAllYears() {
  const years = getYears();
  let totalWeeks = 0;


  years.forEach(year => {
    const weeks = getWeeks(year);
    weeks.forEach(week => {
      computeParlayPicksForWeek(year, week);
      totalWeeks++;
    });
  });


  return { status: "ok", years: years, weeksProcessed: totalWeeks };
}


/**
 * The "current" season is simply the most recent year present in the
 * schedule. (Previously tried to detect in-progress seasons via a
 * Status field, but that column no longer exists in this schema -
 * the most-recent-year rule is simpler and correct for the normal
 * case of one active season at a time.)
 */
function getCurrentSeasonYearImpl_() {
  const rows = sheetToObjects_(SHEET_SCHEDULE).filter(r => r.RowType === "GAME");
  if (!rows.length) return null;
  const years = rows.map(r => Number(r.SeasonYear));
  return Math.max(...years);
}


/**
 * Default week to land on for a season: the earliest week that still
 * has an unplayed/incomplete game, or the last week if the season is over.
 */
/**
 * Default week to land on: the week whose Tue-Mon span contains today.
 * Before the season starts, that's Week 1 (its span hasn't ended yet).
 * After the season ends, falls back to the last week on file.
 */
/**
 * Every week in a season tagged with its pick-window status:
 * "closed" (past, read-only), "open" (picks can be submitted now),
 * or "future" (not open yet). Used to color the Week dropdown so
 * past weeks are visually de-emphasized.
 */
function getWeeksWithStatusImpl_(year) {
  const allGames = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => Number(r.SeasonYear) === Number(year) && (r.RowType === "GAME" || r.RowType === "PRESEASON"));


  const weeks = [...new Set(allGames.map(r => String(r.Week)))].sort(compareWeeks_);


  return weeks.map(w => {
    const weekGames = allGames.filter(r => String(r.Week) === w);
    const win = computePickWindow_(weekGames);
    let status = "open";
    if (win) {
      if (win.isClosed) status = "closed";
      else if (win.isFuture) status = "future";
    }
    return { week: w, status: status };
  });
}


// ------------------------------------------------------------
// TESTING ONLY - lets you simulate "today" being a different date
// so you can test pick windows without waiting for the real season.
// When no override is set, everything uses the real current date as normal.
// ------------------------------------------------------------
function getEffectiveNow_() {
  const override = PropertiesService.getScriptProperties().getProperty("TEST_OVERRIDE_DATE");
  if (override) {
    const d = new Date(override + "T12:00:00");
    if (!isNaN(d.getTime())) return d;
  }
  return new Date();
}


function setTestDate(dateString) {
  PropertiesService.getScriptProperties().setProperty("TEST_OVERRIDE_DATE", dateString);
  bumpDataVersion_();
  return { status: "ok", testDate: dateString };
}


function clearTestDate() {
  PropertiesService.getScriptProperties().deleteProperty("TEST_OVERRIDE_DATE");
  bumpDataVersion_();
  return { status: "ok" };
}


// EDIT THE DATE BELOW, then run this function (from the dropdown) to test
// as if today were that date. Run clearTestDate() when you're done testing.
function DEBUG_setTestDate() {
  return setTestDate("2026-09-09"); // <-- change this date as needed
}


/**
 * Default week to land on: the week whose Tue-Mon span contains today.
 * Before the season starts, that's Week 1 (its span hasn't ended yet).
 * After the season ends, falls back to the last week on file.
 */
function getDefaultWeekImpl_(year) {
  const allGames = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => Number(r.SeasonYear) === Number(year) && (r.RowType === "GAME" || r.RowType === "PRESEASON"));
  if (!allGames.length) return null;


  const weeks = [...new Set(allGames.map(r => String(r.Week)))].sort(compareWeeks_);
  const tz = Session.getScriptTimeZone();
  const todayIso = Utilities.formatDate(getEffectiveNow_(), tz, "yyyy-MM-dd");


  for (const w of weeks) {
    const weekGames = allGames.filter(r => String(r.Week) === w);
    const win = computePickWindow_(weekGames);
    if (!win) continue;


    const endDate = new Date(win.nominalTuesdayIso + "T00:00:00");
    endDate.setDate(endDate.getDate() + 6); // Tuesday + 6 days = the following Monday
    const endIso = Utilities.formatDate(endDate, tz, "yyyy-MM-dd");


    if (todayIso <= endIso) {
      return w; // today is within this week's span, or this is the next upcoming week
    }
  }


  return weeks[weeks.length - 1]; // season's over - land on the last week
}


// ------------------------------------------------------------
// PICK WINDOW EXTENSION - a one-off deadline push for a specific
// week, without changing the normal Tue-Wed rule for any other
// week. Stored as a Script Property, checked by computePickWindow_.
// ------------------------------------------------------------


/**
 * Manually extends a specific week's pick deadline to a given date
 * (yyyy-MM-dd). Only takes effect if it's later than the week's
 * normal Wednesday close - can't shorten a window with this.
 */
function extendPickWindow(year, week, newClosingDateString) {
  const key = "PICK_WINDOW_EXTEND_" + year + "_" + week;
  PropertiesService.getScriptProperties().setProperty(key, newClosingDateString);
  bumpDataVersion_();
  return { status: "ok", year: year, week: week, newClosingDate: newClosingDateString };
}


/**
 * Removes a manual extension for a specific week, reverting it to
 * the normal Tue-Wed rule.
 */
function clearPickWindowExtension(year, week) {
  PropertiesService.getScriptProperties().deleteProperty("PICK_WINDOW_EXTEND_" + year + "_" + week);
  bumpDataVersion_();
  return { status: "ok" };
}


/**
 * Convenience one-click version: extends whichever week is
 * currently "live" through tomorrow, without needing to know the
 * exact year/week. Run this directly from the function dropdown.
 */
function extendCurrentWeekPickWindowThroughTomorrow() {
  const year = getCurrentSeasonYear();
  if (!year) return { status: "error", message: "No current season found" };


  const week = getDefaultWeek(year);
  if (!week) return { status: "error", message: "No current week found" };


  const tz = Session.getScriptTimeZone();
  const tomorrow = new Date(getEffectiveNow_());
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowIso = Utilities.formatDate(tomorrow, tz, "yyyy-MM-dd");


  const result = extendPickWindow(year, week, tomorrowIso);
  Logger.log("Extended pick window for " + year + " Week " + week + " through " + tomorrowIso);
  return result;
}


/**
 * Extends whichever week is currently "live" through the coming
 * Thursday (or today, if today already is Thursday). Same
 * auto-detection as the "through tomorrow" version.
 */
function extendCurrentWeekPickWindowToThursday() {
  const year = getCurrentSeasonYear();
  if (!year) return { status: "error", message: "No current season found" };


  const week = getDefaultWeek(year);
  if (!week) return { status: "error", message: "No current week found" };


  const tz = Session.getScriptTimeZone();
  const today = getEffectiveNow_();
  const thursday = new Date(today);
  const dayOfWeek = thursday.getDay(); // 0=Sun..6=Sat, Thursday=4
  const daysUntilThursday = (4 - dayOfWeek + 7) % 7; // 0 if today IS Thursday
  thursday.setDate(thursday.getDate() + daysUntilThursday);
  const thursdayIso = Utilities.formatDate(thursday, tz, "yyyy-MM-dd");


  const result = extendPickWindow(year, week, thursdayIso);
  Logger.log("Extended pick window for " + year + " Week " + week + " through " + thursdayIso + " (Thursday)");
  return result;
}


// ------------------------------------------------------------
// EARLY OPEN - the mirror image of extendPickWindow: lets a
// specific week's picks open EARLIER than its normal Tuesday,
// without changing the rule for any other week.
// ------------------------------------------------------------


/**
 * Manually opens a specific week's picks starting on a given date
 * (yyyy-MM-dd). Only takes effect if it's EARLIER than the week's
 * normal Tuesday open - can't use this to delay a window, only to
 * open it early (use extendPickWindow for delaying the close).
 */
function openPickWindowEarly(year, week, newOpeningDateString) {
  const key = "PICK_WINDOW_OPEN_EARLY_" + year + "_" + week;
  PropertiesService.getScriptProperties().setProperty(key, newOpeningDateString);
  bumpDataVersion_();
  return { status: "ok", year: year, week: week, newOpeningDate: newOpeningDateString };
}


/**
 * Removes a manual early-open override for a specific week,
 * reverting it to the normal Tuesday-open rule.
 */
function clearPickWindowEarlyOpen(year, week) {
  PropertiesService.getScriptProperties().deleteProperty("PICK_WINDOW_OPEN_EARLY_" + year + "_" + week);
  bumpDataVersion_();
  return { status: "ok" };
}


/**
 * Finds the week that comes immediately after whichever week
 * getDefaultWeek currently considers "current" - used so we don't
 * have to guess at calendar math to figure out "next week."
 */
function getNextWeek_(year) {
  const weeks = getWeeks(year); // already sorted preseason-first, then numeric
  const currentWeek = getDefaultWeek(year);
  const idx = weeks.findIndex(w => String(w) === String(currentWeek));
  if (idx === -1 || idx === weeks.length - 1) return null;
  return weeks[idx + 1];
}
function getPreviousWeek_(year) {
  const weeks = getWeeks(year); // already sorted preseason-first, then numeric
  const currentWeek = getDefaultWeek(year);
  const idx = weeks.findIndex(w => String(w) === String(currentWeek));
  if (idx <= 0) return null;
  return weeks[idx - 1];
}
/**
 * Convenience one-click version: opens whichever week comes next
 * (after the currently-active one) starting today, without needing
 * to know the exact year/week. Run this directly from the function
 * dropdown, or the Sheets menu.
 */
function openNextWeekPickWindowToday() {
  const year = getCurrentSeasonYear();
  if (!year) return { status: "error", message: "No current season found" };


  const nextWeek = getNextWeek_(year);
  if (!nextWeek) return { status: "error", message: "No next week found after the current one" };


  const tz = Session.getScriptTimeZone();
  const todayIso = Utilities.formatDate(getEffectiveNow_(), tz, "yyyy-MM-dd");


  const result = openPickWindowEarly(year, nextWeek, todayIso);
  Logger.log("Opened pick window early for " + year + " Week " + nextWeek + " starting " + todayIso);
  return result;
}


/**
 * DIAGNOSTIC - shows exactly what's currently detected as the
 * "current" year/week, what extension (if any) is stored for it,
 * and what computePickWindow_ actually resolves for that week right
 * now. Use this if an extension doesn't seem to be taking effect.
 */
function checkPickWindowExtension() {
  const year = getCurrentSeasonYear();
  const week = getDefaultWeek(year);


  Logger.log("Detected current year/week: " + year + " / " + week);


  const props = PropertiesService.getScriptProperties();


  const extKey = "PICK_WINDOW_EXTEND_" + year + "_" + week;
  const extValue = props.getProperty(extKey);
  Logger.log("Extension property key: " + extKey);
  Logger.log("Extension property value: " + (extValue || "(not set)"));


  const allGames = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week) && (r.RowType === "GAME" || r.RowType === "PRESEASON"));
  Logger.log("Games found for this week: " + allGames.length);


  const win = computePickWindow_(allGames);
  Logger.log("computePickWindow_ result: " + JSON.stringify(win));


  const nextWeek = getNextWeek_(year);
  let nextWeekInfo = null;
  if (nextWeek) {
    const openKey = "PICK_WINDOW_OPEN_EARLY_" + year + "_" + nextWeek;
    const openValue = props.getProperty(openKey);
    Logger.log("Next week: " + nextWeek + " - early-open value: " + (openValue || "(not set)"));


    const nextWeekGames = sheetToObjects_(SHEET_SCHEDULE)
      .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(nextWeek) && (r.RowType === "GAME" || r.RowType === "PRESEASON"));
    const nextWin = computePickWindow_(nextWeekGames);
    Logger.log("Next week computePickWindow_ result: " + JSON.stringify(nextWin));


    nextWeekInfo = { week: nextWeek, earlyOpenKey: openKey, earlyOpenValue: openValue, pickWindow: nextWin };
  } else {
    Logger.log("No next week found after " + week);
  }


  return { year: year, week: week, extensionKey: extKey, extensionValue: extValue, pickWindow: win, nextWeek: nextWeekInfo };
}


/**
 * Like checkPickWindowExtension() but for any year/week you specify,
 * not just the current one. Used by the Admin tab.
 */
function checkPickWindowForWeek(year, week) {
  const props = PropertiesService.getScriptProperties();
  const extValue = props.getProperty("PICK_WINDOW_EXTEND_" + year + "_" + week);
  const earlyOpenValue = props.getProperty("PICK_WINDOW_OPEN_EARLY_" + year + "_" + week);


  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week) && (r.RowType === "GAME" || r.RowType === "PRESEASON"));


  const win = computePickWindow_(games);


  return {
    year: year,
    week: week,
    extensionValue: extValue || null,
    earlyOpenValue: earlyOpenValue || null,
    pickWindow: win
  };
}


function adminCheckPickWindow(callerUserName, year, week) {
  if (!isUserAdmin_(callerUserName)) return { status: "error", message: "Not authorized." };
  return checkPickWindowForWeek(year, week);
}


function adminExtendPickWindow(callerUserName, year, week, newClosingDateString) {
  if (!isUserAdmin_(callerUserName)) return { status: "error", message: "Not authorized." };
  return extendPickWindow(year, week, newClosingDateString);
}


function adminClearPickWindowExtension(callerUserName, year, week) {
  if (!isUserAdmin_(callerUserName)) return { status: "error", message: "Not authorized." };
  return clearPickWindowExtension(year, week);
}


function adminOpenPickWindowEarly(callerUserName, year, week, newOpeningDateString) {
  if (!isUserAdmin_(callerUserName)) return { status: "error", message: "Not authorized." };
  return openPickWindowEarly(year, week, newOpeningDateString);
}


function adminClearPickWindowEarlyOpen(callerUserName, year, week) {
  if (!isUserAdmin_(callerUserName)) return { status: "error", message: "Not authorized." };
  return clearPickWindowEarlyOpen(year, week);
}


/**
 * Past seasons only - anything that isn't the current season year.
 * Used to populate the History tab's year picker.
 */
function getHistoryYears() {
  const currentYear = getCurrentSeasonYear();
  return getYears().filter(y => y !== currentYear);
}


/**
 * Single call the client makes right after login: everything needed
 * to land on the right season/week without extra round trips.
 */
function getAppBootstrap() {
  const currentYear = getCurrentSeasonYear();
  const defaultWeek = currentYear ? getDefaultWeek(currentYear) : null;
  // One read of the Users sheet feeds all three user lookups (it used to be four).
  const userRows = sheetToObjects_(SHEET_USERS);
  const displayNames = displayNamesFromRows_(userRows);
  return {
    currentYear: currentYear,
    defaultWeek: defaultWeek,
    weeks: currentYear ? getWeeksWithStatus(currentYear) : [],
    users: usersFromRows_(userRows),
    allYears: getYears(),
    displayNames: displayNames,
    shortNames: shortNamesFromRows_(userRows, displayNames)
  };
}


// ------------------------------------------------------------
// LOGIN BY PHONE
// ------------------------------------------------------------
function normalizePhone_(phone) {
  let digits = String(phone || "").replace(/\D/g, "");
  // Strip a leading US country code "1" if present (11 digits starting
  // with 1), so nobody has to type it and it doesn't matter whether the
  // stored number in Users includes it or not - both forms match.
  if (digits.length === 11 && digits.charAt(0) === "1") {
    digits = digits.substring(1);
  }
  return digits;
}


function loginByPhone(phone) {
  const users = sheetToObjects_(SHEET_USERS);
  const target = normalizePhone_(phone);
  if (!target) return null;


  const match = users.find(u => normalizePhone_(u.Phone) === target && target.length > 0);
  return match ? match.UserName : null;
}


/**
 * Active user names + display names, for the "who's picking?" step
 * of login - shown BEFORE any authentication, so it must never
 * include phone numbers or anything else sensitive.
 */
function getLoginUserList() {
  const displayNames = getUserDisplayNames_();


  return sheetToObjects_(SHEET_USERS)
    .filter(u =>
      (u.Active === true || u.Active === "TRUE" || u.Active === "true") &&
      u.UserName !== PARLAY_USER
    )
    .sort((a, b) => (a.SortOrder || 0) - (b.SortOrder || 0))
    .map(u => ({
      userName: u.UserName,
      displayName: displayNames[u.UserName] || u.UserName
    }));
}


/**
 * Second step of login: given a name already picked from the list,
 * verify the phone number matches what's on file for THAT specific
 * user (rather than searching all users by phone, like loginByPhone
 * does). Returns {status:"ok", userName} or {status:"error", message}.
 */
function verifyUserLogin(userName, phone) {
  const target = normalizePhone_(phone);
  if (!target) {
    return { status: "error", message: "Please enter a phone number." };
  }


  const users = sheetToObjects_(SHEET_USERS);
  const user = users.find(u => u.UserName === userName);


  if (!user) {
    return { status: "error", message: "User not found." };
  }


  const isActive = (user.Active === true || user.Active === "TRUE" || user.Active === "true");
  if (!isActive) {
    return { status: "error", message: "This user is not active." };
  }


  if (normalizePhone_(user.Phone) !== target) {
    return { status: "error", message: "That phone number doesn't match our records for this name." };
  }


    return { status: "ok", userName: userName, isAdmin: isUserAdmin_(userName) };


}


// ------------------------------------------------------------
// READ HELPERS
// ------------------------------------------------------------
/**
 * Compares two Week values for sorting. Handles both plain regular-
 * season numbers ("1".."18") and preseason ("P1".."P4") - preseason
 * always sorts before regular season, and within each group sorts
 * numerically rather than as text (so "P2" < "P10", "9" < "10").
 */
function compareWeeks_(a, b) {
  const aStr = String(a);
  const bStr = String(b);
  const aIsPre = aStr.charAt(0) === "P";
  const bIsPre = bStr.charAt(0) === "P";


  if (aIsPre && !bIsPre) return -1;
  if (!aIsPre && bIsPre) return 1;


  const aNum = Number(aIsPre ? aStr.substring(1) : aStr);
  const bNum = Number(bIsPre ? bStr.substring(1) : bStr);
  return aNum - bNum;
}


function sheetToObjects_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}


// ------------------------------------------------------------
// PUBLIC API - called from the client (Index.html)
// ------------------------------------------------------------


function getYearsImpl_() {
  const rows = sheetToObjects_(SHEET_SCHEDULE);
  const years = [...new Set(rows.map(r => Number(r.SeasonYear)))]
    .filter(Boolean)
    .sort((a, b) => b - a);
  return years;
}


function getWeeks(year) {
  const rows = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => Number(r.SeasonYear) === Number(year) && (r.RowType === "GAME" || r.RowType === "PRESEASON"));
  const weeks = [...new Set(rows.map(r => String(r.Week)))]
    .sort(compareWeeks_);
  return weeks;
}


function getUsers() {
  return usersFromRows_(sheetToObjects_(SHEET_USERS));
}

function usersFromRows_(rows) {
  return rows
    .filter(u =>
      (u.Active === true || u.Active === "TRUE" || u.Active === "true") &&
      u.UserName !== PARLAY_USER
    )
    .sort((a, b) => (a.SortOrder || 0) - (b.SortOrder || 0))
    .map(u => u.UserName);
}


/**
 * Pick window for a week: the Tuesday and Wednesday that lead into
 * that week's games (NFL scheduling weeks run Tue-Mon). Computed from
 * the earliest game date in the week. Also reports whether the window
 * is open right now, hasn't opened yet, or has already closed, based
 * on today's date in the script's timezone.
 */
function computePickWindow_(games) {
  const dates = games
    .map(g => g.EventDate ? new Date(g.EventDate) : null)
    .filter(d => d && !isNaN(d.getTime()));


  if (!dates.length) return null;


  const earliest = new Date(Math.min(...dates.map(d => d.getTime())));
  const day = earliest.getDay(); // 0=Sun..6=Sat
  const diffToTuesday = (day - 2 + 7) % 7;


  const tuesday = new Date(earliest);
  tuesday.setDate(tuesday.getDate() - diffToTuesday);


  const wednesday = new Date(tuesday);
  wednesday.setDate(wednesday.getDate() + 2);


  const tz = Session.getScriptTimeZone();
  const fmt = d => Utilities.formatDate(d, tz, "EEE, MMM d");
  const iso = d => Utilities.formatDate(d, tz, "yyyy-MM-dd");


  let effectiveTuesday = tuesday;
  let effectiveWednesday = wednesday;


  // Check for manual one-off overrides for this specific week - temporary
  // adjustments that don't change the Tue-Wed rule for any other week,
  // past or future.
  if (games.length && games[0].SeasonYear && games[0].Week) {
    const props = PropertiesService.getScriptProperties();


    const extKey = "PICK_WINDOW_EXTEND_" + games[0].SeasonYear + "_" + games[0].Week;
    const extOverride = props.getProperty(extKey);
    if (extOverride) {
      const overrideDate = new Date(extOverride + "T00:00:00");
      if (!isNaN(overrideDate.getTime()) && overrideDate > wednesday) {
        effectiveWednesday = overrideDate;
      }
    }


    const openKey = "PICK_WINDOW_OPEN_EARLY_" + games[0].SeasonYear + "_" + games[0].Week;
    const openOverride = props.getProperty(openKey);
    if (openOverride) {
      const overrideDate = new Date(openOverride + "T00:00:00");
      if (!isNaN(overrideDate.getTime()) && overrideDate < tuesday) {
        effectiveTuesday = overrideDate;
      }
    }
  }


  const todayIso = iso(getEffectiveNow_());
  const tuesdayIso = iso(effectiveTuesday);
  const wednesdayIso = iso(effectiveWednesday);


  const isFuture = todayIso < tuesdayIso;
  const isOpen = todayIso >= tuesdayIso && todayIso <= wednesdayIso;
  const isClosed = todayIso > wednesdayIso;
  const nominalTuesdayIso = iso(tuesday);
  return {
    tuesday: fmt(effectiveTuesday),
    wednesday: fmt(effectiveWednesday),
    tuesdayIso: tuesdayIso,
    wednesdayIso: wednesdayIso,
    nominalTuesdayIso: nominalTuesdayIso,
    isFuture: isFuture,
    isOpen: isOpen,
    isClosed: isClosed
  };
}


/**
 * Returns the most recent odds snapshot (by PulledAt) per EventID for a
 * given season/week, read from NFL_OddsHistory. Spread/O-U now live there
 * instead of on the schedule sheet itself.
 */
/**
 * All odds-history rows for one year/week. API_NFL_OddsHistory is the
 * biggest sheet in the project (thousands of rows, growing with every pull),
 * and the Games and Parlay screens each used to read it in full three
 * separate times per load. Call this once per request and pass the result
 * to getLatestOddsByEvent_ / getClosingOddsByEvent_ (their optional
 * weekRows argument) and any other per-week odds filtering.
 */
function getOddsRowsForWeek_(year, week) {
  return sheetToObjects_(SHEET_ODDS)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));
}


function getLatestOddsByEvent_(year, week, weekRows) {
  const rows = weekRows || getOddsRowsForWeek_(year, week);


  const latestByEvent = {};
  rows.forEach(r => {
    const eventId = r.EventID;
    const pulledAtMs = r.PulledAt ? new Date(r.PulledAt).getTime() : 0;
    const existing = latestByEvent[eventId];
    if (!existing || pulledAtMs >= existing._pulledAtMs) {
      latestByEvent[eventId] = {
        spread: r.Spread || "",
        overUnder: (r.OverUnder === "" || r.OverUnder === null || r.OverUnder === undefined) ? "" : r.OverUnder,
        _pulledAtMs: pulledAtMs
      };
    }
  });


  return latestByEvent;
}




function closingMoneylineForTeam_(closing, team, awayTeam, homeTeam) {
  if (!closing || !team) return null;
  if (closing.awayMoneyline && closing.homeMoneyline) {
    const isAway = (team === awayTeam);
    const ml = isAway ? Number(closing.awayMoneyline) : Number(closing.homeMoneyline);
    if (!isNaN(ml)) return { ml: ml, source: "real" };
  }
  if (closing.spread) {
    const parsed = parseSpreadString_(closing.spread);
    if (parsed && parsed.points > 0) {
      const spreadResult = spreadToMoneyline_(parsed.points);
      const isFavorite = (team.toUpperCase() === parsed.favTeam);
      return { ml: isFavorite ? spreadResult.favoriteML : spreadResult.underdogML, source: "estimated" };
    }
  }
  return null;
}


function getClosingOddsByEvent_(year, week, weekRows) {
  const rows = (weekRows || getOddsRowsForWeek_(year, week))
    .filter(r => r.IsClosingSnapshot === true || r.IsClosingSnapshot === "TRUE");


  // If more than one row for the same EventID is flagged IsClosingSnapshot=TRUE
  // (a known data issue — should be exactly one), take the one pulled most
  // recently rather than whichever happens to come last in the sheet.
  const byEvent = {};
  rows.forEach(r => {
    const pulledAtMs = r.PulledAt ? new Date(r.PulledAt).getTime() : 0;
    const existing = byEvent[r.EventID];
    if (!existing || pulledAtMs >= existing._pulledAtMs) {
      byEvent[r.EventID] = {
        spread: r.Spread || "",
        awayMoneyline: r.AwayMoneyline,
        homeMoneyline: r.HomeMoneyline,
        _pulledAtMs: pulledAtMs
      };
    }
  });


  return byEvent;
}


function parseSpreadMagnitude_(spreadText) {
  if (!spreadText) return null;
  const trimmed = String(spreadText).trim();
  if (/^(EVEN|PK)$/i.test(trimmed)) return { favTeam: null, points: 0 };
  const match = trimmed.match(/^([A-Za-z]+)\s*(-?\d+(\.\d+)?)$/);
  if (!match) return null;
  return { favTeam: match[1].toUpperCase(), points: Math.abs(parseFloat(match[2])) };
}






/**
 * Returns everything the UI needs to render one week:
 * games, each user's pick for each game, and final scores if available.
 */
function buildWeekDataPayload_(year, week) {
  ensureSupportSheets();
  computeParlayPicksForWeek(year, week);


  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r =>
      Number(r.SeasonYear) === Number(year) &&
      String(r.Week) === String(week) &&
      (r.RowType === "GAME" || r.RowType === "PRESEASON")
    )
    .sort((a, b) => new Date(a.EventDate) - new Date(b.EventDate));


  const picks = sheetToObjects_(SHEET_PICKS)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      String(p.Week) === String(week) &&
      p.UserName !== PARLAY_USER
    );


  const results = sheetToObjects_(SHEET_RESULTS)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));


  const resultsByEvent = {};
  results.forEach(r => resultsByEvent[r.EventID] = r);


  const picksByEvent = {};
  picks.forEach(p => {
    if (!picksByEvent[p.EventID]) picksByEvent[p.EventID] = {};
    picksByEvent[p.EventID][p.UserName] = p.PickedTeam;
  });


  const parlayRows = sheetToObjects_(SHEET_PARLAY)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));
  const parlayByEvent = {};
  parlayRows.forEach(r => parlayByEvent[r.EventID] = r);


  const users = getUsers();
  const teamRecords = getTeamRecords_(year);
  const weekOddsRows = getOddsRowsForWeek_(year, week); // one read of the big odds sheet, shared below
  const oddsByEvent = getLatestOddsByEvent_(year, week, weekOddsRows);


  const openingOddsRows = weekOddsRows
    .filter(r => r.IsOpeningSnapshot === true || r.IsOpeningSnapshot === "TRUE");
  const openingOddsByEvent = {};
  openingOddsRows.forEach(r => {
    openingOddsByEvent[r.EventID] = {
      spread: r.Spread || "",
      awayMl: r.AwayMoneyline,
      homeMl: r.HomeMoneyline
    };
  });


  const allOddsRows = weekOddsRows;
  
  const latestFullOddsByEvent = {};
    allOddsRows.forEach(r => {
    const existing = latestFullOddsByEvent[r.EventID];
    const rTime = r.PulledAt ? new Date(r.PulledAt).getTime() : 0;
    const existingTime = existing ? existing._pulledAtMs : 0;
    if (!existing || rTime >= existingTime) {
      latestFullOddsByEvent[r.EventID] = {
        spread: r.Spread || "",
        awayMl: r.AwayMoneyline,
        homeMl: r.HomeMoneyline,
        _pulledAtMs: rTime
      };
    }
  });
  const openingSpreadByEvent = {};
  openingOddsRows.forEach(r => { openingSpreadByEvent[r.EventID] = r.Spread || ""; });


  // Standard US state/territory codes - anything outside this list on a venue
  // is treated as an international game (e.g. "VIC" for Melbourne, Australia)
  const US_STATES = new Set([
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA",
    "KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
    "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
    "VA","WA","WV","WI","WY","DC"
  ]);


  const gamesOut = games.map(g => {
    const result = resultsByEvent[g.EventID] || null;
    const parlay = parlayByEvent[g.EventID] || null;


    const eventDateStr = g.EventDate instanceof Date
      ? Utilities.formatDate(g.EventDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : g.EventDate;


    // EventTime may have been auto-converted by Sheets into a real Date/Time value
    const eventTimeStr = g.EventTime instanceof Date
      ? Utilities.formatDate(g.EventTime, Session.getScriptTimeZone(), "h:mm a")
      : g.EventTime;


    let gameTimeDisplay = "";
    if (eventDateStr) {
      const dateObj = new Date(eventDateStr + "T00:00:00");
      if (!isNaN(dateObj.getTime())) {
        const dayPart = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "EEE, MMM d");
        gameTimeDisplay = eventTimeStr ? (dayPart + " \u00b7 " + eventTimeStr + " PT") : dayPart;
      }
    }


     const state = (g.State || "").toUpperCase().trim();
     const country = (g.Country || "").toUpperCase().trim();
     const isInternational = country ? (country !== "USA") : (!!state && !US_STATES.has(state));


      const odds = oddsByEvent[g.EventID] || null;


    let spreadTrend = "";
    let oddsOpen = null;
    let oddsCurrent = null;


    const opening = openingOddsByEvent[g.EventID];
    const current = latestFullOddsByEvent[g.EventID];


    if (opening && current) {
      oddsOpen = { spread: opening.spread, awayMl: opening.awayMl, homeMl: opening.homeMl };
      oddsCurrent = { spread: current.spread, awayMl: current.awayMl, homeMl: current.homeMl };


      const openAwayMl = Number(opening.awayMl);
      const openHomeMl = Number(opening.homeMl);
      const curAwayMl = Number(current.awayMl);
      const curHomeMl = Number(current.homeMl);


      if (openAwayMl && openHomeMl && curAwayMl && curHomeMl) {
        const openFavIsAway = openAwayMl < openHomeMl;
        const curFavIsAway = curAwayMl < curHomeMl;


        if (openFavIsAway === curFavIsAway) {
          const openFavMl = Math.abs(openFavIsAway ? openAwayMl : openHomeMl);
          const curFavMl = Math.abs(curFavIsAway ? curAwayMl : curHomeMl);
          if (curFavMl > openFavMl) spreadTrend = "up";
          else if (curFavMl < openFavMl) spreadTrend = "down";
        }
      } else if (opening.spread && current.spread) {
        const openParsed = parseSpreadMagnitude_(opening.spread);
        const curParsed = parseSpreadMagnitude_(current.spread);
        if (openParsed && curParsed && openParsed.favTeam === curParsed.favTeam) {
          if (curParsed.points > openParsed.points) spreadTrend = "up";
          else if (curParsed.points < openParsed.points) spreadTrend = "down";
        }
      }
    }




    return {
      eventId: g.EventID,
      rowType: g.RowType,
      matchup: g.Matchup,
      awayTeam: g.AwayShort,
      homeTeam: g.HomeShort,
      awayTeamFull: g.AwayTeam,
      homeTeamFull: g.HomeTeam,
      eventDate: eventDateStr,
      eventTime: eventTimeStr,
      gameTimeDisplay: gameTimeDisplay,
      venue: g.Venue || "",
      city: g.City || "",
      state: g.State || "",
      isInternational: isInternational,
      status: g.Status,
      network: g.Network || "",
      spread: odds ? odds.spread : "",
      spreadTrend: spreadTrend,
      awayRecord: formatTeamRecord_(teamRecords[g.AwayShort]),
      homeRecord: formatTeamRecord_(teamRecords[g.HomeShort]),
      oddsOpen: oddsOpen,
      oddsCurrent: oddsCurrent,
      overUnder: odds ? odds.overUnder : "",
      picks: picksByEvent[g.EventID] || {},
      parlay: (parlay && Number(parlay.TotalVotes) > 0) ? {
        pickedTeam: parlay.PickedTeam || "",
        voteCount: parlay.VoteCount,
        totalVotes: parlay.TotalVotes
      } : null,
      final: result ? {
        awayScore: result.AwayScore,
        homeScore: result.HomeScore,
        winningTeam: result.WinningTeam
      } : null
    };
  });


  return { payload: { games: gamesOut, users: users }, pwGames: slimGamesForPickWindow_(games) };
}


/**
 * Lightweight read-only view for the Parlay Pick tab: just the games
 * and their computed majority-vote consensus, no individual user picks.
 * Unlike getWeekData, this doesn't hide the consensus behind a
 * "you must submit first" gate - it's a dedicated results view.
 */
/**
 * Returns the ParlayPicks rows for a week that are still tied
 * (blank PickedTeam) after considering any tie-break votes -
 * i.e. games that genuinely need a re-vote right now.
 */
/**
 * The current #1 user in season standings (by the existing % Accurate
 * ranking), excluding the computed Parlay row. Returns null if nobody
 * has any decided picks yet (too early in the season to have a leader).
 */
function getStandingsLeader_(year) {
  const standings = getStandings(year).filter(r => r.user !== PARLAY_USER);
  if (!standings.length || standings[0].accuracy === null) return null;
  return standings[0].user;
}


/**
 * Whoever had the single best record in the most recent COMPLETED
 * week (week - 1 relative to the week being tie-broken) - not
 * cumulative season standings, just that one week's performance.
 * Returns null if that week isn't decided yet, nobody picked, or
 * there's a tie for best that week (no single clear leader).
 */
function getPreviousWeekLeader_(year, week) {
  const prevWeek = Number(week) - 1;
  if (prevWeek < 1) return null;


  const results = sheetToObjects_(SHEET_RESULTS)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(prevWeek) && r.WinningTeam && r.WinningTeam !== "TIE");


  if (!results.length) return null;


  const resultByEvent = {};
  results.forEach(r => resultByEvent[r.EventID] = r.WinningTeam);


  const picks = sheetToObjects_(SHEET_PICKS)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      String(p.Week) === String(prevWeek) &&
      p.UserName !== PARLAY_USER &&
      p.PickedTeam &&
      p.PickedTeam !== "SKIP"
    );


  const tally = {};
  picks.forEach(p => {
    const winner = resultByEvent[p.EventID];
    if (winner === undefined) return;
    if (!tally[p.UserName]) tally[p.UserName] = { correct: 0, total: 0 };
    tally[p.UserName].total++;
    if (p.PickedTeam === winner) tally[p.UserName].correct++;
  });


  const rows = Object.keys(tally).map(u => ({
    user: u,
    correct: tally[u].correct,
    accuracy: tally[u].total > 0 ? tally[u].correct / tally[u].total : 0
  }));


  if (!rows.length) return null;


  rows.sort((a, b) => (b.accuracy - a.accuracy) || (b.correct - a.correct));


  const top = rows[0];
  const tiedForTop = rows.filter(r => r.accuracy === top.accuracy && r.correct === top.correct);


  return tiedForTop.length > 1 ? null : top.user;
}


/**
 * The tie-break decision-maker: previous week's individual leader
 * first, falling back to the overall season leader if the previous
 * week was tied for best or isn't decided yet. Returns
 * {user, source} or null if nobody qualifies at all (too early in
 * the season).
 */
function getTieBreakAuthority_(year, week) {
  const prevWeekLeader = getPreviousWeekLeader_(year, week);
  if (prevWeekLeader) {
    return { user: prevWeekLeader, source: "previous week" };
  }


  const overallLeader = getStandingsLeader_(year);
  if (overallLeader) {
    return { user: overallLeader, source: "season" };
  }


  return null;
}


/**
 * Lets the current season standings leader make the final call on a
 * still-tied game - a fun fallback for when the group genuinely can't
 * agree even after a re-vote. Verifies server-side that the requesting
 * user really is the current leader (the client only shows this control
 * to the leader, but this re-checks to be safe). The decision is stored
 * separately and takes top priority in computeParlayPicksForWeek, so it
 * sticks even through future re-votes on other games that week.
 */
function resolveTieByLeaderImpl_(year, week, eventId, matchup, pickedTeam, requestingUser) {
  const authority = getTieBreakAuthority_(year, week);
  if (!authority || authority.user !== requestingUser) {
    return { status: "error", message: "Only the current tie-break decision-maker can make this call." };
  }
  const leader = authority.user;


  ensureSupportSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEADER_OVERRIDE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colEventId = headers.indexOf("EventID");


  let rowNum = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colEventId]) === String(eventId)) {
      rowNum = i + 1;
      break;
    }
  }


  const row = new Array(headers.length).fill("");
  row[headers.indexOf("SeasonYear")] = year;
  row[headers.indexOf("Week")] = week;
  row[headers.indexOf("EventID")] = eventId;
  row[headers.indexOf("Matchup")] = matchup;
  row[headers.indexOf("PickedTeam")] = pickedTeam;
  row[headers.indexOf("DecidedBy")] = leader;


  if (rowNum !== -1) {
    sheet.getRange(rowNum, 1, 1, headers.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }


  computeParlayPicksForWeek(year, week);


  try {
    checkAndNotifyTieBreakStatus_(year, week);
  } catch (err) {
    Logger.log("checkAndNotifyTieBreakStatus_ failed after leader resolution: " + err.message);
  }


  return { status: "ok" };
}


function getTiedGamesForWeek(year, week) {
  computeParlayPicksForWeek(year, week); // ensure fresh, tie-break-aware


  return sheetToObjects_(SHEET_PARLAY)
    .filter(r =>
      Number(r.SeasonYear) === Number(year) &&
      String(r.Week) === String(week) &&
      !r.PickedTeam &&
      Number(r.TotalVotes) > 0
    );
}


/**
 * Data for the Tie Break screen: only the games currently tied,
 * with logos/full names like Matchups, plus who's already voted
 * in the tie-break re-vote.
 */
function getTieBreakWeekDataImpl_(year, week) {
  ensureSupportSheets();
  const tiedRows = getTiedGamesForWeek(year, week);


  const users = getUsers();
  const authority = getTieBreakAuthority_(year, week);
  const leader = authority ? authority.user : null;
  const leaderSource = authority ? authority.source : null;


  if (!tiedRows.length) {
    return JSON.stringify({ games: [], users: users, leader: leader, leaderSource: leaderSource });
  }


  const tiedEventIds = new Set(tiedRows.map(r => String(r.EventID)));


  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r =>
      Number(r.SeasonYear) === Number(year) &&
      String(r.Week) === String(week) &&
      (r.RowType === "GAME" || r.RowType === "PRESEASON") &&
      tiedEventIds.has(String(r.EventID))
    )
    .sort((a, b) => new Date(a.EventDate) - new Date(b.EventDate));


  const tieBreakVotes = sheetToObjects_(SHEET_TIEBREAK)
    .filter(p => Number(p.SeasonYear) === Number(year) && String(p.Week) === String(week));


  const votesByEvent = {};
  tieBreakVotes.forEach(p => {
    if (!votesByEvent[p.EventID]) votesByEvent[p.EventID] = {};
    votesByEvent[p.EventID][p.UserName] = p.PickedTeam;
  });


  const tallyByEvent = {};
  tiedRows.forEach(r => tallyByEvent[r.EventID] = r);


  const gamesOut = games.map(g => {
    const tally = tallyByEvent[g.EventID] || {};
    // Has this game already been through at least one full tie-break
    // round? (i.e. every active user has a tie-break vote recorded,
    // and it's STILL tied) - the leader override only becomes available
    // once that's true, so it's a genuine fallback, not a first resort.
    const eventVotes = votesByEvent[g.EventID] || {};
    const votedUsers = Object.keys(eventVotes);
    const hadFullRound = users.length > 0 && users.every(u => votedUsers.indexOf(u) !== -1);


    return {
      eventId: g.EventID,
      matchup: g.Matchup,
      awayTeam: g.AwayShort,
      homeTeam: g.HomeShort,
      awayTeamFull: g.AwayTeam,
      homeTeamFull: g.HomeTeam,
      originalVoteCount: tally.VoteCount || 0,
      originalTotalVotes: tally.TotalVotes || 0,
      picks: eventVotes,
      hadFullRound: hadFullRound
    };
  });


  return JSON.stringify({ games: gamesOut, users: users, leader: leader, leaderSource: leaderSource });
}


/**
 * Saves a user's tie-break re-votes. Separate from submitWeekPicks -
 * this never touches anyone's original Picks row, only TieBreakPicks.
 * Recomputes ParlayPicks afterward so the tie may resolve immediately,
 * and checks whether that resolves every tie for the week.
 */
function submitTieBreakVotesImpl_(year, week, userName, votesArray) {
  ensureSupportSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TIEBREAK);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];


  const colId = headers.indexOf("TieBreakID");
  const colYear = headers.indexOf("SeasonYear");
  const colWeek = headers.indexOf("Week");
  const colEventId = headers.indexOf("EventID");
  const colMatchup = headers.indexOf("Matchup");
  const colUser = headers.indexOf("UserName");
  const colPicked = headers.indexOf("PickedTeam");
  const colTime = headers.indexOf("SubmittedAt");


  const existingRowByKey = {};
  for (let i = 1; i < data.length; i++) {
    existingRowByKey[data[i][colEventId] + "||" + data[i][colUser]] = i + 1;
  }


  const now = new Date();
  const newRows = [];
  let updated = 0;


  votesArray.forEach(v => {
    if (!v.pickedTeam) return;
    const key = v.eventId + "||" + userName;
    if (existingRowByKey[key]) {
      sheet.getRange(existingRowByKey[key], colPicked + 1).setValue(v.pickedTeam);
      sheet.getRange(existingRowByKey[key], colTime + 1).setValue(now);
      updated++;
    } else {
      const row = new Array(headers.length).fill("");
      row[colId] = v.eventId + "-" + userName;
      row[colYear] = year;
      row[colWeek] = week;
      row[colEventId] = v.eventId;
      row[colMatchup] = v.matchup;
      row[colUser] = userName;
      row[colPicked] = v.pickedTeam;
      row[colTime] = now;
      newRows.push(row);
    }
  });


  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }


  computeParlayPicksForWeek(year, week);


  try {
    checkAndNotifyTieBreakStatus_(year, week);
  } catch (err) {
    Logger.log("checkAndNotifyTieBreakStatus_ failed: " + err.message);
  }


  return { status: "ok", updated: updated, created: newRows.length };
}


function buildParlayWeekDataPayload_(year, week) {
  ensureSupportSheets();
  computeParlayPicksForWeek(year, week);


  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r =>
      Number(r.SeasonYear) === Number(year) &&
      String(r.Week) === String(week) &&
      (r.RowType === "GAME" || r.RowType === "PRESEASON")
    )
    .sort((a, b) => new Date(a.EventDate) - new Date(b.EventDate));


  const parlayRows = sheetToObjects_(SHEET_PARLAY)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));
  const parlayByEvent = {};
  parlayRows.forEach(r => parlayByEvent[r.EventID] = r);


  const results = sheetToObjects_(SHEET_RESULTS)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));
  const resultsByEvent = {};
  results.forEach(r => resultsByEvent[r.EventID] = r);


  const picks = sheetToObjects_(SHEET_PICKS)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      String(p.Week) === String(week) &&
      p.UserName !== PARLAY_USER
    );
  const picksByEvent = {};
  picks.forEach(p => {
    if (!picksByEvent[p.EventID]) picksByEvent[p.EventID] = {};
    picksByEvent[p.EventID][p.UserName] = p.PickedTeam;
  });


  const users = getUsers();
  const weekOddsRows = getOddsRowsForWeek_(year, week); // one read of the big odds sheet, shared below
  const oddsByEvent = getLatestOddsByEvent_(year, week, weekOddsRows);


  const moneylineOddsRows = weekOddsRows;
  const latestMoneylineByEvent = {};
  moneylineOddsRows.forEach(r => {
    const existing = latestMoneylineByEvent[r.EventID];
    const rTime = r.PulledAt ? new Date(r.PulledAt).getTime() : 0;
    const existingTime = existing ? existing._pulledAtMs : 0;
    if (!existing || rTime >= existingTime) {
      latestMoneylineByEvent[r.EventID] = {
        spread: r.Spread || "",
        awayMoneyline: r.AwayMoneyline,
        homeMoneyline: r.HomeMoneyline,
        _pulledAtMs: rTime
      };
    }
  });




  const US_STATES = new Set([
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA",
    "KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
    "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
    "VA","WA","WV","WI","WY","DC"
  ]);


  const gamesOut = games.map(g => {
    const parlay = parlayByEvent[g.EventID] || null;
    const result = resultsByEvent[g.EventID] || null;
    const odds = oddsByEvent[g.EventID] || null;


    const eventDateStr = g.EventDate instanceof Date
      ? Utilities.formatDate(g.EventDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : g.EventDate;
    const eventTimeStr = g.EventTime instanceof Date
      ? Utilities.formatDate(g.EventTime, Session.getScriptTimeZone(), "h:mm a")
      : g.EventTime;


    let gameTimeDisplay = "";
    if (eventDateStr) {
      const dateObj = new Date(eventDateStr + "T00:00:00");
      if (!isNaN(dateObj.getTime())) {
        const dayPart = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "EEE, MMM d");
        gameTimeDisplay = eventTimeStr ? (dayPart + " \u00b7 " + eventTimeStr + " PT") : dayPart;
      }
    }


     const state = (g.State || "").toUpperCase().trim();
     const country = (g.Country || "").toUpperCase().trim();
     const isInternational = country ? (country !== "USA") : (!!state && !US_STATES.has(state));


    const parlayInfo = (parlay && Number(parlay.TotalVotes) > 0) ? {
      pickedTeam: parlay.PickedTeam || "",
      voteCount: parlay.VoteCount,
      totalVotes: parlay.TotalVotes,
      resolvedVia: parlay.ResolvedVia || ""
    } : null;


    // WIN / LOSS for the Parlay consensus pick, once the game is final
    let parlayResult = null;
    if (result && result.WinningTeam && result.WinningTeam !== "TIE" && parlayInfo && parlayInfo.pickedTeam) {
      parlayResult = (result.WinningTeam === parlayInfo.pickedTeam) ? "WIN" : "LOSS";
    }




    return {
      eventId: g.EventID,
      matchup: g.Matchup,
      awayTeam: g.AwayShort,
      homeTeam: g.HomeShort,
      awayTeamFull: g.AwayTeam,
      homeTeamFull: g.HomeTeam,
      gameTimeDisplay: gameTimeDisplay,
      venue: g.Venue || "",
      city: g.City || "",
      state: g.State || "",
      isInternational: isInternational,
      spread: odds ? odds.spread : "",
      overUnder: odds ? odds.overUnder : "",
      picks: picksByEvent[g.EventID] || {},
      parlay: parlayInfo,
      parlayResult: parlayResult,
      final: result ? {
        awayScore: result.AwayScore,
        homeScore: result.HomeScore,
        winningTeam: result.WinningTeam
      } : null
    };
  });


  const decidedThreshold = Math.floor(users.length / 2) + 1;
  let combinedDecimal = 1;
  let allDecided = true;
  const parlayLegs = [];


  gamesOut.forEach(g => {
    const pickedTeam = g.parlay && g.parlay.pickedTeam;
    const odds = latestMoneylineByEvent[g.eventId];


    let ml, source;
    if (!pickedTeam) {
      ml = -110; // still tied - flat neutral placeholder
      source = "tie-placeholder";
      allDecided = false;
    } else {
      const isDecided = g.parlay.resolvedVia === 'leader' || g.parlay.resolvedVia === 'tiebreak' || g.parlay.voteCount >= decidedThreshold;
      if (!isDecided) allDecided = false;


      if (odds && odds.awayMoneyline && odds.homeMoneyline) {
        const pickedIsAway = (pickedTeam === g.awayTeam);
        ml = pickedIsAway ? Number(odds.awayMoneyline) : Number(odds.homeMoneyline);
        source = "real";
      } else if (odds && odds.spread) {
        const parsed = parseSpreadString_(odds.spread);
        if (parsed && parsed.points > 0) {
          const result = spreadToMoneyline_(parsed.points);
          const pickedIsFavorite = (pickedTeam.toUpperCase() === parsed.favTeam);
          ml = pickedIsFavorite ? result.favoriteML : result.underdogML;
          source = "formula";
        }
      }
    }


    if (ml == null || isNaN(ml)) return;


    const dec = moneylineToDecimal_(ml);
    combinedDecimal *= dec;
    parlayLegs.push({ eventId: g.eventId, pickedTeam: pickedTeam || "", moneyline: ml, source: source });
  });


  const totalStake = 1 * users.length; // $1 per person
  const combinedMoneyline = decimalToMoneyline_(combinedDecimal);
  const totalPayout = totalStake * combinedDecimal;
  const totalProfit = totalPayout - totalStake;
  const scaledMoneyline = combinedMoneyline * users.length;


  const parlayOdds = parlayLegs.length ? {
    combinedMoneyline: combinedMoneyline,
    scaledMoneyline: scaledMoneyline,
    totalProfit: totalProfit,
    totalPayout: totalPayout,
    allDecided: allDecided,
    gamesIncluded: parlayLegs.length,
    gamesTotal: gamesOut.length
  } : null;


  // "Winning Parlay" — once every game in the week is final, what would the
  // parlay have paid if every leg had been the actual winning team, priced
  // at the CLOSING moneyline? A fun "what could have been" shown alongside
  // the live/leaning banner once results are in. Independent of what was
  // actually picked — see getParlayMoneylineBacktest for the full breakdown.
  const closingOddsForWeek = getClosingOddsByEvent_(year, week, weekOddsRows);
  let winnersCombinedDecimal = 1;
  let winnersIncludedCount = 0;
  let weekIsComplete = gamesOut.length > 0;
  gamesOut.forEach(g => {
    if (!g.final || !g.final.winningTeam) { weekIsComplete = false; return; }
    if (g.final.winningTeam === "TIE") return;
    const closing = closingOddsForWeek[g.eventId] || null;
    const winInfo = closingMoneylineForTeam_(closing, g.final.winningTeam, g.awayTeam, g.homeTeam);
    if (winInfo) {
      winnersCombinedDecimal *= moneylineToDecimal_(winInfo.ml);
      winnersIncludedCount++;
    }
  });
  const winningParlay = (weekIsComplete && winnersIncludedCount > 0) ? {
    combinedMoneyline: decimalToMoneyline_(winnersCombinedDecimal),
    payout: totalStake * winnersCombinedDecimal,
    gamesIncluded: winnersIncludedCount,
    gamesTotal: gamesOut.length
  } : null;


  const placement = getParlayPlacement_(year, week);
  return { payload: { games: gamesOut, users: users, parlayOdds: parlayOdds, winningParlay: winningParlay, placement: placement }, pwGames: null }; // the Parlay screen never reads pickWindow, so it is not computed for it
}


/**
 * Retrospective report: what would the real-money parlay result have been
 * for a given week, if bet at the CLOSING moneyline for each game, judged
 * against the actual final results?
 *
 * - Uses the same "decided" consensus pick as the live parlay (majority
 *   threshold, or leader/tiebreak resolution) — a game that never reached
 *   consensus is excluded, same as a still-tied game is excluded live.
 * - Prefers the real closing moneyline (IsClosingSnapshot=TRUE row). If that
 *   row has no moneyline (a known gap for some early-season weeks), falls
 *   back to the closing spread run through the same spread->moneyline
 *   formula used elsewhere, and flags that leg as "estimated" instead of
 *   "real" so it's clear which legs are actual sportsbook numbers.
 * - A TIE final score is treated as a push: that leg is removed from the
 *   parlay (as if it never happened) rather than counted as a loss.
 * - "Combined" figures (decimal odds / American odds / payout) are the
 *   product across every included leg regardless of win or loss — i.e.
 *   "what this parlay would have paid if every pick had hit." That's shown
 *   alongside the actual result, since a single real loss sinks the whole
 *   real parlay (standard parlay rule) and would otherwise make that number
 *   meaningless most weeks.
 */
function getParlayMoneylineBacktest(year, week) {
  ensureSupportSheets();
  computeParlayPicksForWeek(year, week);


  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r =>
      Number(r.SeasonYear) === Number(year) &&
      String(r.Week) === String(week) &&
      (r.RowType === "GAME" || r.RowType === "PRESEASON")
    )
    .sort((a, b) => new Date(a.EventDate) - new Date(b.EventDate));


  const parlayRows = sheetToObjects_(SHEET_PARLAY)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));
  const parlayByEvent = {};
  parlayRows.forEach(r => parlayByEvent[r.EventID] = r);


  const results = sheetToObjects_(SHEET_RESULTS)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));
  const resultsByEvent = {};
  results.forEach(r => resultsByEvent[r.EventID] = r);


  const users = getUsers();
  const decidedThreshold = Math.floor(users.length / 2) + 1;
  const closingOddsByEvent = getClosingOddsByEvent_(year, week);


  let anyNotFinal = false;
  let sawLoss = false;
  let combinedDecimal = 1;
  let includedCount = 0;

  // "If picks had matched the winners" — combined odds of every game's ACTUAL
  // winning team, regardless of what was picked. Independent of the parlay
  // pick entirely, so it includes every game with a final (non-tie) result
  // and a known closing moneyline for the winner.
  let winnersCombinedDecimal = 1;
  let winnersIncludedCount = 0;


  const legs = games.map(g => {
    const parlay = parlayByEvent[g.EventID] || null;
    const result = resultsByEvent[g.EventID] || null;
    const closing = closingOddsByEvent[g.EventID] || null;


    const winningTeam = (result && result.WinningTeam) || "";
    let winningTeamMl = null, winningTeamMlSource = null;
    if (winningTeam && winningTeam !== "TIE") {
      const winInfo = closingMoneylineForTeam_(closing, winningTeam, g.AwayShort, g.HomeShort);
      if (winInfo) {
        winningTeamMl = winInfo.ml;
        winningTeamMlSource = winInfo.source;
        winnersCombinedDecimal *= moneylineToDecimal_(winInfo.ml);
        winnersIncludedCount++;
      }
    }


    const base = {
      eventId: g.EventID,
      matchup: g.Matchup,
      awayTeam: g.AwayShort,
      homeTeam: g.HomeShort,
      winningTeam: winningTeam,
      winningTeamMl: winningTeamMl,
      winningTeamMlSource: winningTeamMlSource,
      finalScore: result ? { away: result.AwayScore, home: result.HomeScore } : null
    };


    const pickedTeam = parlay && parlay.PickedTeam;
    const voteCount = parlay ? Number(parlay.VoteCount) : 0;
    const resolvedVia = parlay ? (parlay.ResolvedVia || "") : "";
    const isDecided = !!pickedTeam &&
      (resolvedVia === 'leader' || resolvedVia === 'tiebreak' || voteCount >= decidedThreshold);


    if (!isDecided) {
      return Object.assign(base, { pickedTeam: pickedTeam || "", outcome: "EXCLUDED", note: "No consensus pick" });
    }


    if (!result || !result.WinningTeam) {
      anyNotFinal = true;
      return Object.assign(base, { pickedTeam: pickedTeam, outcome: "EXCLUDED", note: "Game not final yet" });
    }


    if (result.WinningTeam === "TIE") {
      return Object.assign(base, { pickedTeam: pickedTeam, outcome: "PUSH", note: "Tie — leg removed" });
    }


    const pickInfo = closingMoneylineForTeam_(closing, pickedTeam, g.AwayShort, g.HomeShort);
    if (!pickInfo) {
      return Object.assign(base, { pickedTeam: pickedTeam, outcome: "EXCLUDED", note: "No closing odds data available" });
    }


    const won = (result.WinningTeam === pickedTeam);
    const dec = moneylineToDecimal_(pickInfo.ml);
    includedCount++;
    combinedDecimal *= dec;
    if (!won) sawLoss = true;


    return Object.assign(base, {
      pickedTeam: pickedTeam,
      outcome: won ? "WIN" : "LOSS",
      moneyline: pickInfo.ml,
      decimalOdds: dec,
      oddsSource: pickInfo.source
    });
  });


  const totalStake = 1 * users.length; // $1 per person, same convention as the live parlay odds


  let actualOutcome;
  if (includedCount === 0) actualOutcome = "NO_LEGS";
  else if (anyNotFinal) actualOutcome = "INCOMPLETE";
  else if (sawLoss) actualOutcome = "LOSS";
  else actualOutcome = "WIN";

  const actualProfit = (actualOutcome === "LOSS") ? -totalStake
    : (actualOutcome === "WIN") ? (totalStake * combinedDecimal) - totalStake
    : null;

  const combinedMoneyline = includedCount ? decimalToMoneyline_(combinedDecimal) : null;
  const wouldBePayout = includedCount ? totalStake * combinedDecimal : null;

  const winnersCombinedMoneyline = winnersIncludedCount ? decimalToMoneyline_(winnersCombinedDecimal) : null;
  const winnersPayout = winnersIncludedCount ? totalStake * winnersCombinedDecimal : null;


  return {
    year: year,
    week: week,
    legs: legs,
    legsIncluded: includedCount,
    legsTotal: games.length,
    winnersIncluded: winnersIncludedCount,
    winnersCombinedMoneyline: winnersCombinedMoneyline,
    winnersPayout: winnersPayout,
    actualOutcome: actualOutcome,
    actualProfit: actualProfit,
    combinedDecimal: includedCount ? combinedDecimal : null,
    combinedMoneyline: combinedMoneyline,
    totalStake: totalStake,
    wouldBePayout: wouldBePayout
  };
}


/**
 * Save/update ALL of one user's picks for a week in one call, used by
 * the "Submit Picks" button. picksArray: [{eventId, matchup, pickedTeam}, ...]
 * Entries with an empty pickedTeam are skipped (no accidental overwrite).
 */
function submitWeekPicksImpl_(year, week, userName, picksArray) {
  ensureSupportSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PICKS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];


  const colPickId = headers.indexOf("PickID");
  const colSeasonYear = headers.indexOf("SeasonYear");
  const colWeek = headers.indexOf("Week");
  const colEventId = headers.indexOf("EventID");
  const colMatchup = headers.indexOf("Matchup");
  const colUser = headers.indexOf("UserName");
  const colPicked = headers.indexOf("PickedTeam");
  const colTime = headers.indexOf("SubmittedAt");


  const existingRowByKey = {};
  for (let i = 1; i < data.length; i++) {
    existingRowByKey[data[i][colEventId] + "||" + data[i][colUser]] = i + 1;
  }


  const now = new Date();
  const newRows = [];
  let updated = 0;


  picksArray.forEach(p => {
    if (!p.pickedTeam) return;
    const key = p.eventId + "||" + userName;
    if (existingRowByKey[key]) {
      sheet.getRange(existingRowByKey[key], colPicked + 1).setValue(p.pickedTeam);
      sheet.getRange(existingRowByKey[key], colTime + 1).setValue(now);
      updated++;
    } else {
      const row = new Array(headers.length).fill("");
      row[colPickId] = p.eventId + "-" + userName;
      row[colSeasonYear] = year;
      row[colWeek] = week;
      row[colEventId] = p.eventId;
      row[colMatchup] = p.matchup;
      row[colUser] = userName;
      row[colPicked] = p.pickedTeam;
      row[colTime] = now;
      newRows.push(row);
    }
  });


  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }


  // Check if this submission completes the week (every user, every game)
  // and text the consensus summary if so. Never let a notification
  // failure break the actual pick submission.
  try {
    checkAndNotifyIfWeekComplete_(year, week);
  } catch (err) {
    Logger.log("checkAndNotifyIfWeekComplete_ failed: " + err.message);
  }


  return { status: "ok", updated: updated, created: newRows.length };
}


function submitPickImpl_(year, week, eventId, matchup, userName, pickedTeam) {
  ensureSupportSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PICKS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];


  const colEventId = headers.indexOf("EventID");
  const colUser = headers.indexOf("UserName");
  const colPicked = headers.indexOf("PickedTeam");
  const colTime = headers.indexOf("SubmittedAt");


  for (let i = 1; i < data.length; i++) {
    if (data[i][colEventId] === eventId && data[i][colUser] === userName) {
      sheet.getRange(i + 1, colPicked + 1).setValue(pickedTeam);
      sheet.getRange(i + 1, colTime + 1).setValue(new Date());
      return { status: "updated" };
    }
  }


  const colPickId = headers.indexOf("PickID");
  const colSeasonYear = headers.indexOf("SeasonYear");
  const colWeek = headers.indexOf("Week");
  const colMatchup = headers.indexOf("Matchup");


  const row = new Array(headers.length).fill("");
  row[colPickId] = eventId + "-" + userName;
  row[colSeasonYear] = year;
  row[colWeek] = week;
  row[colEventId] = eventId;
  row[colMatchup] = matchup;
  row[colUser] = userName;
  row[colPicked] = pickedTeam;
  row[colTime] = new Date();
  sheet.appendRow(row);
  return { status: "created" };
}


/**
 * Pulls final scores from ESPN for a given week/year and writes
 * them into API_NFL_Results. Safe to re-run (upserts by EventID).
 * Only writes games ESPN marks as completed - in-progress games are skipped.
 */
function pullFinalScoresImpl_(year, week) {
  ensureSupportSheets();


  const weekStr = String(week);
  const isPreseason = weekStr.charAt(0) === "P";
  const seasonType = isPreseason ? 1 : 2;
  const espnWeek = isPreseason ? weekStr.substring(1) : weekStr;


  const url =
    "https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard" +
    "?season=" + year +
    "&seasontype=" + seasonType +
    "&week=" + espnWeek;


  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    return { status: "error", message: "ESPN request failed with code " + response.getResponseCode() };
  }


  const data = JSON.parse(response.getContentText());
  const events = data.events || [];


  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
  const existing = sheet.getDataRange().getValues();
  const headers = existing[0];
  const colEventId = headers.indexOf("EventID");


  const existingRowByEvent = {};
  for (let i = 1; i < existing.length; i++) {
    existingRowByEvent[existing[i][colEventId]] = i + 1; // 1-based sheet row
  }


  let updated = 0;
  let skippedIncomplete = 0;


  events.forEach(event => {
    const competition = (event.competitions && event.competitions[0]) || {};
    const status = competition.status || {};
    const isComplete = !!(status.type && status.type.completed === true);


    if (!isComplete) {
      skippedIncomplete++;
      return;
    }


    const competitors = competition.competitors || [];
    const away = competitors.find(c => c.homeAway === "away") || {};
    const home = competitors.find(c => c.homeAway === "home") || {};


    const awayTeam = (away.team && (away.team.abbreviation || away.team.displayName)) || "";
    const homeTeam = (home.team && (home.team.abbreviation || home.team.displayName)) || "";
    const awayScore = Number(away.score);
    const homeScore = Number(home.score);
    const matchup = awayTeam + " @ " + homeTeam;


    let winningTeam = "";
    if (awayScore > homeScore) winningTeam = awayTeam;
    else if (homeScore > awayScore) winningTeam = homeTeam;
    else winningTeam = "TIE";


    const resHeaders = existing[0];
    const rCol = name => resHeaders.indexOf(name);
    const row = new Array(resHeaders.length).fill("");
    row[rCol("SeasonYear")] = year;
    row[rCol("Week")] = week;
    row[rCol("EventID")] = event.id;
    row[rCol("Matchup")] = matchup;
    row[rCol("AwayScore")] = awayScore;
    row[rCol("HomeScore")] = homeScore;
    row[rCol("WinningTeam")] = winningTeam;
    row[rCol("UpdatedAt")] = new Date();


    if (existingRowByEvent[event.id]) {
      sheet.getRange(existingRowByEvent[event.id], 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    updated++;
  });


  try {
    checkAndNotifyRunbackStatus_(year, week);
  } catch (err) {
    Logger.log("checkAndNotifyRunbackStatus_ failed after pullFinalScores: " + err.message);
  }


  return { status: "ok", updated: updated, skippedIncomplete: skippedIncomplete };
}


/**
 * Standings for a season: Correct / Incorrect / Pending counts, plus
 * % Accurate computed as the AVERAGE of each week's accuracy rate
 * (correct / decided that week), not a flat total-correct/total-picks
 * ratio. A user who goes 3/3 one week and 0/5 the next sits at 50%,
 * not 30%. Weeks with no decided games yet don't count toward the average.
 */
function getStandingsImpl_(year, week) {
  const results = sheetToObjects_(SHEET_RESULTS)
    .filter(r => Number(r.SeasonYear) === Number(year) && r.WinningTeam && r.WinningTeam !== "TIE");


  const resultByEvent = {};
  results.forEach(r => resultByEvent[r.EventID] = r.WinningTeam);


   const picks = sheetToObjects_(SHEET_PICKS)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      p.UserName !== PARLAY_USER &&
      p.PickedTeam &&
      p.PickedTeam !== "SKIP" &&
      String(p.Week).charAt(0) !== "P" && // preseason doesn't count toward real standings
      (!week || String(p.Week) === String(week))
    );


  const users = getUsers();
  const byUser = {};
  users.forEach(u => byUser[u] = {});


  const tallyPick = (userName, week, eventId, pickedTeam) => {
    if (!byUser[userName]) byUser[userName] = {};
    if (!byUser[userName][week]) byUser[userName][week] = { correct: 0, incorrect: 0, pending: 0 };
    const winner = resultByEvent[eventId];
    if (winner === undefined) {
      byUser[userName][week].pending++;
    } else if (winner === pickedTeam) {
      byUser[userName][week].correct++;
    } else {
      byUser[userName][week].incorrect++;
    }
  };
    const tieBreakVotes = sheetToObjects_(SHEET_TIEBREAK)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      p.PickedTeam &&
      String(p.Week).charAt(0) !== "P" &&
      (!week || String(p.Week) === String(week))
    );
  const tieBreakByUserEvent = {};
  tieBreakVotes.forEach(p => {
    tieBreakByUserEvent[p.UserName + "||" + p.EventID] = p.PickedTeam;
  });


  picks.forEach(p => {
    const key = p.UserName + "||" + p.EventID;
    const effectivePick = tieBreakByUserEvent[key] || p.PickedTeam;
    tallyPick(p.UserName, String(p.Week), p.EventID, effectivePick);
  });


 
  // Parlay is scored from the computed ParlayPicks table, not from Picks
    const parlayRows = sheetToObjects_(SHEET_PARLAY)
    .filter(r => Number(r.SeasonYear) === Number(year) && r.PickedTeam && String(r.Week).charAt(0) !== "P" && (!week || String(r.Week) === String(week)));
  byUser[PARLAY_USER] = {};
  parlayRows.forEach(r => tallyPick(PARLAY_USER, String(r.Week), r.EventID, r.PickedTeam));


  const rows = Object.keys(byUser).map(u => {
    const weeks = byUser[u];
    let totalCorrect = 0, totalIncorrect = 0, totalPending = 0;
    const weeklyAccuracies = [];


    Object.keys(weeks).forEach(w => {
      const wk = weeks[w];
      totalCorrect += wk.correct;
      totalIncorrect += wk.incorrect;
      totalPending += wk.pending;
      const decided = wk.correct + wk.incorrect;
      if (decided > 0) {
        weeklyAccuracies.push(wk.correct / decided);
      }
    });


    const accuracy = weeklyAccuracies.length
      ? Math.round((weeklyAccuracies.reduce((a, b) => a + b, 0) / weeklyAccuracies.length) * 1000) / 10
      : null; // null = no decided weeks yet, shown as "--" in the UI


    return {
      user: u,
      correct: totalCorrect,
      incorrect: totalIncorrect,
      pending: totalPending,
      accuracy: accuracy
    };
  });


  return rows.sort((a, b) => {
    const aAcc = a.accuracy === null ? -1 : a.accuracy;
    const bAcc = b.accuracy === null ? -1 : b.accuracy;
    return bAcc - aAcc;
  });
}


// ------------------------------------------------------------
// FAST BULK GENERATE - replaces generateAllParlayPicksAllYears.
// That version called computeParlayPicksForWeek once per
// year/week (54 times for 3 seasons), and each call re-read the
// entire NFL_Schedule, Picks, and ParlayPicks sheets plus did up
// to 16 individual write calls - hundreds of operations total,
// which is exactly what triggers Google's "Service Spreadsheets
// failed" rate-limit error on a spreadsheet this size.
//
// This version reads NFL_Schedule, Picks, TieBreakPicks, and
// LeaderOverrides exactly ONCE, computes every week's final pick
// in memory using the SAME priority logic as computeParlayPicksForWeek
// (main vote -> leader override if tied -> tie-break re-vote if tied
// and everyone's voted), then fully rebuilds ParlayPicks in a single
// write. ParlayPicks is entirely derived data (nothing manual lives
// there), so a full rebuild each time is safe and far faster than
// trying to update rows in place.
// ------------------------------------------------------------
/**
 * Determines the minimum SeasonYear that generateAllParlayPicksAllYearsFast
 * should actually recompute. Checks for a manual override in a "Config"
 * sheet (a row where column A = "MinYearToRecompute" and column B has a
 * value) - if present, uses that. Otherwise defaults to current season
 * minus 1, so last season stays in scope for late corrections without
 * ever needing to be manually updated as seasons roll forward. If no
 * Config sheet exists at all, that's fine - it just means no override,
 * same as the default behavior.
 */
function getMinYearToRecompute_() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (configSheet && configSheet.getLastRow() > 1) {
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === "MinYearToRecompute" && data[i][1]) {
        return Number(data[i][1]);
      }
    }
  }
  const currentYear = getCurrentSeasonYear();
  return currentYear ? currentYear - 1 : 0;
}


function generateAllParlayPicksAllYearsFastImpl_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const minYear = getMinYearToRecompute_();
  const headers = ["SeasonYear", "Week", "EventID", "Matchup", "PickedTeam", "VoteCount", "TotalVotes", "GeneratedAt"];


  let sheet = ss.getSheetByName(SHEET_PARLAY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PARLAY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    applyStandardFont_(sheet);
  }


  // Read source data ONLY for years at or above the cutoff - years below
  // minYear are never read here at all, let alone recomputed.
  const scheduleRows = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => (r.RowType === "GAME" || r.RowType === "PRESEASON") && Number(r.SeasonYear) >= minYear);
  const gamesByYearWeek = {};
  scheduleRows.forEach(g => {
    const key = g.SeasonYear + "|" + g.Week;
    if (!gamesByYearWeek[key]) gamesByYearWeek[key] = [];
    gamesByYearWeek[key].push(g);
  });


  const pickRows = sheetToObjects_(SHEET_PICKS)
    .filter(p =>
      p.UserName !== PARLAY_USER &&
      p.PickedTeam &&
      p.PickedTeam !== "SKIP" &&
      Number(p.SeasonYear) >= minYear
    );
  const votesByKey = {};
  pickRows.forEach(p => {
    const key = p.SeasonYear + "|" + p.Week + "|" + p.EventID;
    if (!votesByKey[key]) votesByKey[key] = {};
    votesByKey[key][p.PickedTeam] = (votesByKey[key][p.PickedTeam] || 0) + 1;
  });


  const tieBreakRows = sheetToObjects_(SHEET_TIEBREAK)
    .filter(p => p.PickedTeam && Number(p.SeasonYear) >= minYear);
  const tieBreakVotesByKey = {};
  const tieBreakVotersByKey = {};
  tieBreakRows.forEach(p => {
    const key = p.SeasonYear + "|" + p.Week + "|" + p.EventID;
    if (!tieBreakVotesByKey[key]) tieBreakVotesByKey[key] = {};
    tieBreakVotesByKey[key][p.PickedTeam] = (tieBreakVotesByKey[key][p.PickedTeam] || 0) + 1;
    if (!tieBreakVotersByKey[key]) tieBreakVotersByKey[key] = new Set();
    tieBreakVotersByKey[key].add(p.UserName);
  });


  const leaderOverrideRows = sheetToObjects_(SHEET_LEADER_OVERRIDE)
    .filter(p => p.PickedTeam && Number(p.SeasonYear) >= minYear);
  const leaderOverrideByKey = {};
  leaderOverrideRows.forEach(p => {
    const key = p.SeasonYear + "|" + p.Week + "|" + p.EventID;
    leaderOverrideByKey[key] = p.PickedTeam;
  });


  const activeUsers = getUsers();


  // Same vote-resolution helper as computeParlayPicksForWeek
  const resolveVoteMap_ = voteMap => {
    const teams = Object.keys(voteMap);
    let top = "";
    let count = 0;
    teams.forEach(t => {
      if (voteMap[t] > count) {
        count = voteMap[t];
        top = t;
      }
    });
    const topTeams = teams.filter(t => voteMap[t] === count);
    return { pickedTeam: topTeams.length > 1 ? "" : top, voteCount: count, tied: topTeams.length > 1 };
  };


  // Compute the fresh replacement rows for every in-scope game, entirely
  // in memory - mirrors computeParlayPicksForWeek's exact priority order.
  const now = new Date();
  const freshRows = [];


  Object.keys(gamesByYearWeek).forEach(yearWeekKey => {
    const games = gamesByYearWeek[yearWeekKey];


    games.forEach(g => {
      const key = g.SeasonYear + "|" + g.Week + "|" + g.EventID;
      const voteMap = votesByKey[key] || {};
      const teams = Object.keys(voteMap);


      let pickedTeam = "";
      let voteCount = 0;
      let resolvedVia = "";
      const totalVotes = teams.reduce((sum, t) => sum + voteMap[t], 0);


      if (teams.length) {
        const main = resolveVoteMap_(voteMap);
        pickedTeam = main.pickedTeam;
        voteCount = main.voteCount;


        if (main.tied && leaderOverrideByKey[key]) {
          pickedTeam = leaderOverrideByKey[key];
          resolvedVia = "leader";
        } else if (main.tied) {
          const votersSet = tieBreakVotersByKey[key] || new Set();
          const everyoneVoted = activeUsers.length > 0 && activeUsers.every(u => votersSet.has(u));


          if (everyoneVoted) {
            const tieBreak = resolveVoteMap_(tieBreakVotesByKey[key] || {});
            if (tieBreak.pickedTeam) {
              pickedTeam = tieBreak.pickedTeam;
              resolvedVia = "tiebreak";
            }
          }
        }
      }


      freshRows.push([g.SeasonYear, g.Week, g.EventID, g.Matchup, pickedTeam, voteCount, totalVotes, now, resolvedVia]);
    });
  });


  // Remove ONLY the in-scope rows (year >= minYear) from the sheet - never
  // the whole sheet, never a full delete/recreate. Since the sheet is kept
  // sorted by SeasonYear ascending after every run (see bottom of this
  // function), in-scope rows always form ONE contiguous block at the end,
  // so this is a single deleteRows call regardless of how many rows that
  // spans - not one call per row, which would risk the same rate-limit
  // crash the original slow version had. Historical rows above the cutoff
  // are never selected, read, or touched by this at all.
  const lastRow = sheet.getLastRow();
  let removedCount = 0;


  if (lastRow > 1) {
    const yearColumnValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let firstInScopeRow = -1;
    for (let i = 0; i < yearColumnValues.length; i++) {
      if (Number(yearColumnValues[i][0]) >= minYear) {
        firstInScopeRow = i + 2; // convert to 1-based sheet row
        break;
      }
    }
    if (firstInScopeRow !== -1) {
      removedCount = lastRow - firstInScopeRow + 1;
      sheet.deleteRows(firstInScopeRow, removedCount);
    }
  }


  // Append the freshly computed rows in ONE batch write
  if (freshRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, freshRows.length, headers.length).setValues(freshRows);
  }


  // Keep sorted by SeasonYear/Week ascending so in-scope rows stay one
  // clean contiguous block for the next run's single-range delete
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length)
      .sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
  }


  // Formatting happens LAST, after the actual data is safely written -
  // a formatting failure should never be able to prevent the data write
  applyStandardFont_(sheet);


  Logger.log("ParlayPicks upsert: " + removedCount + " in-scope row(s) removed in a single operation, " +
    freshRows.length + " freshly recomputed row(s) written. Years below " + minYear +
    " were never read, touched, or rewritten.");
  Logger.log("Year/week combinations recomputed: " + Object.keys(gamesByYearWeek).length);


  return {
    status: "ok",
    minYear: minYear,
    removedInScopeRows: removedCount,
    freshRowsWritten: freshRows.length,
    yearWeeksProcessed: Object.keys(gamesByYearWeek).length
  };
}


// Standard normal CDF approximation (Abramowitz & Stegun 7.1.26)
function normalDist_(z) {
  const sign = z < 0 ? -1 : 1;
  z = Math.abs(z) / Math.sqrt(2);
  const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741, a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
  const t = 1 / (1 + p * z);
  const y = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-z * z);
  return 0.5 * (1 + sign * y);
}


function spreadToMoneyline_(spreadPoints) {
  const p = normalDist_(spreadPoints / 13.86); // favorite's win probability
  const favML = Math.round(-100 * (p / (1 - p)));
  const dogML = Math.round(100 * (p / (1 - p)));
  return { favoriteProb: p, favoriteML: favML, underdogML: dogML };
}


function moneylineToDecimal_(ml) {
  return ml < 0 ? 1 + (100 / Math.abs(ml)) : 1 + (ml / 100);
}


function decimalToMoneyline_(dec) {
  return dec >= 2 ? Math.round((dec - 1) * 100) : Math.round(-100 / (dec - 1));
}


function parseSpreadString_(spreadText) {
  if (!spreadText) return null;
  const trimmed = String(spreadText).trim();
  if (/^(EVEN|PK)$/i.test(trimmed)) return { favTeam: null, points: 0 };
  const match = trimmed.match(/^([A-Za-z]+)\s*(-?\d+(\.\d+)?)$/);
  if (!match) return null;
  return { favTeam: match[1].toUpperCase(), points: Math.abs(parseFloat(match[2])) };
}
function ensureParlayPlacementsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ParlayPlacements");
  if (!sheet) {
    sheet = ss.insertSheet("ParlayPlacements");
    sheet.getRange(1, 1, 1, 4).setValues([["SeasonYear", "Week", "PlacedBy", "PlacedAt"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}


function recordParlayPlacementImpl_(year, week, userName) {
  if (!userName) return { status: "error", message: "No name selected." };


  const sheet = ensureParlayPlacementsSheet_();
  const data = sheet.getDataRange().getValues();


  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(year) && String(data[i][1]) === String(week)) {
      rowIndex = i + 1;
      break;
    }
  }


  const now = new Date();
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex, 3, 1, 2).setValues([[userName, now]]);
  } else {
    sheet.appendRow([year, week, userName, now]);
  }

  // Tell the group the Parlay is in (once per week). A notification problem
  // must never undo or fail the placement itself.
  try {
    notifyParlayPlaced_(year, week, userName);
  } catch (err) {
    Logger.log("notifyParlayPlaced_ failed: " + err.message);
  }

    return { status: "ok", placedBy: userName, placedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd HH:mm") };
}


function getParlayPlacement_(year, week) {
  const sheet = ensureParlayPlacementsSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(year) && String(data[i][1]) === String(week)) {
            const placedAt = data[i][3] instanceof Date
        ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), "MM/dd HH:mm")
        : String(data[i][3] || "");
      return { placedBy: data[i][2] || "", placedAt: placedAt };
    }
  }
  return null;
}


function recordRunbackPlacementImpl_(callerUserName, year, week, participants, madeByUserName) {
  if (!isUserAdmin_(callerUserName)) return { status: "error", message: "Not authorized." };
  if (!madeByUserName) return { status: "error", message: "Select who made the runback." };


  const sheet = ensureParlayPlacementsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];


  const colRunBack = headers.indexOf("RunBack");
  const colParticipants = headers.indexOf("RunBackParticipants");
  const colMadeBy = headers.indexOf("RunBackMadeBy");
  const colMadeAt = headers.indexOf("Runback MadeAt");


  if (colRunBack === -1 || colParticipants === -1 || colMadeBy === -1 || colMadeAt === -1) {
    return { status: "error", message: "RunBack columns not found in ParlayPlacements sheet." };
  }


  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(year) && String(data[i][1]) === String(week)) {
      rowIndex = i + 1;
      break;
    }
  }


  const now = new Date();
  if (rowIndex === -1) {
    const newRow = new Array(headers.length).fill("");
    newRow[0] = year;
    newRow[1] = week;
    newRow[colRunBack] = "Yes";
    newRow[colParticipants] = participants;
    newRow[colMadeBy] = madeByUserName;
    newRow[colMadeAt] = now;
    sheet.appendRow(newRow);
  } else {
    sheet.getRange(rowIndex, colRunBack + 1).setValue("Yes");
    sheet.getRange(rowIndex, colParticipants + 1).setValue(participants);
    sheet.getRange(rowIndex, colMadeBy + 1).setValue(madeByUserName);
    sheet.getRange(rowIndex, colMadeAt + 1).setValue(now);
  }

  try {
    notifyRunbackPlaced_(year, week, participants, madeByUserName);
  } catch (err) {
    Logger.log("notifyRunbackPlaced_ failed: " + err.message);
  }

  return {
    status: "ok",
    participants: participants,
    madeBy: madeByUserName,
    madeAt: Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd HH:mm")
  };
}


function getRunbackInfo_(year, week) {
  const sheet = ensureParlayPlacementsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];


  const colRunBack = headers.indexOf("RunBack");
  const colParticipants = headers.indexOf("RunBackParticipants");
  const colMadeBy = headers.indexOf("RunBackMadeBy");
  const colMadeAt = headers.indexOf("Runback MadeAt");
  if (colRunBack === -1) return null;


  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(year) && String(data[i][1]) === String(week)) {
      if (data[i][colRunBack] !== "Yes") return null;
      const madeAtRaw = data[i][colMadeAt];
      const madeAt = madeAtRaw instanceof Date
        ? Utilities.formatDate(madeAtRaw, Session.getScriptTimeZone(), "MM/dd HH:mm")
        : String(madeAtRaw || "");
      return {
        participants: data[i][colParticipants] || "",
        madeBy: data[i][colMadeBy] || "",
        madeAt: madeAt
      };
    }
  }
  return null;
}


function getRunbackForAdmin(callerUserName, year, week) {
  if (!isUserAdmin_(callerUserName)) return { status: "error", message: "Not authorized." };
  return { status: "ok", runback: getRunbackInfo_(year, week) };
}


// ==============================================================
// HOME PAGE - NOTICES
// Two kinds of notice feed into one list on the Home tab:
//   - POSTED notices: any logged-in user can post free text. Stored
//     in the Notices sheet. Author or an admin can delete one.
//   - SYSTEM notices: computed live every time the feed loads,
//     never stored as their own row (e.g. "you haven't picked yet").
//     Each uses a stable synthetic ID (tied to year/week) so a
//     per-user dismissal still sticks even though the notice itself
//     isn't a sheet row.
// "New" vs "dismissed" are two different things: new is just a
// highlight that clears once you've opened Home (tracked via
// UserFeedState.LastSeenAt); dismissed permanently drops a notice
// from YOUR feed only, via a NoticeDismissals row - other people's
// feeds are untouched.
// ==============================================================

function ensureNoticesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NOTICES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NOTICES);
    sheet.getRange(1, 1, 1, 4).setValues([["NoticeID", "PostedBy", "PostedAt", "Text"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    applyStandardFont_(sheet);
  }
  return sheet;
}

function ensureNoticeDismissalsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NOTICE_DISMISSALS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NOTICE_DISMISSALS);
    sheet.getRange(1, 1, 1, 3).setValues([["NoticeID", "UserName", "DismissedAt"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    applyStandardFont_(sheet);
  }
  return sheet;
}

function ensureUserFeedStateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_USER_FEED_STATE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USER_FEED_STATE);
    sheet.getRange(1, 1, 1, 2).setValues([["UserName", "LastSeenAt"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    applyStandardFont_(sheet);
  }
  return sheet;
}

/**
 * Anyone logged in can post a notice - no admin gate. Returns the
 * created notice so the client can render it immediately without
 * a round trip.
 */
function postNotice(userName, text) {
  const cleanText = String(text || "").trim();
  if (!userName) return { status: "error", message: "Not logged in." };
  if (!cleanText) return { status: "error", message: "Notice can't be empty." };

  const sheet = ensureNoticesSheet_();
  const now = new Date();
  const noticeId = "n" + now.getTime() + "_" + Math.floor(Math.random() * 1000);
  sheet.appendRow([noticeId, userName, now, cleanText]);
  bumpDataVersion_();

  return {
    status: "ok",
    notice: { noticeId: noticeId, postedBy: userName, postedAt: now.toISOString(), text: cleanText }
  };
}

/**
 * Writes a group-event announcement (consensus recap, tie alert/
 * resolution, runback passed/failed) into the Notices sheet as a
 * normal posted notice, authored as "System" so it shows up on Home
 * for everyone exactly like a manually posted note - marked "new,"
 * dismissible per-user, and deletable only by an admin (there's no
 * real author to match against). Called from Automation.gs every
 * time sendSmsToAllUsers_ fires for a group-wide category, so Home
 * always carries the same announcement the SMS did - even for
 * someone who's opted out of that text, or just missed it. Not used
 * for the runback-OPEN moment - that already gets a live, interactive
 * notice with Yes/No voting further down in getHomeFeed, and a
 * second static copy of it would just go stale once the vote closes.
 */
function postSystemNotice_(text) {
  const cleanText = String(text || "").trim();
  if (!cleanText) return;
  const sheet = ensureNoticesSheet_();
  const now = new Date();
  const noticeId = "n" + now.getTime() + "_" + Math.floor(Math.random() * 1000);
  sheet.appendRow([noticeId, "System", now, cleanText]);
  bumpDataVersion_();
}

/**
 * Removes a posted notice entirely (for everyone) - only the
 * original author or an admin may do this. System notices have no
 * row to delete, so this only ever touches the Notices sheet.
 */
function deleteNotice(callerUserName, noticeId) {
  const sheet = ensureNoticesSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colId = headers.indexOf("NoticeID");
  const colPostedBy = headers.indexOf("PostedBy");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId]) === String(noticeId)) {
      const postedBy = data[i][colPostedBy];
      if (postedBy !== callerUserName && !isUserAdmin_(callerUserName)) {
        return { status: "error", message: "Only the author or an admin can remove this." };
      }
      sheet.deleteRow(i + 1);
      bumpDataVersion_();
      return { status: "ok" };
    }
  }
  return { status: "error", message: "Notice not found." };
}

/**
 * Hides one notice from THIS user's feed only, forever (or until a
 * system notice's underlying condition recurs with a new synthetic
 * ID, e.g. next week's pick reminder). Never deletes the underlying
 * Notices row, and never affects anyone else's feed.
 */
function dismissNotice(userName, noticeId) {
  if (!userName) return { status: "error", message: "Not logged in." };
  const sheet = ensureNoticeDismissalsSheet_();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(noticeId) && String(data[i][1]) === String(userName)) {
      return { status: "ok" }; // already dismissed
    }
  }
  sheet.appendRow([noticeId, userName, new Date()]);
  return { status: "ok" };
}

/**
 * Whether userName has submitted a pick for every game in this
 * week yet (used for the "you haven't picked" system notice).
 */
function hasUserPickedAllGamesThisWeek_(userName, year, week) {
  const snap = cachedLookup_("weekpicks", year + "|" + week, function () {
    const gameIds = sheetToObjects_(SHEET_SCHEDULE)
      .filter(r =>
        Number(r.SeasonYear) === Number(year) &&
        String(r.Week) === String(week) &&
        (r.RowType === "GAME" || r.RowType === "PRESEASON")
      )
      .map(r => String(r.EventID));
    const picked = {};
    sheetToObjects_(SHEET_PICKS).forEach(p => {
      if (Number(p.SeasonYear) === Number(year) && String(p.Week) === String(week) && p.PickedTeam) {
        const u = String(p.UserName);
        if (!picked[u]) picked[u] = [];
        picked[u].push(String(p.EventID));
      }
    });
    return { gameIds: gameIds, picked: picked };
  });

  if (!snap.gameIds.length) return true;
  const mine = new Set(snap.picked[String(userName)] || []);
  return snap.gameIds.every(id => mine.has(id));
}

/**
 * Everything for the Home tab: posted notices (not dismissed by
 * this user) plus any system notices that currently apply to them,
 * newest first. Marks isNew based on this user's LastSeenAt, then
 * bumps LastSeenAt to now - so the "new" highlight always reflects
 * what's changed since their last visit to Home, not since forever.
 */
function getHomeFeed(userName) {
  if (!userName) return { status: "error", message: "Not logged in." };
  ensureSupportSheets();

  const year = getCurrentSeasonYear();
  const week = year ? getDefaultWeek(year) : null;

  // Opportunistic safety net: also check runback status here, since
  // Home is the page everyone hits most often - this way the vote
  // still surfaces even if the scheduled final-score trigger is the
  // one that's supposed to have kicked it off but didn't fire.
  // Throttled so a burst of Home opens doesn't repeat the same check.
  if (year && week) throttledRunbackCheck_(year, week);

  const dismissedIds = new Set(
    sheetToObjects_(SHEET_NOTICE_DISMISSALS)
      .filter(d => d.UserName === userName)
      .map(d => String(d.NoticeID))
  );

  const lastSeenAt = readFeedState_(userName).lastSeenAt;
  const isAdmin = isUserAdmin_(userName);

  // Posted notices are the same for everyone, so they come from a shared
  // cached copy (refreshed whenever a notice is posted or removed).
  const postedNotices = getSharedNotices_()
    .filter(n => !dismissedIds.has(String(n.noticeId)))
    .map(n => {
      const postedAt = new Date(n.postedAt);
      return {
        noticeId: String(n.noticeId),
        kind: "posted",
        postedBy: n.postedBy,
        postedAt: n.postedAt,
        text: n.text,
        isNew: !lastSeenAt || postedAt > lastSeenAt,
        canDelete: (n.postedBy === userName || isAdmin)
      };
    })
    .sort((a, b) => new Date(b.postedAt) - new Date(a.postedAt));

  const systemNotices = [];
  if (year && week) {
    if (!hasUserPickedAllGamesThisWeek_(userName, year, week)) {
      const noticeId = "sys-pick-" + year + "-" + week;
      if (!dismissedIds.has(noticeId)) {
        systemNotices.push({
          noticeId: noticeId,
          kind: "system",
          postedBy: "",
          postedAt: new Date().toISOString(),
          text: "You haven't picked yet for Week " + week + ".",
          isNew: true,
          canDelete: false
        });
      }
    }

    try {
      const runback = getRunbackStatus(userName, year, week);
      if (runback.status === "ok" && runback.eligible && runback.open) {
        const noticeId = "sys-runback-" + year + "-" + week;
        if (!dismissedIds.has(noticeId)) {
          systemNotices.push({
            noticeId: noticeId,
            kind: "runback",
            postedBy: "",
            postedAt: new Date().toISOString(),
            text: runback.firstGameMatchup + " didn't go our way - run the Parlay back on the rest of the week?",
            isNew: true,
            canDelete: false,
            runback: runback
          });
        }
      }
    } catch (err) {
      Logger.log("getRunbackStatus failed in getHomeFeed: " + err.message);
    }

    getHomeCards_(year, week).forEach(card => systemNotices.push(card));
  }

  // NOTE: this no longer records "last seen" - the page calls markHomeSeen
  // right after it has drawn the feed, so the write doesn't delay the screen.
  return {
    status: "ok",
    year: year,
    week: week,
    notices: systemNotices.concat(postedNotices)
  };
}

/**
 * Records that userName has now seen the Home feed (drives the "new"
 * highlighting next visit). Called by the page after Home has rendered.
 */
function markHomeSeen(userName) {
  if (!userName) return { status: "error", message: "Not logged in." };
  const state = readFeedState_(userName);
  const now = new Date();
  if (state.rowIndex === -1) {
    state.sheet.appendRow([userName, now]);
  } else {
    state.sheet.getRange(state.rowIndex, state.colSeen + 1).setValue(now);
  }
  return { status: "ok" };
}

function readFeedState_(userName) {
  const sheet = ensureUserFeedStateSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colUser = headers.indexOf("UserName");
  const colSeen = headers.indexOf("LastSeenAt");
  let rowIndex = -1;
  let lastSeenAt = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colUser]) === String(userName)) {
      rowIndex = i + 1;
      lastSeenAt = data[i][colSeen] instanceof Date ? data[i][colSeen] : null;
      break;
    }
  }
  return { sheet: sheet, rowIndex: rowIndex, colSeen: colSeen, lastSeenAt: lastSeenAt };
}

// ------------------------------------------------------------
// HOME CARDS: last week's results table + this week's Parlay / Runback.
// Both are the same for everyone, so they are cached (cachedLookup_) and
// pre-built by the background warm-up - Home just reads them.
//  - Results: for the week BEFORE the current one. The current week flips
//    on Tuesday, so the table is replaced then with the new "previous week".
//  - Parlay: shown once this week's Parlay is placed; replaced by the
//    Runback once a runback is confirmed as placed. When the week flips the
//    new week has no placement yet, so the old Parlay card drops away at the
//    same moment the new table appears.
// ------------------------------------------------------------
function getHomeResultsCard_(year, week) {
  const prev = Number(week) - 1;
  if (!(prev >= 1)) return null; // preseason ("P1"...) or week 1: nothing before it
  return cachedLookup_("homeresults", year + "|" + prev, function () {
    const wkRows = getStandings(year, String(prev));
    const players = wkRows.filter(r => r.user !== PARLAY_USER && (r.correct + r.incorrect + r.pending) > 0);
    if (!players.length || players.every(r => r.accuracy === null)) return null;

    const seasonRows = getStandings(year, "").filter(r => r.user !== PARLAY_USER);
    const seasonRank = {};
    seasonRows.forEach(r => {
      const better = seasonRows.filter(o => (o.accuracy === null ? -1 : o.accuracy) > (r.accuracy === null ? -1 : r.accuracy)).length;
      seasonRank[r.user] = better + 1;
    });

    const rows = players.map(r => {
      const better = players.filter(o => (o.accuracy === null ? -1 : o.accuracy) > (r.accuracy === null ? -1 : r.accuracy)).length;
      return {
        user: r.user, correct: r.correct, incorrect: r.incorrect, accuracy: r.accuracy,
        rank: better + 1, seasonRank: seasonRank[r.user] || null
      };
    });

    const pr = wkRows.find(r => r.user === PARLAY_USER);
    return {
      week: prev,
      rows: rows,
      parlay: pr && (pr.correct + pr.incorrect) > 0 ? { correct: pr.correct, incorrect: pr.incorrect, accuracy: pr.accuracy } : null,
      pending: Math.max.apply(null, players.map(r => r.pending || 0))
    };
  });
}

function getHomeParlayCard_(year, week) {
  return cachedLookup_("homeparlay", year + "|" + week, function () {
    const sheet = ensureParlayPlacementsSheet_();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let row = null;
    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(year) && String(data[i][1]) === String(week)) { row = data[i]; break; }
    }
    if (!row) return null;

    const tz = Session.getScriptTimeZone();
    const fmt = d => d instanceof Date ? Utilities.formatDate(d, tz, "EEE h:mm a") : String(d || "");
    const col = n => headers.indexOf(n);

    const runbackMadeBy = col("RunBackMadeBy") !== -1 ? row[col("RunBackMadeBy")] : "";
    if (runbackMadeBy) {
      let odds = null;
      try { odds = getRunbackStatus(null, year, week).runbackOdds; } catch (err) { odds = null; }
      return {
        type: "runback", week: week, placedBy: String(runbackMadeBy),
        placedAt: fmt(row[col("Runback MadeAt")]),
        participants: String(row[col("RunBackParticipants")] || ""),
        legs: odds ? odds.gamesIncluded : null,
        moneyline: odds ? odds.combinedMoneyline : null
      };
    }

    const placedBy = row[2];
    if (!placedBy) return null;
    let odds = null;
    try { odds = JSON.parse(getParlayWeekData(year, week)).parlayOdds; } catch (err) { odds = null; }
    return {
      type: "parlay", week: week, placedBy: String(placedBy), placedAt: fmt(row[3]),
      legs: odds ? odds.gamesIncluded : null,
      moneyline: odds ? odds.combinedMoneyline : null,
      stake: odds ? Number(odds.totalPayout) - Number(odds.totalProfit) : null,
      payout: odds ? Number(odds.totalPayout) : null
    };
  });
}

/** The two cards as Home-feed items (empty array when neither applies). */
function getHomeCards_(year, week) {
  const out = [];
  const nowIso = new Date().toISOString();
  try {
    const p = getHomeParlayCard_(year, week);
    if (p) out.push({ noticeId: "sys-parlay-" + year + "-" + week + "-" + p.type, kind: "parlay", postedBy: "", postedAt: nowIso, text: "", isNew: false, canDelete: false, parlay: p });
  } catch (err) { Logger.log("getHomeParlayCard_ failed: " + err.message); }
  try {
    const r = getHomeResultsCard_(year, week);
    if (r) out.push({ noticeId: "sys-results-" + year + "-" + r.week, kind: "results", postedBy: "", postedAt: nowIso, text: "", isNew: false, canDelete: false, results: r });
  } catch (err) { Logger.log("getHomeResultsCard_ failed: " + err.message); }
  return out;
}

// Posted notices as plain JSON-able objects, shared by every user's feed.
function getSharedNotices_() {
  return cachedLookup_("notices", "-", function () {
    const out = [];
    sheetToObjects_(SHEET_NOTICES).forEach(n => {
      const d = n.PostedAt instanceof Date ? n.PostedAt : new Date(n.PostedAt);
      if (isNaN(d.getTime())) return;
      out.push({ noticeId: String(n.NoticeID), postedBy: n.PostedBy, postedAt: d.toISOString(), text: n.Text });
    });
    return out;
  });
}

// Runs the runback safety-net check at most once per data version per 2 minutes.
function throttledRunbackCheck_(year, week) {
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (err) { cache = null; }
  if (cache) {
    try {
      const key = "rbchk:" + getDataVersion_() + ":" + year + ":" + week;
      if (cache.get(key)) return;
      cache.put(key, "1", 120);
    } catch (err) { /* cache trouble - just run the check */ }
  }
  try { checkAndNotifyRunbackStatus_(year, week); } catch (err) {
    Logger.log("checkAndNotifyRunbackStatus_ failed from getHomeFeed: " + err.message);
  }
}


// ==============================================================
// RUNBACK VOTE
// A genuinely separate second parlay bet on just the remaining
// games, offered to the group the moment the week's Thursday game
// (the earliest kickoff) goes final and the Parlay's consensus
// pick for it lost outright (not a push). Uses the SAME consensus
// picks already locked in for the remaining games - nothing about
// the original picks changes, this only decides whether a second
// real-money bet gets placed on them.
//
// Vote is Yes/No, same majority-decision style as the tie-break
// engine: decided the instant one side crosses a majority of all
// active users, or the vote is treated as failed once the next
// game (the first "remaining" game) kicks off, whichever is first.
// ==============================================================

function ensureRunbackVotesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_RUNBACK_VOTES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RUNBACK_VOTES);
    sheet.getRange(1, 1, 1, 5).setValues([["SeasonYear", "Week", "UserName", "Vote", "VotedAt"]]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    applyStandardFont_(sheet);
  }
  return sheet;
}

/**
 * Combines a date-only Sheets value with a time-only Sheets value
 * into one real Date, in the script's timezone. Sheets stores a
 * time-of-day cell as a Date on an arbitrary epoch date, so this
 * just borrows its hours/minutes onto the real event date.
 */
// The runback vote closes at the end of the day AFTER the first game (script
// time zone) - e.g. Thursday's game -> closes Friday 11:59 PM - or when the
// next game kicks off, whichever is sooner.
const RUNBACK_VOTE_CLOSE_HOUR_ = 23;
const RUNBACK_VOTE_CLOSE_MINUTE_ = 59;

function runbackVoteDeadline_(firstGame, nextGame) {
  const nextKick = combineDateAndTime_(nextGame.EventDate, nextGame.EventTime);
  const firstStart = combineDateAndTime_(firstGame.EventDate, firstGame.EventTime);
  if (!firstStart) return nextKick;
  const close = new Date(firstStart.getTime());
  close.setDate(close.getDate() + 1);
  close.setHours(RUNBACK_VOTE_CLOSE_HOUR_, RUNBACK_VOTE_CLOSE_MINUTE_, 0, 0);
  return (nextKick && nextKick < close) ? nextKick : close;
}

function combineDateAndTime_(dateVal, timeVal) {
  const d = dateVal ? new Date(dateVal) : null;
  if (!d || isNaN(d.getTime())) return null;
  if (timeVal instanceof Date && !isNaN(timeVal.getTime())) {
    d.setHours(timeVal.getHours(), timeVal.getMinutes(), 0, 0);
  }
  return d;
}

/**
 * Full status of the runback vote for a week: whether one is even
 * eligible to happen, whether it's currently open, decided, or the
 * window's closed, the running Yes/No tally, and a payout preview
 * for the remaining games priced off the SAME consensus picks
 * already locked in for them (same math as the live parlay-potential
 * banner, just starting after the first/Thursday leg).
 * userName is optional - pass null for a pure aggregate status
 * check (e.g. from the notification engine); pass a real user to
 * also get their own current vote back.
 */
function getRunbackStatus(userName, year, week) {
  const base = cachedLookup_("rbbase", year + "|" + week, function () { return computeRunbackBase_(year, week); });
  if (!base.eligible) return Object.assign({ status: "ok" }, base);

  ensureRunbackVotesSheet_();

  // Everything below depends on the live clock and the live votes, so it is
  // never cached - but it is only a few lines and one small sheet read.
  const deadline = base.deadlineMs ? new Date(base.deadlineMs) : null;
  const now = new Date();
  const windowClosed = !!(deadline && now >= deadline);

  const votesByUser = {};
  sheetToObjects_(SHEET_RUNBACK_VOTES)
    .filter(v => Number(v.SeasonYear) === Number(year) && String(v.Week) === String(week))
    .forEach(v => votesByUser[v.UserName] = v.Vote);

  const yesVoters = Object.keys(votesByUser).filter(u => votesByUser[u] === "Yes");
  const noVoters = Object.keys(votesByUser).filter(u => votesByUser[u] === "No");

  let decided = null;
  if (yesVoters.length >= base.decidedThreshold) decided = "Yes";
  else if (noVoters.length >= base.decidedThreshold) decided = "No";
  else if (windowClosed) decided = "No";

  return {
    status: "ok",
    eligible: true,
    open: !decided && !windowClosed,
    decided: decided,
    windowClosed: windowClosed,
    deadline: deadline ? Utilities.formatDate(deadline, Session.getScriptTimeZone(), "EEE h:mm a") : "",
    firstGameMatchup: base.firstGameMatchup,
    yesCount: yesVoters.length,
    noCount: noVoters.length,
    yesVoters: yesVoters,
    decidedThreshold: base.decidedThreshold,
    totalVoters: base.totalVoters,
    userVote: userName ? (votesByUser[userName] || null) : null,
    runbackOdds: base.runbackOdds
  };
}

/**
 * The part of the runback status that doesn't depend on the clock or on
 * anyone's vote: whether a runback is even eligible, the first game, the
 * deadline, the vote threshold, and the payout preview. Reads the schedule,
 * parlay picks, results and odds sheets - which is why it's cached.
 */
function computeRunbackBase_(year, week) {
  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r =>
      Number(r.SeasonYear) === Number(year) &&
      String(r.Week) === String(week) &&
      (r.RowType === "GAME" || r.RowType === "PRESEASON")
    )
    .sort((a, b) => new Date(a.EventDate) - new Date(b.EventDate));

  if (games.length < 2) {
    return { eligible: false };
  }

  const firstGame = games[0];
  const remainingGames = games.slice(1);

  const parlayByEvent = {};
  sheetToObjects_(SHEET_PARLAY)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week))
    .forEach(r => parlayByEvent[r.EventID] = r);

  const resultByEvent = {};
  sheetToObjects_(SHEET_RESULTS)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week))
    .forEach(r => resultByEvent[r.EventID] = r);

  const firstParlay = parlayByEvent[firstGame.EventID];
  const firstResult = resultByEvent[firstGame.EventID];
  const firstPicked = firstParlay && firstParlay.PickedTeam;
  const firstFinal = firstResult && firstResult.WinningTeam;

  const eligible = !!(firstFinal && firstFinal !== "TIE" && firstPicked && firstFinal !== firstPicked);
  if (!eligible) {
    return { eligible: false, firstGameMatchup: firstGame.Matchup };
  }

  const deadline = runbackVoteDeadline_(firstGame, remainingGames[0]);

  const users = getUsers();
  const decidedThreshold = Math.floor(users.length / 2) + 1;

  // Payout preview for the remaining games, using the picks already
  // locked in for them - same approach as the live parlay-potential
  // banner (getParlayWeekData's parlayOdds), just starting after leg 1.
  const oddsRows = sheetToObjects_(SHEET_ODDS)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));
  const latestMoneylineByEvent = {};
  oddsRows.forEach(r => {
    const existing = latestMoneylineByEvent[r.EventID];
    const rTime = r.PulledAt ? new Date(r.PulledAt).getTime() : 0;
    const existingTime = existing ? existing._pulledAtMs : 0;
    if (!existing || rTime >= existingTime) {
      latestMoneylineByEvent[r.EventID] = {
        spread: r.Spread || "",
        awayMoneyline: r.AwayMoneyline,
        homeMoneyline: r.HomeMoneyline,
        _pulledAtMs: rTime
      };
    }
  });

  let combinedDecimal = 1;
  let legsIncluded = 0;
  remainingGames.forEach(g => {
    const parlay = parlayByEvent[g.EventID];
    const pickedTeam = parlay && parlay.PickedTeam;
    if (!pickedTeam) return;

    const odds = latestMoneylineByEvent[g.EventID];
    let ml = null;
    if (odds && odds.awayMoneyline && odds.homeMoneyline) {
      const pickedIsAway = (pickedTeam === g.AwayShort);
      ml = pickedIsAway ? Number(odds.awayMoneyline) : Number(odds.homeMoneyline);
    } else if (odds && odds.spread) {
      const parsed = parseSpreadString_(odds.spread);
      if (parsed && parsed.points > 0) {
        const spreadResult = spreadToMoneyline_(parsed.points);
        const pickedIsFavorite = (pickedTeam.toUpperCase() === parsed.favTeam);
        ml = pickedIsFavorite ? spreadResult.favoriteML : spreadResult.underdogML;
      }
    }
    if (ml == null || isNaN(ml)) return;

    combinedDecimal *= moneylineToDecimal_(ml);
    legsIncluded++;
  });

  const totalStake = 1 * users.length; // $1 per person, same convention as the main parlay
  const runbackOdds = legsIncluded ? {
    combinedMoneyline: decimalToMoneyline_(combinedDecimal),
    totalPayout: totalStake * combinedDecimal,
    gamesIncluded: legsIncluded,
    gamesTotal: remainingGames.length
  } : null;

  return {
    eligible: true,
    firstGameMatchup: firstGame.Matchup,
    deadlineMs: deadline ? deadline.getTime() : null,
    decidedThreshold: decidedThreshold,
    totalVoters: users.length,
    runbackOdds: runbackOdds
  };
}

/**
 * Casts (or changes) userName's own Yes/No runback vote for the
 * week, then re-checks whether that decides it - which, if so,
 * fires the SMS + pre-fills the ParlayPlacements log via
 * checkAndNotifyRunbackStatus_ (defined in Automation.gs).
 */
function castRunbackVote(userName, year, week, vote) {
  if (!userName) return { status: "error", message: "Not logged in." };
  if (vote !== "Yes" && vote !== "No") return { status: "error", message: "Invalid vote." };

  const status = getRunbackStatus(userName, year, week);
  if (!status.eligible) return { status: "error", message: "No runback vote is open for this week." };
  if (status.decided) return { status: "error", message: "This runback vote is already decided." };

  const sheet = ensureRunbackVotesSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colYear = headers.indexOf("SeasonYear");
  const colWeek = headers.indexOf("Week");
  const colUser = headers.indexOf("UserName");
  const colVote = headers.indexOf("Vote");
  const colVotedAt = headers.indexOf("VotedAt");

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colYear]) === Number(year) && String(data[i][colWeek]) === String(week) && String(data[i][colUser]) === String(userName)) {
      rowIndex = i + 1;
      break;
    }
  }

  const now = new Date();
  if (rowIndex === -1) {
    const newRow = new Array(headers.length).fill("");
    newRow[colYear] = year;
    newRow[colWeek] = week;
    newRow[colUser] = userName;
    newRow[colVote] = vote;
    newRow[colVotedAt] = now;
    sheet.appendRow(newRow);
  } else {
    sheet.getRange(rowIndex, colVote + 1).setValue(vote);
    sheet.getRange(rowIndex, colVotedAt + 1).setValue(now);
  }

  try {
    checkAndNotifyRunbackStatus_(year, week);
  } catch (err) {
    Logger.log("checkAndNotifyRunbackStatus_ failed after castRunbackVote: " + err.message);
  }

  return { status: "ok", vote: vote };
}

/**
 * Once a runback vote passes, pre-fills the existing
 * ParlayPlacements Runback log (RunBack=Yes, RunBackParticipants=
 * the Yes voters) so the admin's existing "confirm it was actually
 * placed" step (recordRunbackPlacement) has a head start. Never
 * touches RunBackMadeBy/Runback MadeAt - that confirmation stays a
 * deliberate human step, same as before this feature existed.
 */
function recordRunbackVoteResultImpl_(year, week, yesVoters) {
  const sheet = ensureParlayPlacementsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colRunBack = headers.indexOf("RunBack");
  const colParticipants = headers.indexOf("RunBackParticipants");
  if (colRunBack === -1 || colParticipants === -1) return;

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(year) && String(data[i][1]) === String(week)) {
      rowIndex = i + 1;
      break;
    }
  }

  const participantsStr = yesVoters.join(", ");
  if (rowIndex === -1) {
    const newRow = new Array(headers.length).fill("");
    newRow[0] = year;
    newRow[1] = week;
    newRow[colRunBack] = "Yes";
    newRow[colParticipants] = participantsStr;
    sheet.appendRow(newRow);
  } else {
    sheet.getRange(rowIndex, colRunBack + 1).setValue("Yes");
    sheet.getRange(rowIndex, colParticipants + 1).setValue(participantsStr);
  }
}


// ------------------------------------------------------------
// ONE-TIME MIGRATION - adds the SendRunbackInfo notification
// preference column, same pattern/defaults as
// addNotificationPreferenceColumns(). Safe to re-run.
// ------------------------------------------------------------
function addRunbackNotificationColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("SendRunbackInfo") !== -1) {
    return { status: "ok", message: "SendRunbackInfo column already exists" };
  }

  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue("SendRunbackInfo").setFontWeight("bold");

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const defaults = [];
    for (let i = 0; i < lastRow - 1; i++) defaults.push([true]);
    sheet.getRange(2, newCol, lastRow - 1, 1).setValues(defaults);
  }

  return { status: "ok", message: "SendRunbackInfo column added - existing users defaulted to TRUE" };
}



// ==============================================================
// RESPONSE CACHE - makes the Games and Parlay screens fast
//
// The finished getWeekData / getParlayWeekData responses are kept
// in CacheService (Google's in-memory cache - not a sheet), so
// almost every load skips the slow part: reading and crunching
// several sheets. How it stays correct:
//
//   * One shared DATA VERSION lives in Script Properties. The cache
//     key includes it. Every function that changes what those
//     screens show bumps the version (see the wrappers below), so
//     the old cached copies are simply never looked up again.
//   * The pick window (open/closed, depends on today's date and on
//     admin override properties) is recomputed fresh on every
//     request, never served from cache.
//   * Entries expire after PAYLOAD_CACHE_TTL_SECONDS_ as a backstop
//     for changes that don't go through a wrapper (e.g. someone
//     hand-editing a sheet, or code in a file that isn't wrapped).
//     After a hand edit you can run bumpDataVersionNow() to force
//     everything to refresh immediately.
//   * If the cache is unavailable or an entry is too big (100KB
//     limit), it silently falls back to building the response the
//     normal way - it can only ever make things faster, not break.
//   * warmCurrentWeekCache (Automation.gs) rebuilds the current
//     week on a short timer, and score/odds pulls warm it right
//     after they finish, so nobody has to be the one who pays for
//     a rebuild.
// ==============================================================

const DATA_VERSION_PROP_ = "DATA_VERSION";
const PAYLOAD_CACHE_TTL_SECONDS_ = 600;      // 10 minutes
const PAYLOAD_CACHE_MAX_CHARS_ = 90000;      // CacheService limit is 100KB per value

function getDataVersion_() {
  return PropertiesService.getScriptProperties().getProperty(DATA_VERSION_PROP_) || "0";
}

/**
 * Marks "something the cached screens show has changed". The version is a
 * timestamp plus a random suffix rather than a counter, so two bumps can
 * never produce the same value (a counter could, if two writes raced).
 */
function bumpDataVersion_() {
  try {
    const next = Date.now().toString(36) + Math.floor(Math.random() * 1679616).toString(36); // 4 random base-36 chars
    PropertiesService.getScriptProperties().setProperty(DATA_VERSION_PROP_, next);
    return next;
  } catch (err) {
    Logger.log("bumpDataVersion_ failed: " + err.message);
    return null;
  }
}

// Run from the function dropdown after hand-editing a sheet, to make every
// cached screen refresh on its next load instead of waiting out the expiry.
function bumpDataVersionNow() {
  return { status: "ok", version: bumpDataVersion_() };
}

/**
 * Runs fn(...args) and bumps the data version afterwards - even if fn threw,
 * since it may have written part of what it was going to write before failing.
 */
function runThenBumpVersion_(fn, args) {
  try {
    return fn.apply(null, args);
  } finally {
    bumpDataVersion_();
  }
}

// computePickWindow_ only needs these three fields from each game.
function slimGamesForPickWindow_(games) {
  return games.map(g => ({ SeasonYear: g.SeasonYear, Week: g.Week, EventDate: g.EventDate }));
}

/**
 * Cache-through for a per-week payload. builder(year, week) must return
 * { payload: <plain object>, pwGames: <slim games for the pick window, or null> }.
 * Returns the JSON string the client expects. When withPickWindow is true,
 * pickWindow is computed fresh and added (it depends on today's date and admin
 * overrides, so it is never served from cache).
 */
function getCachedWeekPayload_(kind, year, week, builder, withPickWindow) {
  // Capture the version BEFORE reading any data: if a write lands while we're
  // building, the version moves and we refuse to store what may be stale.
  const startVersion = getDataVersion_();
  const key = "wp:" + kind + ":" + startVersion + ":" + year + ":" + week;

  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (err) { cache = null; }

  let entry = null;
  if (cache) {
    try {
      const hit = cache.get(key);
      if (hit) entry = JSON.parse(hit);
    } catch (err) {
      entry = null;
    }
  }

  if (!entry) {
    entry = builder(year, week);
    if (cache) {
      try {
        const serialized = JSON.stringify(entry);
        if (serialized.length <= PAYLOAD_CACHE_MAX_CHARS_ && getDataVersion_() === startVersion) {
          cache.put(key, serialized, PAYLOAD_CACHE_TTL_SECONDS_);
        }
      } catch (err) {
        Logger.log("Cache store skipped for " + key + ": " + err.message);
      }
    }
  }

  const out = Object.assign({}, entry.payload);
  if (withPickWindow) out.pickWindow = computePickWindow_(entry.pwGames);
  return JSON.stringify(out);
}

function getWeekData(year, week) {
  return getCachedWeekPayload_("week", year, week, buildWeekDataPayload_, true);
}

function getParlayWeekData(year, week) {
  return getCachedWeekPayload_("parlay", year, week, buildParlayWeekDataPayload_, false);
}

/**
 * Builds (and caches) both screens for one week. Never throws - it's called
 * from the end of score/odds pulls and from the background timer, and a
 * warm-up failure must not break either of those.
 */
function warmWeekCacheSafe_(year, week) {
  try {
    if (!year || !week || isNaN(Number(year))) return;
    getWeekData(year, week);
    getParlayWeekData(year, week);
  } catch (err) {
    Logger.log("warmWeekCacheSafe_ failed for " + year + "/" + week + ": " + err.message);
  }
  // The Standings tab: the week view (its default) and the whole-season view.
  try {
    getStandings(year, week);
    getStandings(year, "");
  } catch (err) {
    Logger.log("warm standings failed for " + year + "/" + week + ": " + err.message);
  }
  // Home's results table + Parlay card.
  try {
    getHomeCards_(year, week);
  } catch (err) {
    Logger.log("warm home cards failed for " + year + "/" + week + ": " + err.message);
  }
  // The Tie Break tab.
  try {
    getTieBreakWeekData(year, week);
  } catch (err) {
    Logger.log("warm tie break failed for " + year + "/" + week + ": " + err.message);
  }
}

/**
 * Called by the login screen the moment a name is picked, while the person
 * is still typing their phone number. Touches only information that is the
 * same for everyone (group notices, current week, pick and runback snapshots)
 * so that the real Home load right after login finds it already cached.
 * Returns nothing about any person, so it is safe to call before login.
 */
function warmHomeShared() {
  try {
    const year = getCurrentSeasonYear();
    const week = year ? getDefaultWeek(year) : null;
    getSharedNotices_();
    if (year && week) {
      getHomeCards_(year, week);
      hasUserPickedAllGamesThisWeek_("", year, week);
      getRunbackStatus(null, year, week);
    }
  } catch (err) {
    Logger.log("warmHomeShared failed: " + err.message);
  }
  return { status: "ok" };
}

/**
 * Diagnostic: times each screen's server call twice, so you can see the
 * cached speed-up directly. If the background warm-up is running, even the
 * first call is usually already a cache hit. Run from the function dropdown,
 * then check View > Logs / Executions.
 */
function DEBUG_checkCache() {
  const year = getCurrentSeasonYear();
  const week = year ? getDefaultWeek(year) : null;
  if (!year || !week) return { status: "ok", issue: "No current season/week found." };

  const timeIt = fn => {
    const start = Date.now();
    const out = fn();
    return { ms: Date.now() - start, chars: out.length };
  };

  const result = {
    status: "ok",
    dataVersion: getDataVersion_(),
    year: year,
    week: week,
    weekData_call1: timeIt(() => getWeekData(year, week)),
    weekData_call2: timeIt(() => getWeekData(year, week)),
    parlayData_call1: timeIt(() => getParlayWeekData(year, week)),
    parlayData_call2: timeIt(() => getParlayWeekData(year, week))
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


// Public names kept exactly as before (the client and any triggers call them
// by name); each just runs the original logic, then bumps the data version.
function submitWeekPicks(...args)                { return runThenBumpVersion_(submitWeekPicksImpl_, args); }
function submitPick(...args)                     { return runThenBumpVersion_(submitPickImpl_, args); }
function submitTieBreakVotes(...args)            { return runThenBumpVersion_(submitTieBreakVotesImpl_, args); }
function resolveTieByLeader(...args)             { return runThenBumpVersion_(resolveTieByLeaderImpl_, args); }
function recordParlayPlacement(...args)          { return runThenBumpVersion_(recordParlayPlacementImpl_, args); }
function recordRunbackPlacement(...args)         { return runThenBumpVersion_(recordRunbackPlacementImpl_, args); }
function generateAllParlayPicksAllYearsFast(...args) { return runThenBumpVersion_(generateAllParlayPicksAllYearsFastImpl_, args); }
function recordRunbackVoteResult_(...args)       { return runThenBumpVersion_(recordRunbackVoteResultImpl_, args); }

// Pulling final scores changes results and standings, so after bumping the
// version it also rebuilds that week's screens right away (this is the
// "warm the cache once results are in" step).
function pullFinalScores(...args) {
  const result = runThenBumpVersion_(pullFinalScoresImpl_, args);
  warmWeekCacheSafe_(args[0], args[1]);
  return result;
}

/**
 * Run from the editor: DEBUG_winningParlayBreakdown().
 * Scans weeks 1-18 of the year below, logs a one-line summary for every week
 * that has finished games, then prints the full per-game breakdown for the
 * week with the biggest Winning Parlay (or for FORCE_WEEK if you set it).
 */
function DEBUG_winningParlayBreakdown() {
  const year = 2026;      // <-- season to scan
  const FORCE_WEEK = null; // <-- e.g. 3 to break down one specific week instead
  let best = null;
  for (let w = 1; w <= 18; w++) {
    let bt;
    try { bt = getParlayMoneylineBacktest(year, w); } catch (e) { continue; }
    if (!bt || !bt.winnersIncluded) continue;
    const dec = bt.winnersPayout / Math.max(1, getUsers().length);
    Logger.log("Week " + w + ": " + bt.winnersIncluded + "/" + bt.legsTotal + " winners priced | " +
      bt.winnersCombinedMoneyline + " | payout $" + Math.round(bt.winnersPayout) + " | x" + dec.toFixed(1));
    if (!best || bt.winnersPayout > best.bt.winnersPayout) best = { w: w, bt: bt };
  }
  if (FORCE_WEEK) best = { w: FORCE_WEEK, bt: getParlayMoneylineBacktest(year, FORCE_WEEK) };
  if (!best) { Logger.log("No weeks with finished games found for " + year); return; }

  Logger.log("---- Breakdown: " + year + " week " + best.w + " ----");
  let running = 1;
  best.bt.legs.forEach(function (l, i) {
    if (l.winningTeamMl == null) {
      Logger.log((i + 1) + ". " + l.matchup + " | winner: " + (l.winningTeam || "n/a") + " | NOT INCLUDED");
      return;
    }
    const dec = moneylineToDecimal_(l.winningTeamMl);
    running *= dec;
    Logger.log((i + 1) + ". " + l.matchup + " | winner: " + l.winningTeam + " | ML " + l.winningTeamMl +
      " (" + l.winningTeamMlSource + ") | x" + dec.toFixed(3) + " | running x" + running.toFixed(1));
  });
  Logger.log("Combined decimal: " + running.toFixed(2) + " -> American: " + decimalToMoneyline_(running) +
    " | users: " + getUsers().length + " | payout $" + Math.round(getUsers().length * running));
}


// ==============================================================
// SCHEDULE-DERIVED LOOKUPS (cached)
// Current season, the week to land on, the week list with open/closed
// status, and the list of years. Each one used to re-read the whole
// Schedule sheet on every app open and every Home visit. They only change
// when the schedule is written (those writers bump the data version), when a
// pick-window override or test date changes (also bumps), or when the
// calendar day rolls over (the day is part of the cache key). The 5-minute
// expiry is a backstop for hand edits made directly in the Schedule sheet;
// run bumpDataVersionNow to see those immediately.
// ==============================================================

const LOOKUP_CACHE_TTL_SECONDS_ = 300;

function cachedLookup_(name, argKey, compute) {
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (err) { cache = null; }
  const startVersion = getDataVersion_();
  const dayIso = Utilities.formatDate(getEffectiveNow_(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const key = "lk:" + name + ":" + startVersion + ":" + dayIso + ":" + argKey;

  if (cache) {
    try {
      const hit = cache.get(key);
      if (hit) return JSON.parse(hit).v;
    } catch (err) { /* fall through to compute */ }
  }

  const value = compute();
  if (cache) {
    try {
      if (getDataVersion_() === startVersion) {
        cache.put(key, JSON.stringify({ v: value === undefined ? null : value }), LOOKUP_CACHE_TTL_SECONDS_);
      }
    } catch (err) { Logger.log("Lookup cache store skipped for " + key + ": " + err.message); }
  }
  return value;
}

function getCurrentSeasonYear()  { return cachedLookup_("year", "-", getCurrentSeasonYearImpl_); }
function getDefaultWeek(year)    { return cachedLookup_("defweek", String(year), function () { return getDefaultWeekImpl_(year); }); }
function getWeeksWithStatus(year){ return cachedLookup_("weeks", String(year), function () { return getWeeksWithStatusImpl_(year); }); }
function getYears()              { return cachedLookup_("years", "-", getYearsImpl_); }

// Standings only change when scores, picks, tie-break votes or the schedule
// change - all of which bump the data version - so they are cached per
// year + week the same way as the lookups above.
// The Tie Break screen's data only changes when picks, tie-break votes, a
// leader decision, results or the schedule change - all of which bump the
// data version.
function getTieBreakWeekData(year, week) {
  return cachedLookup_("tiebreak", String(year) + "|" + String(week), function () { return getTieBreakWeekDataImpl_(year, week); });
}

function getStandings(year, week) {
  return cachedLookup_("standings", String(year) + "|" + String(week || ""), function () { return getStandingsImpl_(year, week); });
}

/**
 * Run from the editor: DEBUG_cacheSizes().
 * Shows how big each cached item is for the current week against the
 * 100 KB-per-item limit of Google's cache. An item over the limit is simply
 * not cached (that screen just loads the slow way), so this tells you how
 * much headroom there is as the season goes on.
 */
function DEBUG_cacheSizes() {
  const year = getCurrentSeasonYear();
  const week = year ? getDefaultWeek(year) : null;
  if (!year || !week) { Logger.log("No current season/week found."); return; }

  const LIMIT = 100 * 1024;
  const items = [
    ["Games screen (getWeekData)", getWeekData(year, week)],
    ["Parlay screen (getParlayWeekData)", getParlayWeekData(year, week)],
    ["Standings, this week", JSON.stringify(getStandings(year, week))],
    ["Standings, All Weeks", JSON.stringify(getStandings(year, ""))],
    ["Tie Break screen", getTieBreakWeekData(year, week)],
    ["Home notices (shared)", JSON.stringify(getSharedNotices_())],
    ["Week list", JSON.stringify(getWeeksWithStatus(year))]
  ];

  let total = 0;
  items.forEach(function (it) {
    const bytes = Utilities.newBlob(it[1]).getBytes().length; // real size, accents/emoji included
    total += bytes;
    Logger.log(it[0] + ": " + (bytes / 1024).toFixed(1) + " KB (" + Math.round(bytes / LIMIT * 100) + "% of the 100 KB limit)" +
      (bytes > LIMIT ? "  <-- TOO BIG, not cached" : ""));
  });
  Logger.log("All of the above together: " + (total / 1024).toFixed(1) + " KB across " + items.length + " items");
}


// ------------------------------------------------------------
// ONE-TIME MIGRATION - adds the SendParlayInfo notification
// preference column ("week is open" and "Parlay placed" texts),
// same pattern/defaults as addRunbackNotificationColumn().
// Safe to re-run. Until it's run, everyone simply gets these
// texts (the preference defaults to on) but can't opt out.
// ------------------------------------------------------------
function addParlayNotificationColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: "error", message: "Users sheet not found" };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf("SendParlayInfo") !== -1) {
    return { status: "ok", message: "SendParlayInfo column already exists" };
  }

  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue("SendParlayInfo").setFontWeight("bold");

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const defaults = [];
    for (let i = 0; i < lastRow - 1; i++) defaults.push([true]);
    sheet.getRange(2, newCol, lastRow - 1, 1).setValues(defaults);
  }

  return { status: "ok", message: "SendParlayInfo column added - existing users defaulted to TRUE" };
}
