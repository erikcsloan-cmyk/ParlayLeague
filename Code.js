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
    .setTitle("NFL Pick 'Em")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------------------------
// SETUP - creates the Users / Picks / Results sheets if missing
// Safe to call repeatedly - only creates what's missing.
// ------------------------------------------------------------
function ensureSupportSheets() {
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

/**
 * UserName -> DisplayName lookup for the client. Falls back to the
 * UserName itself wherever DisplayName is blank, so this is safe
 * to use everywhere without null-checking on the client.
 */
function getUserDisplayNames_() {
  const users = sheetToObjects_(SHEET_USERS);
  const map = {};
  users.forEach(u => {
    const nick = (u.DisplayName && String(u.DisplayName).trim()) ? String(u.DisplayName).trim() : u.UserName;
    map[u.UserName] = nick;
  });
  return map;
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

  // Tie-break re-votes, only consulted for games where the main vote ties.
  // Track WHO voted per event (not just the tally) - a tie-break only
  // counts as resolved once every active user has actually voted, not
  // just whenever whoever's voted so far happens to agree.
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

  // Leader overrides - highest priority. If the season standings leader
  // has made a call on a still-tied game, that decision always wins,
  // regardless of what any vote tally says.
  const leaderOverrides = sheetToObjects_(SHEET_LEADER_OVERRIDE)
    .filter(p => Number(p.SeasonYear) === Number(year) && String(p.Week) === String(week) && p.PickedTeam);
  const leaderOverrideByEvent = {};
  leaderOverrides.forEach(p => leaderOverrideByEvent[p.EventID] = p.PickedTeam);

  // Picks the top team from a vote map; returns tied:true if there's no
  // single majority (two or more teams sharing the top count)
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

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PARLAY);
  const existing = sheet.getDataRange().getValues();
  const headers = existing[0];
  const colEventId = headers.indexOf("EventID");

  const existingRowByEvent = {};
  for (let i = 1; i < existing.length; i++) {
    existingRowByEvent[existing[i][colEventId]] = i + 1;
  }

  games.forEach(g => {
    const voteMap = votesByEvent[g.EventID] || {};
    const teams = Object.keys(voteMap);

    let pickedTeam = "";
    let voteCount = 0;
    const totalVotes = teams.reduce((sum, t) => sum + voteMap[t], 0);

    if (teams.length) {
      const main = resolveVoteMap_(voteMap);
      pickedTeam = main.pickedTeam;
      voteCount = main.voteCount;

      // Main vote is tied - a leader override always wins first if one exists
      if (main.tied && leaderOverrideByEvent[g.EventID]) {
        pickedTeam = leaderOverrideByEvent[g.EventID];
      } else if (main.tied) {
        // Otherwise, only let a tie-break re-vote resolve it once
        // EVERY active user has actually voted on this specific game
        const votersSet = tieBreakVotersByEvent[g.EventID] || new Set();
        const everyoneVoted = activeUsers.length > 0 && activeUsers.every(u => votersSet.has(u));

        if (everyoneVoted) {
          const tieBreak = resolveVoteMap_(tieBreakVotesByEvent[g.EventID]);
          if (tieBreak.pickedTeam) {
            pickedTeam = tieBreak.pickedTeam;
            // voteCount/totalVotes stay as the original split for context;
            // pickedTeam being set is what signals "resolved"
          }
          // if tieBreak itself ties even with everyone voted, pickedTeam
          // correctly stays "" - genuinely still tied, needs another round
        }
        // if not everyone's voted yet, pickedTeam stays "" regardless of
        // how the partial tie-break votes so far happen to lean
      }
    }

    const row = [year, week, g.EventID, g.Matchup, pickedTeam, voteCount, totalVotes, new Date()];

    if (existingRowByEvent[g.EventID]) {
      sheet.getRange(existingRowByEvent[g.EventID], 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  });

  return { status: "ok", gamesProcessed: games.length };
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
function getCurrentSeasonYear() {
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
function getWeeksWithStatus(year) {
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
  return { status: "ok", testDate: dateString };
}

function clearTestDate() {
  PropertiesService.getScriptProperties().deleteProperty("TEST_OVERRIDE_DATE");
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
function getDefaultWeek(year) {
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

    const endDate = new Date(win.tuesdayIso + "T00:00:00");
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
  return { status: "ok", year: year, week: week, newClosingDate: newClosingDateString };
}

/**
 * Removes a manual extension for a specific week, reverting it to
 * the normal Tue-Wed rule.
 */
function clearPickWindowExtension(year, week) {
  PropertiesService.getScriptProperties().deleteProperty("PICK_WINDOW_EXTEND_" + year + "_" + week);
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
  return { status: "ok", year: year, week: week, newOpeningDate: newOpeningDateString };
}

/**
 * Removes a manual early-open override for a specific week,
 * reverting it to the normal Tuesday-open rule.
 */
function clearPickWindowEarlyOpen(year, week) {
  PropertiesService.getScriptProperties().deleteProperty("PICK_WINDOW_OPEN_EARLY_" + year + "_" + week);
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
  return {
    currentYear: currentYear,
    defaultWeek: defaultWeek,
    weeks: currentYear ? getWeeksWithStatus(currentYear) : [],
    users: getUsers(),
    allYears: getYears(),
    displayNames: getUserDisplayNames_()
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

  return { status: "ok", userName: userName };
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

function getYears() {
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
  return sheetToObjects_(SHEET_USERS)
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
  wednesday.setDate(wednesday.getDate() + 1);

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

  return {
    tuesday: fmt(effectiveTuesday),
    wednesday: fmt(effectiveWednesday),
    tuesdayIso: tuesdayIso,
    wednesdayIso: wednesdayIso,
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
function getLatestOddsByEvent_(year, week) {
  const rows = sheetToObjects_(SHEET_ODDS)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));

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

/**
 * Returns everything the UI needs to render one week:
 * games, each user's pick for each game, and final scores if available.
 */
function getWeekData(year, week) {
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
  const oddsByEvent = getLatestOddsByEvent_(year, week);

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
    const isInternational = !!state && !US_STATES.has(state);

    const odds = oddsByEvent[g.EventID] || null;

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
      spread: odds ? odds.spread : "",
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

  return JSON.stringify({ games: gamesOut, users: users, pickWindow: computePickWindow_(games) });
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
function resolveTieByLeader(year, week, eventId, matchup, pickedTeam, requestingUser) {
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
function getTieBreakWeekData(year, week) {
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
function submitTieBreakVotes(year, week, userName, votesArray) {
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

function getParlayWeekData(year, week) {
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
  const oddsByEvent = getLatestOddsByEvent_(year, week);

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
    const isInternational = !!state && !US_STATES.has(state);

    const parlayInfo = (parlay && Number(parlay.TotalVotes) > 0) ? {
      pickedTeam: parlay.PickedTeam || "",
      voteCount: parlay.VoteCount,
      totalVotes: parlay.TotalVotes
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

  return JSON.stringify({ games: gamesOut, users: users, pickWindow: computePickWindow_(games) });
}

/**
 * Save/update ALL of one user's picks for a week in one call, used by
 * the "Submit Picks" button. picksArray: [{eventId, matchup, pickedTeam}, ...]
 * Entries with an empty pickedTeam are skipped (no accidental overwrite).
 */
function submitWeekPicks(year, week, userName, picksArray) {
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

function submitPick(year, week, eventId, matchup, userName, pickedTeam) {
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
function pullFinalScores(year, week) {
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

  return { status: "ok", updated: updated, skippedIncomplete: skippedIncomplete };
}

/**
 * Standings for a season: Correct / Incorrect / Pending counts, plus
 * % Accurate computed as the AVERAGE of each week's accuracy rate
 * (correct / decided that week), not a flat total-correct/total-picks
 * ratio. A user who goes 3/3 one week and 0/5 the next sits at 50%,
 * not 30%. Weeks with no decided games yet don't count toward the average.
 */
function getStandings(year) {
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
      String(p.Week).charAt(0) !== "P" // preseason doesn't count toward real standings
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

  picks.forEach(p => tallyPick(p.UserName, String(p.Week), p.EventID, p.PickedTeam));

  // Parlay is scored from the computed ParlayPicks table, not from Picks
  const parlayRows = sheetToObjects_(SHEET_PARLAY)
    .filter(r => Number(r.SeasonYear) === Number(year) && r.PickedTeam && String(r.Week).charAt(0) !== "P");
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

function generateAllParlayPicksAllYearsFast() {
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
      const totalVotes = teams.reduce((sum, t) => sum + voteMap[t], 0);

      if (teams.length) {
        const main = resolveVoteMap_(voteMap);
        pickedTeam = main.pickedTeam;
        voteCount = main.voteCount;

        if (main.tied && leaderOverrideByKey[key]) {
          pickedTeam = leaderOverrideByKey[key];
        } else if (main.tied) {
          const votersSet = tieBreakVotersByKey[key] || new Set();
          const everyoneVoted = activeUsers.length > 0 && activeUsers.every(u => votersSet.has(u));

          if (everyoneVoted) {
            const tieBreak = resolveVoteMap_(tieBreakVotesByKey[key] || {});
            if (tieBreak.pickedTeam) {
              pickedTeam = tieBreak.pickedTeam;
            }
          }
        }
      }

      freshRows.push([g.SeasonYear, g.Week, g.EventID, g.Matchup, pickedTeam, voteCount, totalVotes, now]);
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
// ------------------------------------------------------------
// DEBUG - clears the stale notification flags for 2026 Week P3,
// so the next pick submission re-evaluates completeness and tie
// status fresh. Paste in, select DEBUG_clearP3Flags from the
// function dropdown, and run. Safe to delete afterward.
// ------------------------------------------------------------
function DEBUG_clearP3Flags() {
  return clearWeekNotifiedFlag(2026, "P3");
}
// ------------------------------------------------------------
// DEBUG - reverts Week 2's early-open (set by mistake via "Open
// Next Week Today") back to its normal Tuesday open, and opens
// Week 1 early instead, starting today. Safe to delete afterward.
// ------------------------------------------------------------
function DEBUG_fixWeek1EarlyOpen() {
  const year = 2026;
  const tz = Session.getScriptTimeZone();
  const todayIso = Utilities.formatDate(getEffectiveNow_(), tz, "yyyy-MM-dd");

  const cleared = clearPickWindowEarlyOpen(year, "2");
  const opened = openPickWindowEarly(year, "1", todayIso);

  Logger.log("Cleared Week 2 early-open: " + JSON.stringify(cleared));
  Logger.log("Opened Week 1 early starting " + todayIso + ": " + JSON.stringify(opened));

  return { week2Cleared: cleared, week1Opened: opened };
}
