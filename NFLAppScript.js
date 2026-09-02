// ======================================================
// NFL APP SCRIPT - CLEANED VERSION
// ======================================================
//
// This replaces NFLAppScript.gs entirely. It removes:
//   - pullNFLRegularSeasonToSheet() and addNFLByeWeeks_()
//     (the old hardcoded-year, full-sheet-clearing pair that
//     caused the schema corruption)
//   - the duplicate second copy of pullNFLHistoricalResultsByDate()
//   - pullNFLScheduleForYear() / pullNFLScheduleByDateForYear()
//     and their wrappers (both full-year-replace approaches,
//     superseded by the non-destructive upsert below)
//
// Kept / new:
//   1) pullNFLScheduleUpsert(year)  <- use this for schedule pulls
//   2) addNFLByeWeeksForYear_()      <- helper, current schema
//   3) snapshotNFLCurrentOdds()      <- unchanged, still needed
//   4) pullNFLHistoricalResultsByDate() <- single copy, unchanged
//   5) countScheduleRowsByYearAndType() <- diagnostic, kept
//   6) resetScheduleHeadersCompletely() <- emergency reset tool, kept
// ======================================================


// ======================================================
// 1. SCHEDULE PULL - UPSERT BY EVENTID (no overwrites)
//    Target: NFL_Schedule
//
// Fetches a season day-by-day (reliable for both current and
// historical seasons - ESPN's week-based endpoint is not
// trustworthy for completed seasons). For each game:
//   - if a row with that EventID already exists for this year,
//     its cells are updated in place
//   - otherwise a new row is appended
// Nothing is ever deleted. Other years are never touched.
// Refuses to run if NFL_Schedule's headers don't match the
// expected set, rather than risk writing into the wrong columns.
// ======================================================

function pullNFLScheduleUpsert(year) {
  const sheetName = "NFL_Schedule";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const expectedHeaders = [
    "SeasonYear", "Week", "EventID", "RowType", "EventDate", "EventTime",
    "AwayTeam", "AwayShort", "HomeTeam", "HomeShort", "Matchup",
    "Venue", "City", "State"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length)
      .setValues([expectedHeaders])
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
    applyStandardFont_(sheet);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = expectedHeaders.filter(h => headers.indexOf(h) === -1);

  if (missing.length) {
    const msg = "NFL_Schedule is missing expected column(s): " + missing.join(", ") +
      ". Run resetScheduleHeadersCompletely() first, then re-pull each year.";
    Logger.log(msg);
    return { status: "error", message: msg };
  }

  const colMap = {};
  expectedHeaders.forEach(h => colMap[h] = headers.indexOf(h));

  const dateRanges = {
    2024: { start: new Date("2024-09-01"), end: new Date("2025-01-10") },
    2025: { start: new Date("2025-09-01"), end: new Date("2026-01-10") },
    2026: { start: new Date("2026-09-01"), end: new Date("2027-01-10") }
  };

  const range = dateRanges[year];
  if (!range) {
    const msg = "No date range configured for " + year + " - add one to dateRanges in this function.";
    Logger.log(msg);
    return { status: "error", message: msg };
  }

  Logger.log("================================");
  Logger.log("SCHEDULE UPSERT - SEASON " + year);
  Logger.log("================================");

  const recordsByEventId = {};
  const currentDate = new Date(range.start);

  while (currentDate <= range.end) {
    const espnDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyyMMdd");
    const url =
      "https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard" +
      "?dates=" + espnDate;

    try {
      const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });

      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        const events = data.events || [];

        if (events.length) {
          Logger.log(espnDate + ": " + events.length + " game(s)");
        }

        events.forEach(event => {
          const eventSeason = Number((event.season && event.season.year) || (data.season && data.season.year) || 0);
          if (eventSeason && eventSeason !== year) return;

          const seasonType = Number((event.season && event.season.type) || (data.season && data.season.type) || 0);
          if (seasonType && seasonType !== 2) return; // regular season only

          const competition = (event.competitions && event.competitions[0]) || {};
          const competitors = competition.competitors || [];
          const away = competitors.find(c => c.homeAway === "away") || {};
          const home = competitors.find(c => c.homeAway === "home") || {};

          const awayShort = (away.team && away.team.abbreviation) || "";
          const homeShort = (home.team && home.team.abbreviation) || "";
          if (!awayShort || !homeShort) return;

          const awayFull = (away.team && (away.team.displayName || away.team.shortDisplayName)) || awayShort;
          const homeFull = (home.team && (home.team.displayName || home.team.shortDisplayName)) || homeShort;

          const eventDate = event.date ? new Date(event.date) : null;
          const venue = competition.venue || {};
          const address = venue.address || {};
          const week = (event.week && event.week.number) || (data.week && data.week.number) || "";

          recordsByEventId[String(event.id)] = {
            SeasonYear: year,
            Week: week,
            EventID: event.id,
            RowType: "GAME",
            EventDate: eventDate ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
            EventTime: eventDate ? Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "h:mm a") : "",
            AwayTeam: awayFull,
            AwayShort: awayShort,
            HomeTeam: homeFull,
            HomeShort: homeShort,
            Matchup: awayShort + " @ " + homeShort,
            Venue: venue.fullName || "",
            City: address.city || "",
            State: address.state || ""
          };
        });
      }
    } catch (err) {
      Logger.log("ERROR " + espnDate + ": " + err.message);
    }

    currentDate.setDate(currentDate.getDate() + 1);
    Utilities.sleep(75);
  }

  const gameRecords = Object.values(recordsByEventId);
  Logger.log("GAME records fetched: " + gameRecords.length);

  const byeRecords = addNFLByeWeeksForYear_(gameRecords, year, 1, 18);
  Logger.log("BYE records computed: " + byeRecords.length);

  const allRecords = gameRecords.concat(byeRecords);

  // Map existing EventIDs (for THIS year only) to their sheet row number
  const existingData = sheet.getDataRange().getValues();
  const existingRowByEventId = {};
  for (let i = 1; i < existingData.length; i++) {
    if (Number(existingData[i][colMap.SeasonYear]) !== Number(year)) continue;
    existingRowByEventId[String(existingData[i][colMap.EventID])] = i + 1;
  }

  let updated = 0;
  const newRows = [];

  allRecords.forEach(record => {
    const rowArray = expectedHeaders.map(h => (record[h] !== undefined ? record[h] : ""));
    const existingRow = existingRowByEventId[String(record.EventID)];

    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, expectedHeaders.length).setValues([rowArray]);
      updated++;
    } else {
      newRows.push(rowArray);
    }
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, expectedHeaders.length).setValues(newRows);
  }

  // Keep EventID as plain numbers so Sheets can't reinterpret them as dates
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, colMap.EventID + 1, sheet.getLastRow() - 1, 1).setNumberFormat("0");
  }

  // Light re-sort for readability - doesn't delete or move data destructively
  if (sheet.getLastRow() > 1) {
    const sortSpec = [
      { column: colMap.SeasonYear + 1, ascending: true },
      { column: colMap.Week + 1, ascending: true },
      { column: colMap.EventDate + 1, ascending: true },
      { column: colMap.RowType + 1, ascending: true }
    ];
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).sort(sortSpec);
  }

  Logger.log("================================");
  Logger.log("DONE - " + year + ": " + updated + " rows updated in place, " + newRows.length + " new rows inserted");
  Logger.log("================================");

  return { status: "ok", year: year, updated: updated, inserted: newRows.length };
}

// Convenience wrappers - run these directly from the function dropdown
function pullNFLScheduleUpsertFor2024() {
  return pullNFLScheduleUpsert(2024);
}

function pullNFLScheduleUpsertFor2025() {
  return pullNFLScheduleUpsert(2025);
}

function pullNFLScheduleUpsertFor2026() {
  return pullNFLScheduleUpsert(2026);
}


// ======================================================
// 2. BYE-WEEK HELPER (current schema - AwayShort/HomeShort)
// ======================================================

function addNFLByeWeeksForYear_(gameRecords, season, firstWeek, lastWeek) {
  const nflTeams = [
    "ARI", "ATL", "BAL", "BUF", "CAR", "CHI", "CIN", "CLE",
    "DAL", "DEN", "DET", "GB", "HOU", "IND", "JAX", "KC",
    "LV", "LAC", "LAR", "MIA", "MIN", "NE", "NO", "NYG",
    "NYJ", "PHI", "PIT", "SF", "SEA", "TB", "TEN", "WSH"
  ];

  const byeRecords = [];

  for (let week = firstWeek; week <= lastWeek; week++) {
    const teamsPlaying = new Set();

    gameRecords.forEach(record => {
      if (Number(record.Week) !== week) return;
      if (record.RowType !== "GAME") return;
      if (record.AwayShort) teamsPlaying.add(String(record.AwayShort).trim());
      if (record.HomeShort) teamsPlaying.add(String(record.HomeShort).trim());
    });

    if (teamsPlaying.size === 0) {
      Logger.log("Week " + week + ": no games found - skipping bye calculation");
      continue;
    }

    const byeTeams = nflTeams.filter(t => !teamsPlaying.has(t));

    byeTeams.forEach(team => {
      byeRecords.push({
        SeasonYear: season,
        Week: week,
        EventID: "BYE-" + season + "-" + week + "-" + team,
        RowType: "BYE",
        EventDate: "",
        EventTime: "",
        AwayTeam: team,
        AwayShort: team,
        HomeTeam: "",
        HomeShort: "",
        Matchup: team + " \u2014 BYE",
        Venue: "",
        City: "",
        State: ""
      });
    });
  }

  return byeRecords;
}


// ======================================================
// 3. SNAPSHOT CURRENT NFL ODDS (unchanged)
//    Target: API_NFL_OddsHistory
// ======================================================

function snapshotNFLCurrentOdds() {
  const season = 2026;
  const firstWeek = 1;
  const lastWeek = 18;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "API_NFL_OddsHistory";

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const headers = [
    "OddsSnapshotID", "EventID", "SeasonYear", "Week", "Matchup",
    "GameDateTime", "PulledAt", "OddsProvider", "Spread", "OverUnder",
    "AwayMoneyline", "HomeMoneyline", "IsOpeningSnapshot", "IsClosingSnapshot"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    applyStandardFont_(sheet);
  }

  const pulledAt = new Date();
  const records = [];

  for (let week = firstWeek; week <= lastWeek; week++) {
    const url =
      "https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard" +
      "?season=" + season + "&seasontype=2&week=" + week;

    try {
      const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        Logger.log("NFL odds Week " + week + " failed: HTTP " + response.getResponseCode());
        continue;
      }

      const data = JSON.parse(response.getContentText());

      (data.events || []).forEach(event => {
        const competition = (event.competitions && event.competitions[0]) || {};
        const competitors = competition.competitors || [];
        const away = competitors.find(c => c.homeAway === "away") || {};
        const home = competitors.find(c => c.homeAway === "home") || {};

        const awayTeam = (away.team && (away.team.abbreviation || away.team.displayName)) || "";
        const homeTeam = (home.team && (home.team.abbreviation || home.team.displayName)) || "";
        const matchup = awayTeam + " @ " + homeTeam;

        const odds = (competition.odds && competition.odds[0]) || {};
        if (!odds.details && odds.overUnder == null && !odds.awayTeamOdds && !odds.homeTeamOdds) return;

        const awayOdds = odds.awayTeamOdds || {};
        const homeOdds = odds.homeTeamOdds || {};
        const provider = odds.provider || {};
        const snapshotId = event.id + "-" + Utilities.formatDate(pulledAt, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");

        records.push({
          OddsSnapshotID: snapshotId,
          EventID: event.id || "",
          SeasonYear: season,
          Week: (event.week && event.week.number) || (data.week && data.week.number) || week,
          Matchup: matchup,
          GameDateTime: event.date ? new Date(event.date) : "",
          PulledAt: pulledAt,
          OddsProvider: provider.name || provider.displayName || provider.nameDisplay || "",
          Spread: odds.details || "",
          OverUnder: odds.overUnder ?? "",
          AwayMoneyline: awayOdds.moneyLine ?? "",
          HomeMoneyline: homeOdds.moneyLine ?? "",
          IsOpeningSnapshot: false,
          IsClosingSnapshot: false
        });
      });

      Utilities.sleep(150);
    } catch (err) {
      Logger.log("NFL odds Week " + week + " error: " + err.message);
    }
  }

  const rows = records.map(record => headers.map(h => record[h] ?? ""));

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  Logger.log("NFL odds snapshots added: " + rows.length);
}


// ======================================================
// 4. HISTORICAL NFL RESULTS - single copy (unchanged logic)
//    Target: API_NFL_Results
// ======================================================

function pullNFLHistoricalResultsByDate() {
  const seasons = [
    { season: 2024, startDate: new Date("2024-09-05"), endDate: new Date("2025-01-05") },
    { season: 2025, startDate: new Date("2025-09-04"), endDate: new Date("2026-01-04") }
  ];

  const sheetName = "API_NFL_Results";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const headers = ["SeasonYear", "Week", "EventID", "Matchup", "AwayScore", "HomeScore", "WinningTeam", "UpdatedAt"];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  applyStandardFont_(sheet);

  const recordsByEventId = {};

  seasons.forEach(config => {
    const season = config.season;
    const currentDate = new Date(config.startDate);

    Logger.log("================================");
    Logger.log("NFL HISTORICAL SEASON: " + season);
    Logger.log("================================");

    while (currentDate <= config.endDate) {
      const espnDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyyMMdd");
      const url = "https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard?dates=" + espnDate;

      try {
        const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
        const code = response.getResponseCode();

        if (code !== 200) {
          currentDate.setDate(currentDate.getDate() + 1);
          continue;
        }

        const data = JSON.parse(response.getContentText());
        const events = data.events || [];

        events.forEach(event => {
          const competition = (event.competitions && event.competitions[0]) || {};
          const competitors = competition.competitors || [];
          const away = competitors.find(c => c.homeAway === "away") || {};
          const home = competitors.find(c => c.homeAway === "home") || {};

          const eventSeason = Number((event.season && event.season.year) || (data.season && data.season.year) || 0);
          if (eventSeason && eventSeason !== season) return;

          const seasonType = Number((event.season && event.season.type) || (data.season && data.season.type) || 0);
          if (seasonType && seasonType !== 2) return;

          const awayTeam = (away.team && (away.team.abbreviation || away.team.shortDisplayName || away.team.displayName)) || "";
          const homeTeam = (home.team && (home.team.abbreviation || home.team.shortDisplayName || home.team.displayName)) || "";
          if (!awayTeam || !homeTeam) return;

          const matchup = awayTeam + " @ " + homeTeam;

          const awayScore = (away.score !== undefined && away.score !== null && away.score !== "") ? Number(away.score) : "";
          const homeScore = (home.score !== undefined && home.score !== null && home.score !== "") ? Number(home.score) : "";

          const completed = (competition.status && competition.status.type && competition.status.type.completed === true) ||
            (event.status && event.status.type && event.status.type.completed === true);

          let winningTeam = "";
          if (completed && awayScore !== "" && homeScore !== "") {
            if (awayScore > homeScore) winningTeam = awayTeam;
            else if (homeScore > awayScore) winningTeam = homeTeam;
            else winningTeam = "TIE";
          }

          const week = (event.week && event.week.number) || (data.week && data.week.number) || "";

          const record = {
            SeasonYear: season, Week: week, EventID: event.id || "", Matchup: matchup,
            AwayScore: awayScore, HomeScore: homeScore, WinningTeam: winningTeam, UpdatedAt: new Date()
          };

          if (record.EventID) {
            recordsByEventId[String(record.EventID)] = record;
          }
        });
      } catch (err) {
        Logger.log("ERROR " + season + " " + espnDate + ": " + err.message);
      }

      currentDate.setDate(currentDate.getDate() + 1);
      Utilities.sleep(75);
    }
  });

  const records = Object.values(recordsByEventId);

  records.sort((a, b) => {
    const seasonDiff = Number(a.SeasonYear) - Number(b.SeasonYear);
    if (seasonDiff !== 0) return seasonDiff;
    const weekDiff = Number(a.Week) - Number(b.Week);
    if (weekDiff !== 0) return weekDiff;
    return String(a.EventID).localeCompare(String(b.EventID));
  });

  const rows = records.map(record => headers.map(h => record[h] ?? ""));

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log("================================");
  Logger.log("NFL HISTORICAL RESULTS COMPLETE - TOTAL GAMES WRITTEN: " + records.length);
  Logger.log("================================");
}


// ======================================================
// 5. DIAGNOSTIC - row counts by year and RowType
// ======================================================

function countScheduleRowsByYearAndType() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NFL_Schedule");
  if (!sheet) {
    Logger.log("NFL_Schedule sheet not found under that exact name.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colYear = headers.indexOf("SeasonYear");
  const colRowType = headers.indexOf("RowType");

  Logger.log("Total rows in sheet (including header): " + data.length);
  Logger.log("Headers found: " + headers.join(", "));

  if (colYear === -1 || colRowType === -1) {
    Logger.log("Could not find SeasonYear or RowType column - check header names above.");
    return;
  }

  const counts = {};
  for (let i = 1; i < data.length; i++) {
    const year = data[i][colYear];
    const rowType = data[i][colRowType];
    const key = year || "(blank year)";
    if (!counts[key]) counts[key] = { GAME: 0, BYE: 0, other: 0 };
    if (rowType === "GAME") counts[key].GAME++;
    else if (rowType === "BYE") counts[key].BYE++;
    else counts[key].other++;
  }

  Logger.log("=== Row breakdown by SeasonYear and RowType ===");
  Object.keys(counts).sort().forEach(year => {
    const c = counts[year];
    Logger.log("  " + year + ": GAME=" + c.GAME + ", BYE=" + c.BYE + ", other/blank=" + c.other);
  });
}


// ======================================================
// 6. EMERGENCY RESET - only use if headers are ever wrong
//    again. Wipes NFL_Schedule entirely and writes the
//    correct header row. Then re-run pullNFLScheduleUpsert
//    for each year to rebuild.
// ======================================================

function resetScheduleHeadersCompletely() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("NFL_Schedule");

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet("NFL_Schedule");

  const headers = [
    "SeasonYear", "Week", "EventID", "RowType", "EventDate", "EventTime",
    "AwayTeam", "AwayShort", "HomeTeam", "HomeShort", "Matchup",
    "Venue", "City", "State"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.setFrozenRows(1);

  Logger.log("NFL_Schedule reset - correct headers written, all data cleared.");
  Logger.log("Next: run pullNFLScheduleUpsertFor2024/2025/2026 to rebuild.");

  return { status: "ok", message: "Schedule reset - rebuild each year now" };
}
// ======================================================
// 7. REMOVE PRESEASON - deletes PRESEASON-tagged rows for
//    a given year from NFL_Schedule. GAME and BYE rows, and
//    all other years, are never touched. Deletes bottom-up
//    in merged contiguous chunks so row indices never shift
//    out from under an in-progress delete.
// ======================================================

function removePreseasonRowsFromSchedule(year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NFL_Schedule");
  if (!sheet) {
    return { status: "error", message: "NFL_Schedule sheet not found.", removed: 0 };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { status: "ok", removed: 0 };
  }

  const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = data[0];
  const colYear = headers.indexOf("SeasonYear");
  const colRowType = headers.indexOf("RowType");

  if (colYear === -1 || colRowType === -1) {
    return { status: "error", message: "SeasonYear or RowType column not found.", removed: 0 };
  }

  // Walk bottom-up, deleting matching rows in merged contiguous
  // chunks. Bottom-up means already-deleted rows never shift the
  // index of rows still waiting to be checked/deleted.
  let removed = 0;
  let i = data.length - 1;
  while (i >= 1) {
    const matches = Number(data[i][colYear]) === Number(year) && data[i][colRowType] === "PRESEASON";
    if (matches) {
      let chunkEnd = i;
      while (i >= 1 && Number(data[i][colYear]) === Number(year) && data[i][colRowType] === "PRESEASON") {
        i--;
      }
      const chunkStart = i + 1; // first matching row in this contiguous run
      const chunkSize = chunkEnd - chunkStart + 1;
      sheet.deleteRows(chunkStart + 1, chunkSize); // +1 for 1-based sheet row
      removed += chunkSize;
    } else {
      i--;
    }
  }

  Logger.log("Removed " + removed + " PRESEASON row(s) for " + year + " from NFL_Schedule.");
  return { status: "ok", removed: removed };
}

function DEBUG_removePreseason2026() {
  return removePreseasonRowsFromSchedule(2026);
}

function DEBUG_whichSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Name: " + ss.getName());
  Logger.log("ID: " + ss.getId());
  Logger.log("URL: " + ss.getUrl());
}