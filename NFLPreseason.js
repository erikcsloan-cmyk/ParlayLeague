// ------------------------------------------------------------
// PRESEASON CONNECTION TEST - pulls real, currently-live preseason
// games into a SEPARATE, isolated sheet (NFL_Schedule_PreseasonTest),
// never touching the real NFL_Schedule. Preseason uses the same
// Week numbering (1-4) as early regular season, so mixing them into
// the same sheet would cause exactly the kind of collision/corruption
// we've already fought through once this project - this keeps them
// fully separate on purpose.
//
// Good for: proving the ESPN connection, date-based fetch, and
// write-to-sheet pipeline all work end-to-end against real live
// data, without any risk to production.
// ------------------------------------------------------------
function testPullPreseasonConnection(year) {
  const sheetName = "NFL_Schedule_PreseasonTest";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const headers = [
    "SeasonYear", "Week", "EventID", "RowType", "EventDate", "EventTime",
    "AwayTeam", "AwayShort", "HomeTeam", "HomeShort", "Matchup",
    "Venue", "City", "State", "Status"
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.setFrozenRows(1);

  // Preseason typically runs early-to-mid August
  const startDate = new Date(year + "-08-01");
  const endDate = new Date(year + "-08-31");

  Logger.log("================================");
  Logger.log("PRESEASON CONNECTION TEST - " + year);
  Logger.log("Date range: " + startDate.toDateString() + " to " + endDate.toDateString());
  Logger.log("================================");

  const recordsByEventId = {};
  const currentDate = new Date(startDate);
  let daysChecked = 0;
  let daysWithGames = 0;

  while (currentDate <= endDate) {
    const espnDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyyMMdd");
    const url =
      "https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard" +
      "?dates=" + espnDate;

    daysChecked++;

    try {
      const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
      const code = response.getResponseCode();

      if (code !== 200) {
        Logger.log(espnDate + ": HTTP " + code);
        currentDate.setDate(currentDate.getDate() + 1);
        continue;
      }

      const data = JSON.parse(response.getContentText());
      const events = data.events || [];

      if (events.length) {
        daysWithGames++;
        Logger.log(espnDate + ": " + events.length + " game(s) found");
      }

      events.forEach(event => {
        const seasonType = Number((event.season && event.season.type) || (data.season && data.season.type) || 0);
        if (seasonType !== 1) return; // preseason ONLY - explicitly the opposite filter from the regular-season puller

        const eventSeason = Number((event.season && event.season.year) || (data.season && data.season.year) || 0);
        if (eventSeason && eventSeason !== year) return;

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
        const status = (event.status && event.status.type && event.status.type.description) || "";

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
          State: address.state || "",
          Status: status
        };
      });

    } catch (err) {
      Logger.log("ERROR " + espnDate + ": " + err.message);
    }

    currentDate.setDate(currentDate.getDate() + 1);
    Utilities.sleep(75);
  }

  const records = Object.values(recordsByEventId);
  const rows = records.map(r => headers.map(h => (r[h] !== undefined ? r[h] : "")));

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.autoResizeColumns(1, headers.length);

  Logger.log("================================");
  Logger.log("DONE - days checked: " + daysChecked + ", days with any games: " + daysWithGames);
  Logger.log("PRESEASON games written to " + sheetName + ": " + records.length);
  Logger.log("================================");

  return { status: "ok", daysChecked: daysChecked, daysWithGames: daysWithGames, gamesWritten: records.length };
}

// Convenience wrapper - run this directly from the function dropdown
function testPullPreseasonConnection2026() {
  return testPullPreseasonConnection(2026);
}

/**
 * Removes the isolated test sheet entirely once you're done
 * verifying the connection works.
 */
function removePreseasonTestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("NFL_Schedule_PreseasonTest");
  if (sheet) {
    ss.deleteSheet(sheet);
    return { status: "ok", message: "Test sheet removed" };
  }
  return { status: "ok", message: "Test sheet didn't exist" };
}

// ------------------------------------------------------------
// PRESEASON - INTEGRATED PIPELINE TEST
//
// Pulls real preseason games directly into the live NFL_Schedule
// sheet, tagged RowType="PRESEASON" (not "GAME"). Every function
// in this app filters strictly on RowType === "GAME", so these
// rows are completely invisible to getWeeks, computeParlayPicksForWeek,
// getWeekData, the bye-week calculator, etc. - they coexist safely
// with real season data and can't collide with the Week 1-4
// numbering regular season also uses.
//
// Upserts by EventID (never deletes/touches existing GAME/BYE rows).
// Run removePreseasonRowsFromSchedule afterward to clean up.
// ------------------------------------------------------------
function pullPreseasonIntoScheduleForTesting(year) {
  const sheetName = "NFL_Schedule";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { status: "error", message: "NFL_Schedule not found" };
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colEventId = headers.indexOf("EventID");
  if (colEventId === -1) {
    return { status: "error", message: "EventID column not found in NFL_Schedule" };
  }

  const startDate = new Date(year + "-08-01");
  const endDate = new Date(year + "-08-31");

  Logger.log("================================");
  Logger.log("PULLING PRESEASON INTO NFL_Schedule (tagged RowType=PRESEASON) - " + year);
  Logger.log("================================");

  const recordsByEventId = {};
  const currentDate = new Date(startDate);

  while (currentDate <= endDate) {
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
          Logger.log(espnDate + ": " + events.length + " game(s) found");
        }

        events.forEach(event => {
          const seasonType = Number((event.season && event.season.type) || (data.season && data.season.type) || 0);
          if (seasonType !== 1) return; // preseason ONLY

          const eventSeason = Number((event.season && event.season.year) || (data.season && data.season.year) || 0);
          if (eventSeason && eventSeason !== year) return;

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
          const rawWeek = (event.week && event.week.number) || (data.week && data.week.number) || "";
          const week = rawWeek ? ("P" + rawWeek) : ""; // e.g. "P1" - visually distinct from regular season "1" at a glance

          recordsByEventId[String(event.id)] = {
            SeasonYear: year,
            Week: week,
            EventID: event.id,
            RowType: "PRESEASON",
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

  const records = Object.values(recordsByEventId);

  const existingData = sheet.getDataRange().getValues();
  const existingRowByEventId = {};
  for (let i = 1; i < existingData.length; i++) {
    existingRowByEventId[String(existingData[i][colEventId])] = i + 1;
  }

  let updated = 0;
  const newRows = [];

  records.forEach(record => {
    const rowArray = headers.map(h => (record[h] !== undefined ? record[h] : ""));
    const existingRow = existingRowByEventId[String(record.EventID)];
    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, headers.length).setValues([rowArray]);
      updated++;
    } else {
      newRows.push(rowArray);
    }
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }

  Logger.log("================================");
  Logger.log("DONE - " + updated + " PRESEASON rows updated, " + newRows.length + " new rows added");
  Logger.log("Tagged RowType=PRESEASON - invisible to all live app logic, which filters on RowType=GAME.");
  Logger.log("================================");

  return { status: "ok", updated: updated, inserted: newRows.length };
}

/**
 * Removes every PRESEASON-tagged row for a given year from
 * NFL_Schedule, leaving all real GAME/BYE rows completely untouched.
 */
function removePreseasonRowsFromSchedule(year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NFL_Schedule");
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: "ok", removed: 0 };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colYear = headers.indexOf("SeasonYear");
  const colRowType = headers.indexOf("RowType");

  let removed = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (Number(data[i][colYear]) === Number(year) && data[i][colRowType] === "PRESEASON") {
      sheet.deleteRow(i + 1);
      removed++;
    }
  }

  Logger.log("Removed " + removed + " PRESEASON rows for " + year + " from NFL_Schedule.");
  return { status: "ok", removed: removed };
}

// Convenience wrappers - run these directly from the function dropdown
function pullPreseasonIntoSchedule2026() {
  return pullPreseasonIntoScheduleForTesting(2026);
}

function removePreseason2026() {
  return removePreseasonRowsFromSchedule(2026);
}