// ======================================================
// AUTOMATION - SCHEDULED ODDS PULLS + "PICKS ARE IN" SMS
// ======================================================
//
// Two pieces:
//   1) Four scheduled odds pulls (Tue 6am/4pm, Wed 6am/3pm),
//      each skipping itself if all picks are already in for
//      the current week - no point tracking line movement
//      once nobody's decision depends on it anymore.
//   2) A Twilio SMS sent once, automatically, the moment every
//      active user has picked every game for the week. Message
//      is a short consensus summary (Parlay's pick per game),
//      not a full per-user breakdown.
//
// Setup required before any of this works:
//   1) Run DEBUG_setupTwilioCredentials() once, after filling
//      in your real Account SID / Auth Token / Twilio number.
//   2) Run installOddsTriggers() once to schedule the 5 pulls.
//   3) Wire checkAndNotifyIfWeekComplete_ into submitWeekPicks
//      in Code.gs (see the separate note - one line to add).
// ======================================================


// ------------------------------------------------------------
// TWILIO CREDENTIALS - stored in Script Properties, never in
// code. Run this once with your real values, from the function
// dropdown (can't pass arguments through the dropdown directly,
// hence this wrapper - same pattern as DEBUG_setTestDate).
// ------------------------------------------------------------
function DEBUG_setupTwilioCredentials() {
  return setupTwilioCredentials(
    "ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx", // <-- your Twilio Account SID
    "your_auth_token_here",               // <-- your Twilio Auth Token
    "+15551234567"                        // <-- your Twilio phone number, E.164 format
  );
}

function setupTwilioCredentials(accountSid, authToken, fromNumber) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("TWILIO_ACCOUNT_SID", accountSid);
  props.setProperty("TWILIO_AUTH_TOKEN", authToken);
  props.setProperty("TWILIO_FROM_NUMBER", fromNumber);
  return { status: "ok", message: "Twilio credentials saved to Script Properties" };
}

/**
 * Converts a phone number to E.164 format (+1XXXXXXXXXX) as
 * Twilio requires. Assumes US numbers.
 */
function toE164_(phone) {
  const digits = String(phone || "").replace(/\D/g, "");
  if (!digits) return "";
  if (digits.length === 10) return "+1" + digits;
  if (digits.length === 11 && digits.charAt(0) === "1") return "+" + digits;
  return "+" + digits;
}

/**
 * Sends one SMS via Twilio's REST API. Returns {status:"ok"} or
 * {status:"error", message} - never throws, so a Twilio failure
 * can't break whatever called it.
 */
function sendTwilioSms_(toPhone, body) {
  const props = PropertiesService.getScriptProperties();
  const sid = props.getProperty("TWILIO_ACCOUNT_SID");
  const token = props.getProperty("TWILIO_AUTH_TOKEN");
  const from = props.getProperty("TWILIO_FROM_NUMBER");

  if (!sid || !token || !from) {
    Logger.log("Twilio credentials not set - run DEBUG_setupTwilioCredentials() first.");
    return { status: "error", message: "Twilio credentials missing" };
  }
  if (!toPhone) {
    return { status: "error", message: "No phone number to send to" };
  }

  const url = "https://api.twilio.com/2010-04-01/Accounts/" + sid + "/Messages.json";
  const options = {
    method: "post",
    payload: { To: toPhone, From: from, Body: body },
    headers: { Authorization: "Basic " + Utilities.base64Encode(sid + ":" + token) },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      return { status: "ok" };
    }
    Logger.log("Twilio send to " + toPhone + " failed (" + code + "): " + response.getContentText());
    return { status: "error", code: code, message: response.getContentText() };
  } catch (err) {
    Logger.log("Twilio send to " + toPhone + " threw: " + err.message);
    return { status: "error", message: err.message };
  }
}


// ------------------------------------------------------------
// WEEK COMPLETION CHECK - true only when every active user has
// a real pick (not blank, not "SKIP") for every GAME that week.
// ------------------------------------------------------------
function isWeekComplete_(year, week) {
  const users = getUsers();
  if (!users.length) return false;

  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week) && (r.RowType === "GAME" || r.RowType === "PRESEASON"));
  if (!games.length) return false;

  const picks = sheetToObjects_(SHEET_PICKS)
    .filter(p => Number(p.SeasonYear) === Number(year) && String(p.Week) === String(week) && p.UserName !== PARLAY_USER);

  const pickedSet = new Set();
  picks.forEach(p => {
    if (p.PickedTeam && p.PickedTeam !== "SKIP") {
      pickedSet.add(p.UserName + "||" + p.EventID);
    }
  });

  for (let i = 0; i < users.length; i++) {
    for (let j = 0; j < games.length; j++) {
      if (!pickedSet.has(users[i] + "||" + games[j].EventID)) {
        return false;
      }
    }
  }
  return true;
}

/**
 * Builds the SMS text: a short header plus the Parlay consensus
 * pick per game (not every individual user's picks).
 */
function buildConsensusSummaryText_(year, week) {
  computeParlayPicksForWeek(year, week); // make sure ParlayPicks is fresh

  const games = sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week) && (r.RowType === "GAME" || r.RowType === "PRESEASON"))
    .sort((a, b) => new Date(a.EventDate) - new Date(b.EventDate));

  const parlayRows = sheetToObjects_(SHEET_PARLAY)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week));
  const parlayByEvent = {};
  parlayRows.forEach(r => parlayByEvent[r.EventID] = r);

  const lines = ["\uD83C\uDFC8 Week " + week + " picks are all in! Here's the consensus:", ""];

  games.forEach(g => {
    const p = parlayByEvent[g.EventID];
    const pick = (p && p.PickedTeam) ? p.PickedTeam : "Tied - no consensus";
    lines.push(g.Matchup + ": " + pick);
  });

  return lines.join("\n");
}

/**
 * Called after every pick submission. If this completes the week
 * (every user has picked every game) and we haven't already sent
 * the notification for this exact year/week, texts every active
 * user the consensus summary. Safe to call repeatedly - only
 * sends once per week thanks to the Script Properties flag.
 */
/**
 * Sends the same SMS body to every active user with a phone number
 * on file. Shared by the completion, tie-break, and tie-resolved
 * notifications so the send logic only lives in one place.
 */
/**
 * Reads a Users-sheet boolean column that should default to TRUE
 * when blank/missing, so adding a new preference column never
 * silently opts existing users out of something they were
 * already getting.
 */
function boolFieldDefaultTrue_(value) {
  if (value === "" || value === undefined || value === null) return true;
  return (value === true || value === "TRUE" || value === "true");
}

/**
 * Sends the same SMS body to every active, opted-in user. category
 * controls which granular preference column (in addition to the
 * SendText master switch) gates this particular message:
 *   "recap"     -> SendRecapInfo (normal "picks are in" consensus)
 *   "tie"       -> SendTieInfo (initial tie alert + final resolution)
 *   "stillTied" -> SendStillTied (cycling "still tied, vote again")
 *   anything else -> only the SendText master switch applies
 */
function sendSmsToAllUsers_(body, category) {
  const usersSheetData = sheetToObjects_(SHEET_USERS);
  const results = [];

  usersSheetData.forEach(u => {
    const isActive = (u.Active === true || u.Active === "TRUE" || u.Active === "true");
    const wantsTexts = boolFieldDefaultTrue_(u.SendText);

    if (!isActive || !wantsTexts || !u.Phone || u.UserName === PARLAY_USER) return;

    let categoryOk = true;
    if (category === "recap") categoryOk = boolFieldDefaultTrue_(u.SendRecapInfo);
    else if (category === "tie") categoryOk = boolFieldDefaultTrue_(u.SendTieInfo);
    else if (category === "stillTied") categoryOk = boolFieldDefaultTrue_(u.SendStillTied);

    if (!categoryOk) return;

    const toPhone = toE164_(u.Phone);
    results.push({ user: u.UserName, result: sendTwilioSms_(toPhone, body) });
  });

  return results;
}

/**
 * Called after every pick submission. If this completes the week
 * (every user has picked every game) and we haven't already notified
 * for this exact year/week, delegates to checkAndNotifyTieBreakStatus_
 * to handle either the tie path or the normal consensus summary.
 * Only ever fires the initial completion check once per week thanks
 * to the Script Properties flag - but the tie logic itself stays
 * live for the rest of the week via checkAndNotifyTieBreakStatus_.
 */
function checkAndNotifyIfWeekComplete_(year, week) {
  const propKey = "NOTIFIED_" + year + "_" + week;
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty(propKey)) {
    return { status: "ok", alreadyNotified: true };
  }

  if (!isWeekComplete_(year, week)) {
    return { status: "ok", complete: false };
  }

  props.setProperty(propKey, "true");

  return checkAndNotifyTieBreakStatus_(year, week);
}

/**
 * Builds a deterministic string representing the CURRENT tie
 * situation for a week: which games are tied, plus everyone's
 * current tie-break votes on those games. Used to detect whether
 * anything has actually changed since the last notification, so
 * we never send a duplicate for a state we've already announced -
 * but WILL correctly re-announce if a NEW or DIFFERENT tie shows
 * up later, even after an earlier tie this same week fully resolved.
 */
function buildTieBreakSignature_(year, week, tiedEventIds) {
  const sortedIds = Array.from(tiedEventIds).sort();

  const votes = sheetToObjects_(SHEET_TIEBREAK)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      String(p.Week) === String(week) &&
      tiedEventIds.has(String(p.EventID))
    );

  const voteParts = votes.map(v => v.EventID + ":" + v.UserName + "=" + v.PickedTeam).sort();

  return sortedIds.join(",") + "||" + voteParts.join("|");
}

/**
 * The single source of truth for tie-break notifications. Call this
 * any time the tie state might have changed (after original picks
 * complete, and after every tie-break vote). Always looks at the
 * CURRENT actual tie state rather than trusting a one-shot flag, so
 * it correctly handles a brand new tie appearing even after an
 * earlier tie this week already fully resolved. Three outcomes:
 *
 *   1) No games tied right now:
 *      - if we'd previously alerted about a tie this week, send the
 *        "tie resolved" SMS with the final consensus, then clear
 *        our record of it so a FUTURE tie this week starts fresh
 *      - otherwise (never was a tie), send the normal consensus SMS
 *
 *   2) Games are tied, but not everyone's voted in this round yet:
 *      do nothing, wait for more votes.
 *
 *   3) Games are tied and (for a re-vote round) everyone's voted, or
 *      this is the very first time we're seeing this tie: send the
 *      appropriate alert, but only if the situation has actually
 *      changed since the last alert we sent (signature-based dedup).
 */
function checkAndNotifyTieBreakStatus_(year, week) {
  const props = PropertiesService.getScriptProperties();
  const sigKey = "TIEBREAK_WEEK_SIG_" + year + "_" + week;
  const lastSig = props.getProperty(sigKey);

  const stillTied = getTiedGamesForWeek(year, week);

  // Outcome 1: nothing tied right now
  if (!stillTied.length) {
    if (lastSig) {
      // We'd previously alerted about a tie this week - announce resolution, then clear
      const messageBody = buildConsensusSummaryText_(year, week);
      const lines = ["\u2705 Tie resolved! Final Week " + week + " consensus:", "", messageBody];
      const results = sendSmsToAllUsers_(lines.join("\n"), "tie");
      props.deleteProperty(sigKey);
      Logger.log("Week " + year + "/" + week + " tie(s) resolved - notified " + results.length + " users.");
      return { status: "ok", resolved: true, sent: results };
    }

    // Never was a tie this week - just the normal completion summary
    const messageBody = buildConsensusSummaryText_(year, week);
    const results = sendSmsToAllUsers_(messageBody, "recap");
    Logger.log("Week " + year + "/" + week + " complete, no ties - notified " + results.length + " users.");
    return { status: "ok", complete: true, tied: 0, sent: results };
  }

  // There ARE tied games right now
  const tiedEventIds = new Set(stillTied.map(r => String(r.EventID)));

  const tieBreakVotes = sheetToObjects_(SHEET_TIEBREAK)
    .filter(p =>
      Number(p.SeasonYear) === Number(year) &&
      String(p.Week) === String(week) &&
      p.PickedTeam &&
      tiedEventIds.has(String(p.EventID))
    );

  const isReVoteRound = tieBreakVotes.length > 0;

  // Outcome 2: if we're in a re-vote round, only alert once EVERYONE
  // has voted on every currently-tied game. (No such gate applies to
  // the very first alert - that goes out immediately once the week's
  // original picks complete and a tie is found.)
  if (isReVoteRound) {
    const users = getUsers();
    const votedSet = new Set(tieBreakVotes.map(v => v.UserName + "||" + v.EventID));

    let allVoted = true;
    users.forEach(u => {
      tiedEventIds.forEach(eventId => {
        if (!votedSet.has(u + "||" + eventId)) allVoted = false;
      });
    });

    if (!allVoted) {
      return { status: "ok", resolved: false, waitingOnVotes: true };
    }
  }

  // Outcome 3: time to alert - but only if this exact situation
  // hasn't already been announced
  const currentSig = buildTieBreakSignature_(year, week, tiedEventIds);
  if (currentSig === lastSig) {
    return { status: "ok", resolved: false, alreadyNotifiedThisState: true };
  }

  props.setProperty(sigKey, currentSig);

  const lines = isReVoteRound
    ? ["\uD83D\uDD01 Still tied after the re-vote! Week " + week + " - please vote again on:", ""]
    : ["\u26A0\uFE0F Week " + week + " picks are in, but we have a tie!", "", "Please re-pick just these game(s) in the Tie Break screen:"];
  stillTied.forEach(g => lines.push("\u2022 " + g.Matchup));

  const results = sendSmsToAllUsers_(lines.join("\n"), isReVoteRound ? "stillTied" : "tie");

  Logger.log("Week " + year + "/" + week + " tie alert sent (" + (isReVoteRound ? "re-vote round" : "initial") + ") to " + results.length + " users.");

  return { status: "ok", resolved: false, tied: stillTied.length, sent: results };
}

/**
 * Testing helper - clears the "already notified" flags for a week
 * so you can re-trigger notifications during testing.
 */
function clearWeekNotifiedFlag(year, week) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("NOTIFIED_" + year + "_" + week);
  props.deleteProperty("TIEBREAK_WEEK_SIG_" + year + "_" + week);
  // Clean up any leftover flags from the older PENDING/RESOLVED design
  props.deleteProperty("TIEBREAK_PENDING_" + year + "_" + week);
  props.deleteProperty("TIEBREAK_RESOLVED_" + year + "_" + week);
  props.deleteProperty("TIEBREAK_SIGNATURE_" + year + "_" + week);
  return { status: "ok" };
}


// ------------------------------------------------------------
// SCHEDULED ODDS PULLS - four per week, each skipping itself if
// all picks are already locked in. Pulls only the CURRENT week's
// odds (not all 18 weeks like the manual snapshotNFLCurrentOdds),
// since that's the only week anyone's decision-making depends on.
// ------------------------------------------------------------
function snapshotOddsForWeek_(year, week, isOpening, isClosing) {
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

  const weekStr = String(week);
  const isPreseason = weekStr.charAt(0) === "P";
  const seasonType = isPreseason ? 1 : 2;
  const espnWeek = isPreseason ? weekStr.substring(1) : weekStr;

  const url =
    "https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard" +
    "?season=" + year + "&seasontype=" + seasonType + "&week=" + espnWeek;

  const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    Logger.log("Odds pull failed for " + year + " Week " + week + ": HTTP " + response.getResponseCode());
    return { status: "error", message: "ESPN request failed" };
  }

  const data = JSON.parse(response.getContentText());
  const pulledAt = new Date();
  const rows = [];

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

    rows.push([
      snapshotId, event.id || "", year, week, matchup,
      event.date ? new Date(event.date) : "", pulledAt,
      provider.name || provider.displayName || provider.nameDisplay || "",
      odds.details || "", odds.overUnder ?? "",
      awayOdds.moneyLine ?? "", homeOdds.moneyLine ?? "",
      !!isOpening, !!isClosing
    ]);
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }

  Logger.log("Odds snapshot " + year + " Week " + week + ": " + rows.length +
    " rows added (opening=" + !!isOpening + ", closing=" + !!isClosing + ")");

  return { status: "ok", rowsAdded: rows.length };
}

/**
 * Shared entry point for all 4 scheduled pulls: figures out the
 * current year/week, skips entirely if picks are already
 * complete, otherwise pulls fresh odds for that week only.
 */
function scheduledOddsPull_(isOpening, isClosing) {
  const year = getCurrentSeasonYear();
  if (!year) {
    Logger.log("No current season found in NFL_Schedule - skipping odds pull.");
    return { status: "skipped", reason: "no current season" };
  }

  const week = getDefaultWeek(year);
  if (!week) {
    Logger.log("No current week found for " + year + " - skipping odds pull.");
    return { status: "skipped", reason: "no current week" };
  }

  if (isWeekComplete_(year, week)) {
    Logger.log("All picks already in for " + year + " Week " + week + " - skipping this scheduled odds pull.");
    return { status: "skipped", reason: "picks already complete" };
  }

  return snapshotOddsForWeek_(year, week, isOpening, isClosing);
}

// The 5 actual trigger targets - install these via installOddsTriggers()
function snapshotOddsTue1205am() { return scheduledOddsPull_(true, false); }  // opening line - first pull, right after picks open
function snapshotOddsTue6am() { return scheduledOddsPull_(false, false); }
function snapshotOddsTue4pm() { return scheduledOddsPull_(false, false); }
function snapshotOddsWed6am() { return scheduledOddsPull_(false, false); }
function snapshotOddsWed3pm() { return scheduledOddsPull_(false, true); }   // closing-ish, last pull before games

/**
 * Run ONCE to install the 4 scheduled triggers. Safe to re-run -
 * it deletes any existing triggers for these 4 functions first,
 * so re-running never creates duplicates. Trigger times use
 * whatever timezone is set in Project Settings - verify that's
 * America/Los_Angeles (or your intended zone) before relying on
 * exact times. Apps Script fires time-based triggers "near" the
 * requested hour, not to the exact minute.
 */
function installOddsTriggers() {
  const handlerNames = ["snapshotOddsTue1205am", "snapshotOddsTue6am", "snapshotOddsTue4pm", "snapshotOddsWed6am", "snapshotOddsWed3pm"];

  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("snapshotOddsTue1205am").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(0).nearMinute(5).create();
  ScriptApp.newTrigger("snapshotOddsTue6am").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(6).nearMinute(0).create();
  ScriptApp.newTrigger("snapshotOddsTue4pm").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(16).nearMinute(0).create();
  ScriptApp.newTrigger("snapshotOddsWed6am").timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(6).nearMinute(0).create();
  ScriptApp.newTrigger("snapshotOddsWed3pm").timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(15).nearMinute(0).create();

  Logger.log("Odds pull triggers installed: Tue 12:05am, Tue 6am, Tue 4pm, Wed 6am, Wed 3pm (script timezone: " + Session.getScriptTimeZone() + ")");

  return { status: "ok", timezone: Session.getScriptTimeZone() };
}

/**
 * Removes all 4 scheduled odds-pull triggers, if you ever want
 * to pause or stop this automation.
 */
function removeOddsTriggers() {
  const handlerNames = ["snapshotOddsTue1205am", "snapshotOddsTue6am", "snapshotOddsTue4pm", "snapshotOddsWed6am", "snapshotOddsWed3pm"];
  let removed = 0;

  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  Logger.log("Removed " + removed + " odds-pull triggers.");
  return { status: "ok", removed: removed };
}

function scheduledPullFinalScores_() {
  const year = getCurrentSeasonYear();
  if (!year) {
    Logger.log("No current season found in NFL_Schedule - skipping results pull.");
    return { status: "skipped", reason: "no current season" };
  }

  const week = getDefaultWeek(year);
  if (!week) {
    Logger.log("No current week found for " + year + " - skipping results pull.");
    return { status: "skipped", reason: "no current week" };
  }

  return pullFinalScores(year, week);
}

// The 3 actual trigger targets - install these via installResultsTriggers()
function pullFinalScoresFri5am() { return scheduledPullFinalScores_(); }
function pullFinalScoresMon5am() { return scheduledPullFinalScores_(); }
function pullFinalScoresTue5am() { return scheduledPullFinalScores_(); }

/**
 * Run ONCE to install the 3 scheduled results-pull triggers. Safe
 * to re-run - deletes any existing triggers for these 3 functions
 * first, so re-running never creates duplicates. Trigger times use
 * whatever timezone is set in Project Settings - verify that's
 * America/Los_Angeles (or your intended zone) before relying on
 * exact times.
 */
function installResultsTriggers() {
  const handlerNames = ["pullFinalScoresFri5am", "pullFinalScoresMon5am", "pullFinalScoresTue5am"];

  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("pullFinalScoresFri5am").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(5).nearMinute(0).create();
  ScriptApp.newTrigger("pullFinalScoresMon5am").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(5).nearMinute(0).create();
  ScriptApp.newTrigger("pullFinalScoresTue5am").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(5).nearMinute(0).create();

  Logger.log("Results pull triggers installed: Fri 5am, Mon 5am, Tue 5am (script timezone: " + Session.getScriptTimeZone() + ")");

  return { status: "ok", timezone: Session.getScriptTimeZone() };
}

/**
 * Removes all 3 scheduled results-pull triggers, if you ever want
 * to pause or stop this automation.
 */
function removeResultsTriggers() {
  const handlerNames = ["pullFinalScoresFri5am", "pullFinalScoresMon5am", "pullFinalScoresTue5am"];
  let removed = 0;

  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  Logger.log("Removed " + removed + " results-pull triggers.");
  return { status: "ok", removed: removed };
}