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
//   2) Run installOddsTriggers() once to schedule the 4 pulls.
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

  const lines = ["\uD83C\uDFC8 Week " + week + " consensus is in:", ""];

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
 *   "runback"   -> SendRunbackInfo (Thursday-loss runback vote open/decided)
 *   "parlay"    -> SendParlayInfo (week is open for picks; Parlay / runback placed)
 *   anything else -> only the SendText master switch applies
 */
function sendSmsToAllUsers_(body, category) {
  const usersSheetData = sheetToObjects_(SHEET_USERS);
  const results = [];
  const seenPhones = {}; // one text per phone number, even if two Users rows share it

  usersSheetData.forEach(u => {
    const isActive = (u.Active === true || u.Active === "TRUE" || u.Active === "true");
    const wantsTexts = boolFieldDefaultTrue_(u.SendText);

    if (!isActive || !wantsTexts || !u.Phone || u.UserName === PARLAY_USER) return;

    let categoryOk = true;
    if (category === "recap") categoryOk = boolFieldDefaultTrue_(u.SendRecapInfo);
    else if (category === "tie") categoryOk = boolFieldDefaultTrue_(u.SendTieInfo);
    else if (category === "stillTied") categoryOk = boolFieldDefaultTrue_(u.SendStillTied);
    else if (category === "runback") categoryOk = boolFieldDefaultTrue_(u.SendRunbackInfo);
    else if (category === "parlay") categoryOk = boolFieldDefaultTrue_(u.SendParlayInfo);

    if (!categoryOk) return;

    const phoneKey = normalizePhone_(u.Phone);
    if (seenPhones[phoneKey]) return;
    seenPhones[phoneKey] = true;

    const toPhone = toE164_(u.Phone);
    results.push({ user: u.UserName, result: sendTwilioSms_(toPhone, body) });
  });

  return results;
}

/**
 * Atomic "do this only once" flags. The old read-flag-then-set-flag
 * pattern let two overlapping runs (two triggers, two near-simultaneous
 * submissions) both see "not sent yet" and both text everyone. These
 * take the script lock for just the read+write, so exactly one caller
 * wins.
 */
function claimOnce_(key) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(key)) return false;
    props.setProperty(key, "true");
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Sets key to value; true only if that changed it (so only one caller announces a given state). */
function claimChange_(key, value) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(key) === value) return false;
    props.setProperty(key, value);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Removes key; true only if it existed (so only one caller announces the resolution). */
function claimRemove_(key) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty(key)) return false;
    props.deleteProperty(key);
    return true;
  } finally {
    lock.releaseLock();
  }
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

  if (!claimOnce_(propKey)) return { status: "ok", alreadyNotified: true };

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
 * earlier tie this week already fully resolved. Every SMS this
 * function sends (recap, tie alert/re-vote, tie resolved) also gets
 * posted to Home via postSystemNotice_ (Code.gs), independent of
 * anyone's text preferences - so it's still visible to someone who
 * opted out of that category of text, or simply missed it. Three
 * outcomes:
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
      // (text only: tie alerts no longer post to Home)
      if (!claimRemove_(sigKey)) return { status: "ok", alreadyNotifiedThisState: true };
      const messageBody = buildConsensusSummaryText_(year, week);
      const lines = ["\u2705 Tie resolved! Week " + week + " consensus:", "", messageBody];
      const results = sendSmsToAllUsers_(lines.join("\n"), "tie");
      Logger.log("Week " + year + "/" + week + " tie(s) resolved - notified " + results.length + " users.");
      return { status: "ok", resolved: true, sent: results };
    }

    // Never was a tie this week - just the normal completion summary
    const messageBody = buildConsensusSummaryText_(year, week);
    const results = sendSmsToAllUsers_(messageBody, "recap");
    postSystemNotice_(messageBody);
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

  if (!claimChange_(sigKey, currentSig)) {
    return { status: "ok", resolved: false, alreadyNotifiedThisState: true };
  }

  const lines = isReVoteRound
    ? ["\u2696\uFE0F Still tied! Week " + week + " \u2014 vote again:", ""]
    : ["\u26A0\uFE0F Week " + week + " tie! Re-pick on the Ties tab:", ""];
  stillTied.forEach(g => lines.push("\u2022 " + g.Matchup));

  const results = sendSmsToAllUsers_(lines.join("\n"), isReVoteRound ? "stillTied" : "tie");
  // text only - tie alerts no longer post to Home

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
  props.deleteProperty("RUNBACK_OPENED_" + year + "_" + week);
  props.deleteProperty("RUNBACK_DECIDED_" + year + "_" + week);
  props.deleteProperty("KICKOFF_SENT_" + year + "_" + week);
  props.deleteProperty("PARLAY_PLACED_NOTIFIED_" + year + "_" + week);
  props.deleteProperty("RUNBACK_PLACED_NOTIFIED_" + year + "_" + week);
  return { status: "ok" };
}


// ------------------------------------------------------------
// RUNBACK VOTE NOTIFICATIONS - the single source of truth for
// texting the group about a runback vote, same signature/flag
// pattern as checkAndNotifyTieBreakStatus_ above. Safe to call
// repeatedly and from multiple places (after final scores are
// pulled, after every vote, and opportunistically from the Home
// tab) - each of the two texts it can send only ever goes out once
// per week thanks to the Script Properties flags. The passed/failed
// outcomes also get posted to Home via postSystemNotice_ (Code.gs)
// so anyone who missed or opted out of the text still sees the
// result; the OPEN moment does not, since Home already shows a
// live, interactive Yes/No notice for that while the vote is open.
// getRunbackStatus() and recordRunbackVoteResult_() live in
// Code.gs, but Apps Script shares one global scope across files so
// they're callable from here directly.
// ------------------------------------------------------------
// Flip to true once the Twilio number + webhook are set up (setupSmsWebhook),
// so the vote text can invite people to reply YES / NO. Until then it points
// to the Home page instead, since a text reply wouldn't be counted.
const SMS_REPLY_VOTING_LIVE_ = false;

function runbackOpenText_(week, st) {
  const closes = st.deadline ? (" (closes " + st.deadline + ")") : "";
  if (SMS_REPLY_VOTING_LIVE_) {
    return "🔄 Runback vote is open. Reply RB YES or RB NO" + closes + ".";
  }
  let link = "";
  try { link = ScriptApp.getService().getUrl() || ""; } catch (err) { link = ""; }
  // Replies aren't wired up yet, so also point to the Home page.
  return "🔄 Runback vote is open. Reply RB YES or RB NO" + closes + ", or vote on the Home page." + (link ? "\n" + link : "");
}

/**
 * Run from the function dropdown when a runback vote is open but a text
 * didn't arrive. Shows who WOULD be texted and why anyone wouldn't, plus the
 * state of the "already sent" flags. Sends nothing.
 */
function DEBUG_runbackTextStatus() {
  const year = getCurrentSeasonYear();
  const week = year ? getDefaultWeek(year) : null;
  const props = PropertiesService.getScriptProperties();
  const st = (year && week) ? getRunbackStatus(null, year, week) : null;
  const users = sheetToObjects_(SHEET_USERS).filter(u => u.UserName !== PARLAY_USER).map(u => {
    const active = (u.Active === true || u.Active === "TRUE" || u.Active === "true");
    const reasons = [];
    if (!active) reasons.push("not Active");
    if (!boolFieldDefaultTrue_(u.SendText)) reasons.push("master Text switch off");
    if (!boolFieldDefaultTrue_(u.SendRunbackInfo)) reasons.push("Runback votes box off");
    if (!u.Phone) reasons.push("no phone");
    return { user: u.UserName, phone: u.Phone ? "on file" : "none", willReceive: reasons.length === 0, why: reasons.join(", ") };
  });
  const out = {
    year: year, week: week,
    runbackEligible: !!(st && st.eligible), voteOpen: !!(st && st.open), decided: st ? st.decided : null,
    openTextFlagSet: !!props.getProperty("RUNBACK_OPENED_" + year + "_" + week),
    decidedFlagSet: !!props.getProperty("RUNBACK_DECIDED_" + year + "_" + week),
    users: users
  };
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

/**
 * Re-sends the "runback vote is open" text to everyone it applies to
 * (ignores the already-sent flag). Anyone who already got it will get a
 * second copy, so check DEBUG_runbackTextStatus first.
 */
function DEBUG_resendRunbackOpenText() {
  const year = getCurrentSeasonYear();
  const week = year ? getDefaultWeek(year) : null;
  const st = (year && week) ? getRunbackStatus(null, year, week) : null;
  if (!st || !st.eligible || !st.open || st.decided) {
    Logger.log("No open runback vote right now - nothing sent.");
    return { status: "ok", sent: 0, reason: "no open runback vote" };
  }
  const results = sendSmsToAllUsers_(runbackOpenText_(week, st), "runback");
  Logger.log("Runback text re-sent to " + results.length + " users: " + JSON.stringify(results));
  return { status: "ok", sent: results.length, results: results };
}

function checkAndNotifyRunbackStatus_(year, week) {
  const props = PropertiesService.getScriptProperties();
  const decidedKey = "RUNBACK_DECIDED_" + year + "_" + week;
  if (props.getProperty(decidedKey)) {
    return { status: "ok", alreadyFinal: true };
  }

  const st = getRunbackStatus(null, year, week);
  if (!st.eligible) {
    return { status: "ok", eligible: false };
  }

  const openKey = "RUNBACK_OPENED_" + year + "_" + week;
  if (!props.getProperty(openKey) && claimOnce_(openKey)) {
    const lines = [runbackOpenText_(week, st)];
    const results = sendSmsToAllUsers_(lines.join("\n"), "runback");
    Logger.log("Runback vote opened for " + year + "/" + week + " - notified " + results.length + " users.");
  }

  if (st.decided === "Yes") {
    if (!claimOnce_(decidedKey)) return { status: "ok", alreadyFinal: true };
    recordRunbackVoteResult_(year, week, st.yesVoters || []);

    const lines = ["✅ Runback vote passed."];
    const results = sendSmsToAllUsers_(lines.join("\n"), "runback");
    postSystemNotice_(lines.join("\n"));
    Logger.log("Runback vote passed for " + year + "/" + week + " - notified " + results.length + " users.");
    return { status: "ok", decided: "Yes", sent: results };
  }

  if (st.decided === "No") {
    if (!claimOnce_(decidedKey)) return { status: "ok", alreadyFinal: true };
    const lines = ["❌ Runback vote failed."];
    const results = sendSmsToAllUsers_(lines.join("\n"), "runback");
    postSystemNotice_(lines.join("\n"));
    Logger.log("Runback vote closed with no majority for " + year + "/" + week + " - notified " + results.length + " users.");
    return { status: "ok", decided: "No", sent: results };
  }

  return { status: "ok", decided: null, open: st.open };
}


// ------------------------------------------------------------
// PICK REMINDERS - self-managed, opt-in, up to 3 per user. Each
// slot is a day (Tue/Wed/Thu) + time, stored on the Users sheet
// (PickReminder1/2/3 Enabled/Day/Time). Slot 3's message is more
// direct/urgent than slots 1 and 2 (see checkPickReminderSlot_).
// A single background check
// runs every 15 minutes (installPickReminderTrigger) and fires a
// reminder to a user's OWN phone only, skipping anyone who's
// already picked the full week. Message is fixed/standard - there's
// no per-user custom text, matching how every other automated text
// in this app works. Unlike sendSmsToAllUsers_, this never broadcasts
// to the group - each send targets exactly one phone number.
// ------------------------------------------------------------
function installPickReminderTrigger() {
  removePickReminderTrigger();
  ScriptApp.newTrigger("checkPickReminders").timeBased().everyMinutes(15).create();
  return { status: "ok", message: "Pick reminder check installed - runs every 15 minutes." };
}

function removePickReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "checkPickReminders") ScriptApp.deleteTrigger(t);
  });
  return { status: "ok" };
}

/**
 * Minutes since midnight for an "HH:mm" string, or null if it
 * doesn't parse - used to match a user's chosen reminder time
 * against a 15-minute-wide window ending now, so a slightly-late
 * trigger run still catches a reminder due earlier in that window.
 */
function minutesSinceMidnight_(hhmm) {
  // normalizeReminderTime_ (Code.gs) handles the case where Sheets
  // auto-converted the cell into a Date instead of keeping "HH:mm"
  // text - this reads straight from the sheet via sheetToObjects_,
  // not through getUserProfile, so it needs that same normalization.
  const normalized = normalizeReminderTime_(hhmm);
  const m = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(normalized);
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]);
}

/**
 * Fires (or skips) one reminder slot for one user. Dedup key is
 * per user+slot+year+week+day-of-week, so a slot only ever sends
 * once per week even if the 15-minute check runs more than once
 * inside the matching window.
 */
function checkPickReminderSlot_(u, slotNum, todayName, nowMinutes, year, week, alreadyPicked) {
  const enabled = u["PickReminder" + slotNum + "Enabled"];
  const isEnabled = (enabled === true || enabled === "TRUE" || enabled === "true");
  if (!isEnabled) return;
  const dayName = u["PickReminder" + slotNum + "Day"];
  if (dayName !== todayName) return;

  const reminderMinutes = minutesSinceMidnight_(u["PickReminder" + slotNum + "Time"]);
  if (reminderMinutes == null) return;
  if (nowMinutes < reminderMinutes || nowMinutes >= reminderMinutes + 15) return;

  if (alreadyPicked) return; // skip - already picked this week, no need to nag

  const wantsTexts = boolFieldDefaultTrue_(u.SendText);
  if (!wantsTexts || !u.Phone) return;

  const props = PropertiesService.getScriptProperties();
  const dedupKey = "PICKREMINDER_SENT_" + u.UserName + "_" + slotNum + "_" + year + "_" + week + "_" + dayName;
  if (!claimOnce_(dedupKey)) return;

  const body = (slotNum === "3")
    ? "‼️Yo, you are holding up the Parlay. Submit your picks."
    : "⏰ Week " + week + " reminder: you haven't picked yet!";
  const result = sendTwilioSms_(toE164_(u.Phone), body);
  // sendTwilioSms_ never throws, so a bad phone number or bad Twilio
  // credentials would otherwise fail completely silently - log it so
  // it shows up in Executions even though nothing threw.
  Logger.log("Pick reminder slot " + slotNum + " for " + u.UserName + " -> " + toE164_(u.Phone) + ": " + JSON.stringify(result));
}

/**
 * The trigger entry point, run every 15 minutes. Cheap to run when
 * there's no active week or nobody has a reminder configured -
 * bails out immediately in either case.
 */
function checkPickReminders() {
  const year = getCurrentSeasonYear();
  if (!year) return { status: "ok", skipped: "no current season" };
  const week = getDefaultWeek(year);
  if (!week) return { status: "ok", skipped: "no current week" };

  // "The week is open" text rides on this same 15-minute check.
  try {
    checkWeekKickoff_(year, week);
  } catch (err) {
    Logger.log("checkWeekKickoff_ failed: " + err.message);
  }

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayName = Utilities.formatDate(now, tz, "EEEE");
  const nowMinutes = now.getHours() * 60 + now.getMinutes();

  const users = sheetToObjects_(SHEET_USERS).filter(u =>
    (u.Active === true || u.Active === "TRUE" || u.Active === "true") && u.UserName !== PARLAY_USER
  );

  users.forEach(u => {
    const alreadyPicked = hasUserPickedAllGamesThisWeek_(u.UserName, year, week);
    ["1", "2", "3"].forEach(slotNum => {
      checkPickReminderSlot_(u, slotNum, todayName, nowMinutes, year, week, alreadyPicked);
    });
  });

  return { status: "ok" };
}


/**
 * Clears the "already sent" dedup flags for one user's pick
 * reminders for a given week (all 3 slots, all 3 possible days) -
 * useful while testing, so you can re-trigger a reminder without
 * waiting for next week. Does NOT touch whether the reminder is
 * enabled or its day/time - only the "have I already texted this
 * person this week" memory.
 */
function clearPickReminderFlags(userName, year, week) {
  const props = PropertiesService.getScriptProperties();
  PICK_REMINDER_DAYS_.forEach(dayName => {
    ["1", "2", "3"].forEach(slotNum => {
      props.deleteProperty("PICKREMINDER_SENT_" + userName + "_" + slotNum + "_" + year + "_" + week + "_" + dayName);
    });
  });
  return { status: "ok" };
}

/**
 * Diagnostic - walks through every check checkPickReminderSlot_
 * makes for one user's three reminder slots RIGHT NOW, and reports
 * exactly which one (if any) is blocking a send. Run via the
 * DEBUG_checkMyPickReminder() wrapper below (fill in your UserName
 * there first) since the function dropdown can't pass arguments
 * directly - same pattern as DEBUG_setupTwilioCredentials.
 */
function debugPickReminder(userName) {
  const year = getCurrentSeasonYear();
  if (!year) return { status: "ok", issue: "No current season detected - getCurrentSeasonYear() returned nothing." };
  const week = getDefaultWeek(year);
  if (!week) return { status: "ok", issue: "No current week detected - getDefaultWeek() returned nothing." };

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayName = Utilities.formatDate(now, tz, "EEEE");
  const nowMinutes = now.getHours() * 60 + now.getMinutes();

  const users = sheetToObjects_(SHEET_USERS);
  const u = users.find(x => x.UserName === userName);
  if (!u) return { status: "error", message: "User not found: " + userName };

  const alreadyPicked = hasUserPickedAllGamesThisWeek_(u.UserName, year, week);
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "checkPickReminders");

  const slots = ["1", "2", "3"].map(slotNum => {
    const enabledRaw = u["PickReminder" + slotNum + "Enabled"];
    const isEnabled = (enabledRaw === true || enabledRaw === "TRUE" || enabledRaw === "true");
    const day = u["PickReminder" + slotNum + "Day"];
    const rawTime = u["PickReminder" + slotNum + "Time"];
    const time = normalizeReminderTime_(rawTime);
    const wasStoredAsDate = (rawTime instanceof Date);
    const reminderMinutes = minutesSinceMidnight_(time);

    let willFireInCurrentWindow = false;
    let reasonNotFiring = "";
    if (!isEnabled) {
      reasonNotFiring = "Slot is not enabled.";
    } else if (day !== todayName) {
      reasonNotFiring = "Day is set to " + day + ", today is " + todayName + " - won't fire today.";
    } else if (reminderMinutes == null) {
      reasonNotFiring = "Time value \"" + time + "\" didn't parse as HH:mm.";
    } else if (nowMinutes < reminderMinutes || nowMinutes >= reminderMinutes + 15) {
      reasonNotFiring = "Time is " + time + ", current script time is " + Utilities.formatDate(now, tz, "HH:mm") +
        " (" + tz + ") - outside the 15-minute firing window.";
    } else if (alreadyPicked) {
      reasonNotFiring = "You've already picked every game this week - reminder is intentionally skipped.";
    } else if (!boolFieldDefaultTrue_(u.SendText)) {
      reasonNotFiring = "Your \"Text me\" master switch (SendText) is off.";
    } else if (!u.Phone) {
      reasonNotFiring = "No phone number on file for you.";
    } else {
      willFireInCurrentWindow = true;
    }

    const dedupKey = "PICKREMINDER_SENT_" + u.UserName + "_" + slotNum + "_" + year + "_" + week + "_" + day;

    return {
      slot: slotNum,
      enabledRaw: enabledRaw,
      isEnabled: isEnabled,
      day: day,
      time: time,
      timeWasStoredAsDateNotText: wasStoredAsDate,
      willFireInCurrentWindow: willFireInCurrentWindow,
      reasonNotFiring: reasonNotFiring,
      alreadySentThisWeekFlag: !!PropertiesService.getScriptProperties().getProperty(dedupKey)
    };
  });

  return {
    status: "ok",
    scriptTimeZone: tz,
    nowInScriptTimeZone: Utilities.formatDate(now, tz, "EEEE h:mm a"),
    year: year,
    week: week,
    alreadyPickedThisWeek: alreadyPicked,
    phoneOnFile: !!u.Phone,
    sendTextMasterSwitch: boolFieldDefaultTrue_(u.SendText),
    pickReminderTriggerInstalled: triggers.length > 0,
    pickReminderTriggerCount: triggers.length,
    slots: slots
  };
}

// Fill in your UserName (exactly as it appears on the Users sheet),
// run this from the function dropdown, then check View > Logs (or
// View > Executions) for the result.
function DEBUG_checkMyPickReminder() {
  const result = debugPickReminder("PUT_YOUR_USERNAME_HERE");
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


// ------------------------------------------------------------
// SCHEDULED ODDS PULLS - four per week, each skipping itself if
// all picks are already locked in. Pulls only the CURRENT week's
// odds (not all 18 weeks like the manual snapshotNFLCurrentOdds),
// since that's the only week anyone's decision-making depends on.
// ------------------------------------------------------------

/**
 * Pulls a moneyline number for one side ("away" or "home") out of an
 * ESPN odds object, tolerating both schemas ESPN has used: the older
 * flat odds.awayTeamOdds.moneyLine / odds.homeTeamOdds.moneyLine
 * number, and the newer odds.moneyline.away/home.close.odds (falling
 * back to .open.odds) string like "-185" or "+154". ESPN migrated to
 * the new shape mid-season and dropped the old field entirely, which
 * is why moneylines started coming back blank on some/all pulls -
 * the old field simply isn't there anymore. Returns "" only when
 * neither shape has a value, so a genuinely-not-yet-posted line still
 * reads as blank instead of 0.
 */
function extractMoneyline_(odds, side) {
  const sideKey = side === "away" ? "awayTeamOdds" : "homeTeamOdds";
  const legacyVal = odds[sideKey] && odds[sideKey].moneyLine;
  if (legacyVal !== undefined && legacyVal !== null && legacyVal !== "") {
    return Number(legacyVal);
  }

  const ml = odds.moneyline && odds.moneyline[side];
  if (ml) {
    const rawVal = (ml.close && ml.close.odds) || (ml.open && ml.open.odds);
    if (rawVal !== undefined && rawVal !== null && rawVal !== "") {
      return Number(rawVal);
    }
  }

  return "";
}

function snapshotOddsForWeekImpl_(year, week, isOpening, isClosing) {
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
    if (!odds.details && odds.overUnder == null && !odds.awayTeamOdds && !odds.homeTeamOdds && !odds.moneyline) return;

    const provider = odds.provider || {};
    const snapshotId = event.id + "-" + Utilities.formatDate(pulledAt, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");

    rows.push([
      snapshotId, event.id || "", year, week, matchup,
      event.date ? new Date(event.date) : "", pulledAt,
      provider.name || provider.displayName || provider.nameDisplay || "",
      odds.details || "", odds.overUnder ?? "",
      extractMoneyline_(odds, "away"), extractMoneyline_(odds, "home"),
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

  // Tuesday's opening pull is when the week flips: refresh Home's results
  // table (now the week just finished) and drop the old Parlay card.
  if (isOpening) {
    try { getHomeCards_(year, week); } catch (err) { Logger.log("home cards warm failed: " + err.message); }
  }

  if (isWeekComplete_(year, week)) {
    Logger.log("All picks already in for " + year + " Week " + week + " - skipping this scheduled odds pull.");
    return { status: "skipped", reason: "picks already complete" };
  }

  return snapshotOddsForWeek_(year, week, isOpening, isClosing);
}

// The 4 actual trigger targets - install these via installOddsTriggers()
function snapshotOddsTue6am() { return scheduledOddsPull_(true, false); }   // opening line
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
  const handlerNames = ["snapshotOddsTue6am", "snapshotOddsTue4pm", "snapshotOddsWed6am", "snapshotOddsWed3pm"];

  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("snapshotOddsTue6am").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(6).nearMinute(0).create();
  ScriptApp.newTrigger("snapshotOddsTue4pm").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(16).nearMinute(0).create();
  ScriptApp.newTrigger("snapshotOddsWed6am").timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(6).nearMinute(0).create();
  ScriptApp.newTrigger("snapshotOddsWed3pm").timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(15).nearMinute(0).create();

  Logger.log("Odds pull triggers installed: Tue 6am, Tue 4pm, Wed 6am, Wed 3pm (script timezone: " + Session.getScriptTimeZone() + ")");

  return { status: "ok", timezone: Session.getScriptTimeZone() };
}

/**
 * Removes all 4 scheduled odds-pull triggers, if you ever want
 * to pause or stop this automation.
 */
function removeOddsTriggers() {
  const handlerNames = ["snapshotOddsTue6am", "snapshotOddsTue4pm", "snapshotOddsWed6am", "snapshotOddsWed3pm"];
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



// ------------------------------------------------------------
// ODDS PULL WRAPPER - the public name is unchanged; it runs the pull,
// bumps the cache data version (new odds change what the Games and
// Parlay screens show), and rebuilds that week's cached screens right
// away so nobody has to wait on the first load after a pull.
// ------------------------------------------------------------
function snapshotOddsForWeek_(year, week, isOpening, isClosing) {
  const result = runThenBumpVersion_(snapshotOddsForWeekImpl_, [year, week, isOpening, isClosing]);
  if (result && result.status === "ok" && result.rowsAdded > 0) {
    warmWeekCacheSafe_(year, week);
  }
  return result;
}


// ------------------------------------------------------------
// CACHE WARM-UP TRIGGER - every 5 minutes, rebuilds the current
// week's cached Games and Parlay responses if they aren't already
// cached (when they are, this is just two quick cache lookups). Keeps
// the cache warm after pick submissions, after entries expire, and
// across the week, so the first person to open the app rarely pays
// for a rebuild. Run installCacheWarmTrigger() once; safe to re-run.
// ------------------------------------------------------------
function installCacheWarmTrigger() {
  removeCacheWarmTrigger();
  ScriptApp.newTrigger("warmCurrentWeekCache").timeBased().everyMinutes(5).create();
  return { status: "ok", message: "Cache warm-up installed - runs every 5 minutes." };
}

function removeCacheWarmTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "warmCurrentWeekCache") ScriptApp.deleteTrigger(t);
  });
  return { status: "ok" };
}

function warmCurrentWeekCache() {
  const year = getCurrentSeasonYear();
  if (!year) return { status: "ok", skipped: "no current season" };
  const week = getDefaultWeek(year);
  if (!week) return { status: "ok", skipped: "no current week" };
  try { getWeeksWithStatus(year); getYears(); } catch (err) { Logger.log("warm lookups failed: " + err.message); }
  warmWeekCacheSafe_(year, week);
  // Closes (and announces) a runback vote promptly once its deadline passes,
  // instead of waiting for someone to open Home. Cheap: cached and flag-guarded.
  try { checkAndNotifyRunbackStatus_(year, week); } catch (err) { Logger.log("runback check failed: " + err.message); }
  return { status: "ok", year: year, week: week };
}


// ======================================================
// GROUP TEXTS: WEEK OPEN / PARLAY PLACED / RUNBACK PLACED
// All use the "parlay" category (SendParlayInfo + the SendText master
// switch). Which of them also post to Home is set right here:
// ======================================================
const POST_KICKOFF_TO_HOME_ = false;        // "Week N is open" - Home already shows the "you haven't picked" reminder
const POST_PARLAY_PLACED_TO_HOME_ = true;   // a lasting record of the bet
const POST_RUNBACK_PLACED_TO_HOME_ = true;
const KICKOFF_EARLIEST_HOUR_ = 8;           // don't send the "week is open" text before 8am script time

function formatMoney_(n) {
  return "$" + String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function signedMl_(ml) {
  const r = Math.round(Number(ml) || 0);
  return (r > 0 ? "+" : "") + String(r).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/**
 * Texts the group once per week the first time the pick window is open
 * (after KICKOFF_EARLIEST_HOUR_). Called from checkPickReminders.
 */
function checkWeekKickoff_(year, week) {
  const props = PropertiesService.getScriptProperties();
  const flagKey = "KICKOFF_SENT_" + year + "_" + week;
  if (props.getProperty(flagKey)) return { status: "ok", alreadySent: true };

  if (new Date().getHours() < KICKOFF_EARLIEST_HOUR_) return { status: "ok", tooEarly: true };

  // Only on the day the window opens (Tuesday). Without this, installing or
  // re-enabling mid-week would text "Week N is open" for a week that's
  // already well under way.
  if (Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE") !== "Tuesday") {
    return { status: "ok", notTuesday: true };
  }

  const games = sheetToObjects_(SHEET_SCHEDULE).filter(r =>
    Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week) &&
    (r.RowType === "GAME" || r.RowType === "PRESEASON"));
  if (!games.length) return { status: "ok", noGames: true };

  const win = computePickWindow_(games);
  if (!win || !win.isOpen) return { status: "ok", windowNotOpen: true };

  if (!claimOnce_(flagKey)) return { status: "ok", alreadySent: true };

  let link = "";
  try { link = ScriptApp.getService().getUrl() || ""; } catch (err) { link = ""; }

  const lines = [
    "🏈 Week " + week + " is open! " + games.length + " games - picks close " + win.wednesday + "."
  ];
  if (link) lines.push(link);
  const body = lines.join("\n");

  const results = sendSmsToAllUsers_(body, "parlay");
  if (POST_KICKOFF_TO_HOME_) postSystemNotice_("Week " + week + " is open - picks close " + win.wednesday + ".");
  Logger.log("Week " + year + "/" + week + " kickoff text sent to " + results.length + " users.");
  return { status: "ok", sent: results };
}

/** Texts the group (and posts to Home) the first time a week's Parlay is recorded as placed. */
function notifyParlayPlaced_(year, week, placedBy) {
  const props = PropertiesService.getScriptProperties();
  const flagKey = "PARLAY_PLACED_NOTIFIED_" + year + "_" + week;
  if (!claimOnce_(flagKey)) return { status: "ok", alreadyNotified: true };

  let odds = null;
  try { odds = JSON.parse(getParlayWeekData(year, week)).parlayOdds; } catch (err) { odds = null; }

  const names = getUserDisplayNames_();
  const who = names[placedBy] || placedBy;
  const lines = ["🎟️ Week " + week + " Parlay is placed" + (who ? " by " + who : "") + "!"];
  if (odds) {
    const stake = Number(odds.totalPayout) - Number(odds.totalProfit);
    lines.push(odds.gamesIncluded + " legs · " + signedMl_(odds.combinedMoneyline) + " · " +
      formatMoney_(stake) + " in, pays " + formatMoney_(odds.totalPayout) + " if it hits.");
  }
  const body = lines.join("\n");

  const results = sendSmsToAllUsers_(body, "parlay");
  if (POST_PARLAY_PLACED_TO_HOME_) postSystemNotice_(body);
  return { status: "ok", sent: results };
}

/** Same, for the runback bet. */
function notifyRunbackPlaced_(year, week, participants, madeBy) {
  const props = PropertiesService.getScriptProperties();
  const flagKey = "RUNBACK_PLACED_NOTIFIED_" + year + "_" + week;
  if (!claimOnce_(flagKey)) return { status: "ok", alreadyNotified: true };

  let odds = null;
  try { odds = getRunbackStatus(null, year, week).runbackOdds; } catch (err) { odds = null; }

  const names = getUserDisplayNames_();
  const who = names[madeBy] || madeBy;
  const lines = ["🔁 Week " + week + " Runback is placed" + (who ? " by " + who : "") + "!"];
  if (participants) lines.push("In: " + participants);
  if (odds) lines.push(odds.gamesIncluded + " legs · " + signedMl_(odds.combinedMoneyline));
  const body = lines.join("\n");

  const results = sendSmsToAllUsers_(body, "parlay");
  if (POST_RUNBACK_PLACED_TO_HOME_) postSystemNotice_(body);
  return { status: "ok", sent: results };
}


// ======================================================
// REPLY BY TEXT (inbound SMS)
// Twilio posts each incoming text to this web app's doPost. People can:
//   - text team abbreviations to PICK games (while the pick window is open)
//     or to cast TIE-BREAK votes (for games that are currently tied)
//   - STATUS        where they and the group stand
//   - RB YES / RB NO   vote on an open runback
//   - COMMANDS      the list (Twilio itself answers HELP, so it isn't used)
// Replies are sent back through the normal Twilio send, so the HTTP reply
// to Twilio is deliberately empty.
//
// Security: Apps Script can't verify Twilio's signature header, so the
// webhook URL carries a secret token (made by setupSmsWebhook) and a text is
// only acted on if it comes from a phone number on an active user.
// ======================================================
const SMS_WEBHOOK_TOKEN_PROP_ = "SMS_WEBHOOK_TOKEN";

// Reply-by-text is ON for runback votes only (YES / NO). Picks, tie-break
// votes and STATUS are built but switched off until you want them.
const SMS_PICKS_AND_STATUS_ENABLED_ = false;

/** Run once from the function dropdown; the log shows the URL to paste into Twilio. */
function setupSmsWebhook() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty(SMS_WEBHOOK_TOKEN_PROP_);
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, "");
    props.setProperty(SMS_WEBHOOK_TOKEN_PROP_, token);
  }
  const base = ScriptApp.getService().getUrl();
  const url = base + (base.indexOf("?") === -1 ? "?" : "&") + "t=" + token;
  Logger.log("Paste this into Twilio (Phone Number > Messaging > 'A message comes in' > Webhook, HTTP POST):");
  Logger.log(url);
  return { status: "ok", webhookUrl: url };
}

// Every inbound text is recorded on the "SmsLog" sheet so you can see that
// replies are arriving and what the app did with them. Messages from numbers
// that aren't on the Users sheet (for example replies meant for another app
// sharing the Twilio number) are logged WITHOUT their text.
const SHEET_SMS_LOG_ = "SmsLog";

function logInboundSms_(fromPhone, userName, bodyText, outcome) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_SMS_LOG_);
    if (!sh) {
      sh = ss.insertSheet(SHEET_SMS_LOG_);
      sh.getRange(1, 1, 1, 5).setValues([["ReceivedAt", "From", "User", "Text", "Outcome"]]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
    const digits = String(fromPhone || "").replace(/\D/g, "");
    const shown = userName ? ("***" + digits.slice(-4)) : ("***" + digits.slice(-4));
    sh.appendRow([new Date(), shown, userName || "(not a league number)", userName ? String(bodyText || "").slice(0, 200) : "", String(outcome || "").slice(0, 200)]);
  } catch (err) {
    Logger.log("logInboundSms_ failed: " + err.message);
  }
}

function doPost(e) {
  const empty = ContentService.createTextOutput("<Response></Response>")
    .setMimeType(ContentService.MimeType.XML);
  try {
    const p = (e && e.parameter) || {};
    const expected = PropertiesService.getScriptProperties().getProperty(SMS_WEBHOOK_TOKEN_PROP_);
    if (!expected || p.t !== expected) {
      Logger.log("Inbound SMS ignored: missing or wrong token.");
      logInboundSms_(p.From, null, "", "ignored: missing or wrong token");
      return empty;
    }

    // Twilio can deliver the same message twice; handle each MessageSid once.
    const sid = p.MessageSid || p.SmsSid || "";
    if (sid) {
      try {
        const cache = CacheService.getScriptCache();
        if (cache.get("sms:" + sid)) return empty;
        cache.put("sms:" + sid, "1", 21600);
      } catch (err) { /* no cache - carry on */ }
    }

    const from = p.From || "";
    const reply = handleInboundSms_(from, p.Body || "");
    const sender = findUserByPhone_(from);
    logInboundSms_(from, sender ? sender.UserName : null, p.Body || "",
      !sender ? "ignored: unknown number" : (reply ? "replied: " + reply.replace(/\s+/g, " ") : "no reply"));
    if (reply) {
      const r = sendTwilioSms_(from, reply);
      if (r.status !== "ok") Logger.log("Reply to " + from + " failed: " + JSON.stringify(r));
    }
  } catch (err) {
    Logger.log("doPost (inbound SMS) failed: " + err.message);
  }
  return empty;
}

/**
 * Tests the reply-by-text path WITHOUT Twilio: runs exactly what an incoming
 * text would run, as if USER had texted TEXT from their own number, logs it
 * on the SmsLog sheet, and sends nothing. NOTE it is real - a YES/NO here
 * records an actual runback vote (change it again from the Home card).
 * Fill in the two values, run from the function dropdown, check the log.
 */
function DEBUG_testTextReply() {
  const USER = "PUT_YOUR_USERNAME_HERE";
  const TEXT = "YES";
  const u = sheetToObjects_(SHEET_USERS).find(x => x.UserName === USER);
  if (!u) { Logger.log("User not found: " + USER); return { status: "error", message: "User not found" }; }
  const reply = handleInboundSms_(u.Phone, TEXT);
  logInboundSms_(u.Phone, USER, TEXT, "TEST (no SMS sent) - " + (reply ? "would reply: " + reply.replace(/\s+/g, " ") : "no reply"));
  Logger.log("Would reply: " + reply);
  return { status: "ok", reply: reply };
}

function findUserByPhone_(phone) {
  const target = normalizePhone_(phone);
  if (!target) return null;
  return sheetToObjects_(SHEET_USERS).find(u =>
    (u.Active === true || u.Active === "TRUE" || u.Active === "true") &&
    u.UserName !== PARLAY_USER && normalizePhone_(u.Phone) === target) || null;
}

function smsHelpText_() {
  return [
    "Parlay texts:",
    "• RB YES or RB NO - vote when a runback vote is open",
    "• COMMANDS - this list"
  ].join("\n");
}

function smsGamesForWeek_(year, week) {
  return sheetToObjects_(SHEET_SCHEDULE)
    .filter(r => Number(r.SeasonYear) === Number(year) && String(r.Week) === String(week) &&
      (r.RowType === "GAME" || r.RowType === "PRESEASON"))
    .sort((a, b) => new Date(a.EventDate) - new Date(b.EventDate));
}

// userName -> Set of EventIDs with a real pick, for the week
function smsPickedByUser_(year, week) {
  const byUser = {};
  sheetToObjects_(SHEET_PICKS).forEach(p => {
    if (Number(p.SeasonYear) === Number(year) && String(p.Week) === String(week) &&
        p.PickedTeam && p.PickedTeam !== "SKIP") {
      if (!byUser[p.UserName]) byUser[p.UserName] = new Set();
      byUser[p.UserName].add(String(p.EventID));
    }
  });
  return byUser;
}

function smsList_(items, max) {
  if (items.length <= max) return items.join(", ");
  return items.slice(0, max).join(", ") + " +" + (items.length - max) + " more";
}

function smsMyProgress_(userName, games, byUser) {
  const mine = byUser[userName] || new Set();
  const missing = games.filter(g => !mine.has(String(g.EventID)));
  return { picked: games.length - missing.length, total: games.length, missing: missing };
}

function smsStatusText_(user, year, week) {
  const games = smsGamesForWeek_(year, week);
  const byUser = smsPickedByUser_(year, week);
  const me = smsMyProgress_(user.UserName, games, byUser);
  const names = getUserDisplayNames_();
  const lines = [];

  if (me.missing.length) {
    lines.push("Week " + week + ": you've picked " + me.picked + "/" + me.total + ". Missing: " +
      smsList_(me.missing.map(g => g.Matchup), 4));
  } else {
    lines.push("Week " + week + ": you're all picked (" + me.total + "/" + me.total + ").");
  }

  const waiting = getUsers().filter(u => smsMyProgress_(u, games, byUser).missing.length > 0);
  lines.push(waiting.length ? "Waiting on: " + waiting.map(u => names[u] || u).join(", ") : "Everyone has picked.");

  try {
    const tied = getTiedGamesForWeek(year, week);
    if (tied.length) lines.push("Tie-break open: " + smsList_(tied.map(g => g.Matchup), 3) + " - text a team to vote.");
  } catch (err) { /* status is best-effort */ }

  try {
    const rb = getRunbackStatus(user.UserName, year, week);
    if (rb.eligible && rb.open) {
      lines.push("Runback vote open" + (rb.userVote ? " (you voted " + rb.userVote + ")" : "") + " - text RB YES or RB NO.");
    }
  } catch (err) { /* best-effort */ }

  return lines.join("\n");
}

function smsRunbackVote_(user, year, week, vote, silentIfClosed) {
  const st = getRunbackStatus(user.UserName, year, week);
  if (!st.eligible || st.decided || !st.open) return silentIfClosed ? null : "No runback vote is open right now.";
  const res = castRunbackVote(user.UserName, year, week, vote);
  if (!res || res.status !== "ok") return "Couldn't record that vote" + (res && res.message ? ": " + res.message : ".");
  const after = getRunbackStatus(user.UserName, year, week);
  return "🔁 Runback vote saved: " + vote.toUpperCase() + ". Now " + after.yesCount + " Yes / " + after.noCount +
    " No (need " + after.decidedThreshold + ").";
}

/**
 * Turns one incoming text into an action and returns the reply text, or
 * null to stay silent (unknown number). Never throws.
 */
function handleInboundSms_(fromPhone, bodyText) {
  try {
    const user = findUserByPhone_(fromPhone);
    if (!user) return null;

    const year = getCurrentSeasonYear();
    const week = year ? getDefaultWeek(year) : null;
    if (!year || !week) return "There's no active week right now.";

    const upper = String(bodyText || "").trim().toUpperCase().replace(/[.!\s]+$/, "");
    if (!upper) return null;

    if (upper === "COMMANDS") return smsHelpText_();

    // RB YES / RB NO. An "RB YES" with no vote open gets told so; anything
    // else that isn't recognized stays silent.
    // The RB / RUNBACK prefix is required, so a plain YES / NO meant for some
    // other app on the same number is never mistaken for a runback vote.
    const rb = /^(RB|RUNBACK)\s*(YES|Y|NO|N)$/.exec(upper);
    if (rb) return smsRunbackVote_(user, year, week, rb[2].charAt(0) === "Y" ? "Yes" : "No", false);

    if (!SMS_PICKS_AND_STATUS_ENABLED_) return null;
    if (upper === "?" || upper === "MENU") return smsHelpText_();
    if (upper === "STATUS") return smsStatusText_(user, year, week);

    // Otherwise: a list of team abbreviations, optionally led by PICK/PICKS/VOTE/TIE
    let tokens = upper.split(/[^A-Z0-9]+/).filter(Boolean);
    if (tokens.length && ["PICK", "PICKS", "VOTE", "TIE"].indexOf(tokens[0]) !== -1) tokens = tokens.slice(1);
    if (!tokens.length) return "Text team abbreviations to pick (e.g. MIN BUF NE), STATUS, or COMMANDS.";

    const games = smsGamesForWeek_(year, week);
    const byTeam = {};
    games.forEach(g => {
      [g.AwayShort, g.HomeShort].forEach(t => { if (t) byTeam[String(t).toUpperCase()] = { game: g, team: String(t) }; });
    });

    const unknown = tokens.filter(t => !byTeam[t]);
    if (unknown.length) {
      return "Didn't recognize " + unknown.join(", ") + ". Use team abbreviations from this week's games (like MIN or BUF), or text COMMANDS.";
    }

    // One team per game
    const chosenByEvent = {};
    for (let i = 0; i < tokens.length; i++) {
      const hit = byTeam[tokens[i]];
      const id = String(hit.game.EventID);
      if (chosenByEvent[id] && chosenByEvent[id].team !== hit.team) {
        return "You sent both " + chosenByEvent[id].team + " and " + hit.team + " (same game: " + hit.game.Matchup + "). Nothing saved - resend with one.";
      }
      chosenByEvent[id] = hit;
    }

    let tiedIds = new Set();
    try { tiedIds = new Set(getTiedGamesForWeek(year, week).map(r => String(r.EventID))); } catch (err) { tiedIds = new Set(); }

    const asItem = hit => ({ eventId: hit.game.EventID, matchup: hit.game.Matchup, pickedTeam: hit.team });
    const hits = Object.keys(chosenByEvent).map(id => chosenByEvent[id]);
    const tieItems = hits.filter(h => tiedIds.has(String(h.game.EventID))).map(asItem);
    const pickHits = hits.filter(h => !tiedIds.has(String(h.game.EventID)));

    const win = computePickWindow_(games);
    const windowOpen = !!(win && win.isOpen);
    const pickItems = windowOpen ? pickHits.map(asItem) : [];

    const parts = [];

    if (pickItems.length) {
      const res = submitWeekPicks(year, week, user.UserName, pickItems);
      if (!res || res.status !== "ok") return "Couldn't save your picks right now - try the app.";
      const me = smsMyProgress_(user.UserName, games, smsPickedByUser_(year, week));
      parts.push("✅ Saved: " + pickItems.map(i => i.pickedTeam).join(", ") + ".");
      parts.push(me.missing.length
        ? "You've picked " + me.picked + "/" + me.total + ". Left: " + smsList_(me.missing.map(g => g.Matchup), 4)
        : "You're all set for Week " + week + "! 🎉");
    } else if (pickHits.length) {
      parts.push("Picks for Week " + week + " are closed, so " + pickHits.map(h => h.team).join(", ") + " wasn't saved.");
    }

    if (tieItems.length) {
      const res = submitTieBreakVotes(year, week, user.UserName, tieItems);
      if (!res || res.status !== "ok") return parts.concat(["Couldn't save your tie-break vote right now - try the app."]).join("\n");
      parts.push("✅ Tie-break vote saved: " + tieItems.map(i => i.pickedTeam).join(", ") + ".");
    }

    return parts.join("\n");
  } catch (err) {
    Logger.log("handleInboundSms_ failed: " + err.message);
    return "Sorry, something went wrong on our end - try the app.";
  }
}


// ======================================================
// SCHEDULED FINAL-SCORE PULLS
// The three 5am triggers (Fri / Mon / Tue) call these. The functions that
// used to sit behind those trigger names were no longer in the project, so
// every run failed with "function not found"; defining them again under the
// SAME names means the existing triggers start working with no re-install.
// Each pull covers the current week plus the one before it (Tuesday's pull
// is what finishes Monday night's games, right as the week flips), then
// re-warms the cached screens, Home's results table included.
// ======================================================
function pullFinalScoresScheduled_() {
  const year = getCurrentSeasonYear();
  if (!year) return { status: "skipped", reason: "no current season" };
  const week = getDefaultWeek(year);
  if (!week) return { status: "skipped", reason: "no current week" };

  const weeks = [String(week)];
  const prev = Number(week) - 1;
  if (prev >= 1) weeks.push(String(prev));

  const out = [];
  weeks.forEach(w => {
    try {
      const res = pullFinalScores(year, w);
      out.push({ week: w, result: res });
      Logger.log("Final scores pulled for " + year + " week " + w + ": " + JSON.stringify(res));
    } catch (err) {
      out.push({ week: w, error: err.message });
      Logger.log("Final score pull failed for " + year + " week " + w + ": " + err.message);
    }
    try { warmWeekCacheSafe_(year, w); } catch (err) { /* warm-up must never fail the pull */ }
  });
  return { status: "ok", year: year, pulled: out };
}

function pullFinalScoresFri5am() { return pullFinalScoresScheduled_(); }
function pullFinalScoresMon5am() { return pullFinalScoresScheduled_(); }
function pullFinalScoresTue5am() { return pullFinalScoresScheduled_(); }

/**
 * Run from the function dropdown: lists every trigger on the project and
 * flags any whose function no longer exists (those fail on every run).
 */
function DEBUG_listTriggers() {
  const g = (function () { return this; })();
  const rows = ScriptApp.getProjectTriggers().map(t => {
    const name = t.getHandlerFunction();
    return { handler: name, exists: typeof g[name] === "function", type: String(t.getEventType()) };
  });
  rows.forEach(r => Logger.log((r.exists ? "OK       " : "MISSING  ") + r.handler + "  (" + r.type + ")"));
  return rows;
}
