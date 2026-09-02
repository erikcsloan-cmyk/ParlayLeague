function testNFLHistoricalGameTimeOdds() {

  const testDates = [
    {
      season: 2024,
      date: "20240905"
    },
    {
      season: 2025,
      date: "20250904"
    }
  ];

  testDates.forEach(test => {

    const url =
      "https://site.web.api.espn.com/apis/site/v2/sports/" +
      "football/nfl/scoreboard" +
      "?dates=" + test.date;

    Logger.log("====================================");
    Logger.log("SEASON: " + test.season);
    Logger.log("DATE: " + test.date);
    Logger.log("URL: " + url);
    Logger.log("====================================");

    const response =
      UrlFetchApp.fetch(url, {
        method: "get",
        muteHttpExceptions: true
      });

    const code =
      response.getResponseCode();

    Logger.log("HTTP: " + code);

    if (code !== 200) {
      Logger.log(
        response
          .getContentText()
          .substring(0, 2000)
      );
      return;
    }

    const data =
      JSON.parse(
        response.getContentText()
      );

    const events =
      data.events || [];

    Logger.log(
      "EVENT COUNT: " +
      events.length
    );

    events.forEach(event => {

      const competition =
        event.competitions?.[0] || {};

      const competitors =
        competition.competitors || [];

      const away =
        competitors.find(
          c => c.homeAway === "away"
        ) || {};

      const home =
        competitors.find(
          c => c.homeAway === "home"
        ) || {};

      const awayTeam =
        away.team?.abbreviation || "";

      const homeTeam =
        home.team?.abbreviation || "";

      Logger.log("------------------------------------");

      Logger.log(
        "EVENT ID: " +
        event.id
      );

      Logger.log(
        "MATCHUP: " +
        awayTeam +
        " @ " +
        homeTeam
      );

      Logger.log(
        "GAME START: " +
        event.date
      );

      Logger.log(
        "STATUS: " +
        (
          competition.status
            ?.type
            ?.description || ""
        )
      );

      Logger.log("ODDS RAW:");

      Logger.log(
        JSON.stringify(
          competition.odds || [],
          null,
          2
        )
      );
    });
  });
}

function testCoversHistoricalNFL() {

  const url =
    "https://www.covers.com/sportsoddshistory/nfl-game-odds/";

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
        "AppleWebKit/537.36 Chrome/131.0 Safari/537.36"
    }
  });

  const code =
    response.getResponseCode();

  const body =
    response.getContentText();

  Logger.log("HTTP CODE: " + code);
  Logger.log("BODY LENGTH: " + body.length);

  Logger.log(
    body.substring(0, 5000)
  );
 }

function testCoversNFL2025Season() {

  const url =
    "https://www.covers.com/sportsoddshistory/" +
    "nfl-game-season/?y=2025";

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
        "AppleWebKit/537.36 Chrome/131.0 Safari/537.36"
    }
  });

  const code = response.getResponseCode();
  const body = response.getContentText();

  Logger.log("HTTP CODE: " + code);
  Logger.log("BODY LENGTH: " + body.length);

  if (code !== 200) {
    Logger.log(body.substring(0, 3000));
    return;
  }

  // Look for known teams from Week 1
  const searchTerms = [
    "Dallas Cowboys",
    "Philadelphia Eagles",
    "Sep 4",
    "Sep 5",
    "2025"
  ];

  searchTerms.forEach(term => {

    const index = body.indexOf(term);

    Logger.log(
      "SEARCH: " +
      term +
      " | INDEX: " +
      index
    );

    if (index >= 0) {

      Logger.log(
        body.substring(
          Math.max(0, index - 1500),
          Math.min(body.length, index + 4000)
        )
      );
    }
  });
 }
function pullCoversNFLHistoricalOdds() {

  const seasons = [2024, 2025];

  const sheetName = "API_NFL_OddsHistory";

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet =
      ss.insertSheet(sheetName);
  }


  // ========================================
  // REQUIRED HEADERS
  // ========================================

  const requiredHeaders = [
    "OddsSnapshotID",
    "EventID",
    "SeasonYear",
    "Week",
    "Matchup",
    "GameDateTime",
    "PulledAt",
    "OddsProvider",
    "Spread",
    "OverUnder",
    "AwayMoneyline",
    "HomeMoneyline",
    "IsOpeningSnapshot",
    "IsClosingSnapshot"
  ];


  // ========================================
  // CREATE HEADERS IF SHEET EMPTY
  // ========================================

  if (sheet.getLastRow() === 0) {

    sheet
      .getRange(
        1,
        1,
        1,
        requiredHeaders.length
      )
      .setValues([requiredHeaders])
      .setFontWeight("bold");

  }


  // ========================================
  // READ ACTUAL SHEET HEADERS
  // ========================================

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(h => String(h).trim());


  // ========================================
  // VERIFY REQUIRED HEADERS
  // ========================================

  requiredHeaders.forEach(header => {

    if (!headers.includes(header)) {

      throw new Error(
        "API_NFL_OddsHistory missing required header: " +
        header
      );

    }

  });


  const records = [];


  // ========================================
  // ESPN EVENT LOOKUP
  // ========================================

  const espnLookup =
    buildNFLHistoricalEventLookup_();


  // ========================================
  // LOOP SEASONS
  // ========================================

  seasons.forEach(season => {

    Logger.log(
      "================================"
    );

    Logger.log(
      "COVERS NFL ODDS " + season
    );

    Logger.log(
      "================================"
    );


    const url =
      "https://www.covers.com/" +
      "sportsoddshistory/" +
      "nfl-game-season/?y=" +
      season;


    const response =
      UrlFetchApp.fetch(
        url,
        {
          method: "get",

          muteHttpExceptions:
            true,

          headers: {

            "User-Agent":
              "Mozilla/5.0 " +
              "(Windows NT 10.0; Win64; x64) " +
              "AppleWebKit/537.36 " +
              "Chrome/131.0 Safari/537.36"

          }
        }
      );


    const code =
      response.getResponseCode();


    Logger.log(
      "HTTP CODE: " + code
    );


    if (code !== 200) {

      Logger.log(
        response
          .getContentText()
          .substring(0, 2000)
      );

      return;
    }


    const html =
      response.getContentText();


    const seasonRecords =
      parseCoversNFLSeasonHtml_(
        html,
        season,
        espnLookup
      );


    Logger.log(
      season +
      " RECORDS PARSED: " +
      seasonRecords.length
    );


    records.push(
      ...seasonRecords
    );


    Utilities.sleep(300);

  });


  // ========================================
  // REMOVE EXISTING COVERS CLOSING RECORDS
  // BEFORE REWRITING THEM
  // ========================================

  removeExistingCoversClosingOdds_(
    sheet,
    headers,
    [2024, 2025]
  );


  // ========================================
  // MAP RECORDS TO SHEET BY HEADER NAME
  // ========================================

  const rows =
    records.map(record => {

      return headers.map(
        header =>
          record[header] ?? ""
      );

    });


  // ========================================
  // APPEND
  // ========================================

  if (rows.length) {

    sheet
      .getRange(
        sheet.getLastRow() + 1,
        1,
        rows.length,
        headers.length
      )
      .setValues(rows);

  }


  sheet.setFrozenRows(1);

  sheet.autoResizeColumns(
    1,
    headers.length
  );


  Logger.log(
    "================================"
  );

  Logger.log(
    "COVERS HISTORICAL ODDS COMPLETE"
  );

  Logger.log(
    "TOTAL CLOSING LINES WRITTEN: " +
    records.length
  );

  Logger.log(
    "================================"
  );

 }

function parseCoversNFLSeasonHtml_(
  html,
  season,
  espnLookup
 ) {

  const records = [];


  // ========================================
  // WEEK TABLES
  // ========================================

  const weekRegex =
    /<h3>\s*(\d{4})\s+Regular Season\s*-\s*Week\s+(\d+)<\/h3>[\s\S]*?<tbody>([\s\S]*?)<\/tbody>/gi;


  let weekMatch;


  while (
    (weekMatch = weekRegex.exec(html)) !== null
  ) {

    const parsedSeason =
      Number(
        weekMatch[1]
      );

    const week =
      Number(
        weekMatch[2]
      );


    if (
      parsedSeason !== season
    ) {
      continue;
    }


    const tbody =
      weekMatch[3];


    Logger.log(
      `Parsing Covers ${season} Week ${week}`
    );


    // ======================================
    // GAME ROWS
    // ======================================

    const rowRegex =
      /<tr[^>]*>([\s\S]*?)<\/tr>/gi;


    let rowMatch;


    while (
      (rowMatch = rowRegex.exec(tbody)) !== null
    ) {

      const cells =
        extractHtmlTableCells_(
          rowMatch[1]
        );


      if (
        cells.length < 10
      ) {
        continue;
      }


      const gameDateText =
        cleanHtmlText_(
          cells[1]
        );


      const gameTimeText =
        cleanHtmlText_(
          cells[2]
        );


      const favoriteLocation =
        cleanHtmlText_(
          cells[3]
        );


      const favorite =
        cleanHtmlText_(
          cells[4]
        );


      const scoreText =
        cleanHtmlText_(
          cells[5]
        );


      const spreadText =
        cleanHtmlText_(
          cells[6]
        );


      const underdogLocation =
        cleanHtmlText_(
          cells[7]
        );


      const underdog =
        cleanHtmlText_(
          cells[8]
        );


      const totalText =
        cleanHtmlText_(
          cells[9]
        );


      const notes =
        cells.length > 10
          ? cleanHtmlText_(cells[10])
          : "";


      if (
        !gameDateText ||
        !favorite ||
        !underdog
      ) {
        continue;
      }


      // ====================================
      // TEAM SHORT NAMES
      // ====================================

      const favoriteShort =
        coversNFLTeamToShort_(
          favorite
        );


      const underdogShort =
        coversNFLTeamToShort_(
          underdog
        );


      if (
        !favoriteShort ||
        !underdogShort
      ) {

        Logger.log(
          "TEAM MAP FAILED: " +
          favorite +
          " / " +
          underdog
        );

        continue;
      }


      // ====================================
      // DETERMINE AWAY / HOME
      // ====================================

      let awayShort = "";
      let homeShort = "";


      /*
       * Covers placement:
       *
       * @ before favorite:
       * underdog @ favorite
       *
       * @ before underdog:
       * favorite @ underdog
       *
       * N = neutral site
       */


      if (
        favoriteLocation === "@"
      ) {

        awayShort =
          underdogShort;

        homeShort =
          favoriteShort;

      } else if (
        underdogLocation === "@"
      ) {

        awayShort =
          favoriteShort;

        homeShort =
          underdogShort;

      } else {

        /*
         * Neutral-site games.
         *
         * Keep Covers listed ordering:
         * Favorite vs Underdog.
         *
         * ESPN lookup below will help
         * identify the actual away/home
         * combination when possible.
         */

        awayShort =
          favoriteShort;

        homeShort =
          underdogShort;

      }


      // ====================================
      // MATCHUP
      // ====================================

      let matchup =
        `${awayShort} @ ${homeShort}`;


      // ====================================
      // GAME DATE/TIME
      // Covers times are Eastern
      // ====================================

      const gameDateTime =
        parseCoversNFLDateTimeET_(
          gameDateText,
          gameTimeText
        );


      // ====================================
      // ESPN EVENT ID LOOKUP
      // ====================================

      let eventID = "";


      const lookupKey =
        `${season}|${week}|${matchup}`;


      if (
        espnLookup[lookupKey]
      ) {

        eventID =
          espnLookup[lookupKey];

      } else {

        /*
         * Try reverse orientation,
         * especially useful for
         * neutral-site games.
         */

        const reverseMatchup =
          `${homeShort} @ ${awayShort}`;


        const reverseKey =
          `${season}|${week}|${reverseMatchup}`;


        if (
          espnLookup[reverseKey]
        ) {

          eventID =
            espnLookup[reverseKey];

          matchup =
            reverseMatchup;


          const temp =
            awayShort;

          awayShort =
            homeShort;

          homeShort =
            temp;
        }
      }


      // ====================================
      // CLOSING SPREAD
      // ====================================

      /*
       * Covers gives:
       *
       * W -8
       * L -3
       * P -2
       *
       * Remove ATS result indicator.
       */

      const spreadNumber =
        parseCoversOddsNumber_(
          spreadText
        );


      /*
       * Spread always belongs to
       * the FAVORITE.
       */

      const spread =
        favoriteShort +
        " " +
        spreadNumber;


      // ====================================
      // CLOSING TOTAL
      // ====================================

      const closingTotal =
        parseCoversOddsNumber_(
          totalText
        );


      // ====================================
      // SNAPSHOT ID
      // ====================================

      const snapshotID =
        eventID
          ? `COVERS-CLOSE-${eventID}`
          : (
              "COVERS-CLOSE-" +
              season +
              "-" +
              week +
              "-" +
              awayShort +
              "-" +
              homeShort
            );


      // ====================================
      // BUILD ODDS HISTORY RECORD
      // ====================================

      records.push({

        OddsSnapshotID:
          snapshotID,

        EventID:
          eventID,

        SeasonYear:
          season,

        Week:
          week,

        Matchup:
          matchup,

        GameDateTime:
          gameDateTime,

        PulledAt:
          new Date(),

        OddsProvider:
          "Covers",

        Spread:
          spread,

        OverUnder:
          closingTotal,

        AwayMoneyline:
          "",

        HomeMoneyline:
          "",

        IsOpeningSnapshot:
          false,

        IsClosingSnapshot:
          true

      });

    }

  }


  return records;

 }

function extractHtmlTableCells_(
  rowHtml
 ) {

  const cells = [];

  const cellRegex =
    /<td[^>]*>([\s\S]*?)<\/td>/gi;

  let match;

  while (
    (match = cellRegex.exec(rowHtml)) !== null
  ) {

    cells.push(
      match[1]
    );

  }

  return cells;

 }

function cleanHtmlText_( 
  html
 ) {

  if (!html) {
    return "";
  }

  return String(html)

    .replace(
      /<br\s*\/?>/gi,
      " "
    )

    .replace(
      /<[^>]+>/g,
      ""
    )

    .replace(
      /&nbsp;/gi,
      " "
    )

    .replace(
      /&amp;/gi,
      "&"
    )

    .replace(
      /&#39;/gi,
      "'"
    )

    .replace(
      /&quot;/gi,
      '"'
    )

    .replace(
      /\s+/g,
      " "
    )

    .trim();

 }

function parseCoversOddsNumber_(
  value
 ) {

  const text =
    String(
      value || ""
    ).trim();


  /*
   * Examples:
   *
   * W -8
   * L -3
   * U 47.5
   * O 44
   * P 42
   */

  const match =
    text.match(
      /(-?\d+(?:\.\d+)?)/
    );


  return match
    ? Number(match[1])
    : "";

 }



function parseCoversNFLDateTimeET_(
  dateText,
  timeText
 ) {

  /*
   * Covers:
   * Sep 4, 2025
   * 8:20
   *
   * Treat time as Eastern.
   */

  const combined =
    dateText +
    " " +
    timeText;


  const parsed =
    new Date(
      combined +
      " ET"
    );


  if (
    isNaN(
      parsed.getTime()
    )
  ) {

    return "";
  }


  return parsed;

 }

function buildNFLHistoricalEventLookup_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      "API_NFL_Results"
    );


  if (!sheet) {

    throw new Error(
      "API_NFL_Results sheet not found"
    );

  }


  const data =
    sheet
      .getDataRange()
      .getValues();


  if (
    data.length < 2
  ) {
    return {};
  }


  const headers =
    data[0]
      .map(
        h =>
          String(h).trim()
      );


  const col = {};


  headers.forEach(
    (header, index) => {

      if (header) {
        col[header] =
          index;
      }

    }
  );


  const required = [
    "SeasonYear",
    "Week",
    "EventID",
    "Matchup"
  ];


  required.forEach(
    header => {

      if (
        col[header] === undefined
      ) {

        throw new Error(
          "API_NFL_Results missing required header: " +
          header
        );

      }

    }
  );


  const lookup = {};


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const season =
      row[
        col.SeasonYear
      ];


    const week =
      row[
        col.Week
      ];


    const eventID =
      row[
        col.EventID
      ];


    const matchup =
      row[
        col.Matchup
      ];


    if (
      !season ||
      !week ||
      !eventID ||
      !matchup
    ) {
      continue;
    }


    const key =
      `${season}|${week}|${matchup}`;


    lookup[key] =
      String(eventID);

  }


  return lookup;

 }
function removeExistingCoversClosingOdds_( 
  sheet,
  headers,
  seasons
 ) {

  if (
    sheet.getLastRow() <= 1
  ) {
    return;
  }


  const headerMap = {};


  headers.forEach(
    (header, index) => {

      headerMap[header] =
        index;

    }
  );


  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        headers.length
      )
      .getValues();


  const keptRows =
    data.filter(row => {

      const provider =
        String(
          row[
            headerMap.OddsProvider
          ] || ""
        ).trim();


      const isClosing =
        row[
          headerMap.IsClosingSnapshot
        ];


      const season =
        Number(
          row[
            headerMap.SeasonYear
          ]
        );


      const isCoversHistorical =
        provider === "Covers" &&
        (
          isClosing === true ||
          String(isClosing)
            .toLowerCase() === "true"
        ) &&
        seasons.includes(
          season
        );


      return !isCoversHistorical;

    });


  sheet
    .getRange(
      2,
      1,
      sheet.getMaxRows() - 1,
      headers.length
    )
    .clearContent();


  if (
    keptRows.length
  ) {

    sheet
      .getRange(
        2,
        1,
        keptRows.length,
        headers.length
      )
      .setValues(
        keptRows
      );

  }

 }

 function coversNFLTeamToShort_(teamName) {

  const map = {
    "Arizona Cardinals": "ARI",
    "Atlanta Falcons": "ATL",
    "Baltimore Ravens": "BAL",
    "Buffalo Bills": "BUF",

    "Carolina Panthers": "CAR",
    "Chicago Bears": "CHI",
    "Cincinnati Bengals": "CIN",
    "Cleveland Browns": "CLE",

    "Dallas Cowboys": "DAL",
    "Denver Broncos": "DEN",
    "Detroit Lions": "DET",
    "Green Bay Packers": "GB",

    "Houston Texans": "HOU",
    "Indianapolis Colts": "IND",
    "Jacksonville Jaguars": "JAX",
    "Kansas City Chiefs": "KC",

    "Las Vegas Raiders": "LV",
    "Los Angeles Chargers": "LAC",
    "Los Angeles Rams": "LAR",
    "Miami Dolphins": "MIA",

    "Minnesota Vikings": "MIN",
    "New England Patriots": "NE",
    "New Orleans Saints": "NO",
    "New York Giants": "NYG",

    "New York Jets": "NYJ",
    "Philadelphia Eagles": "PHI",
    "Pittsburgh Steelers": "PIT",
    "San Francisco 49ers": "SF",

    "Seattle Seahawks": "SEA",
    "Tampa Bay Buccaneers": "TB",
    "Tennessee Titans": "TEN",
    "Washington Commanders": "WSH"
  };

  const cleanName =
    String(teamName || "")
      .replace(/\s+/g, " ")
      .trim();

  const shortName =
    map[cleanName] || "";

  if (!shortName) {
    Logger.log(
      "COVERS TEAM NOT MAPPED: [" +
      cleanName +
      "]"
    );
  }

  return shortName;
 }