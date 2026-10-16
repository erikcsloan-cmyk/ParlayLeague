-- Parlay League D1 Database
-- Migration 0005: Shared game references
-- Supports regular-season and preseason results and odds.

-- 1. Create a shared registry of game EventIDs.

CREATE TABLE IF NOT EXISTS game_references (
    event_id TEXT PRIMARY KEY,
    game_type TEXT NOT NULL
        CHECK (game_type IN ('REGULAR', 'PRESEASON'))
);

-- 2. Register existing regular-season games.

INSERT OR IGNORE INTO game_references (event_id, game_type)
SELECT event_id, 'REGULAR'
FROM nfl_games;

-- 3. Register existing preseason games.

INSERT OR IGNORE INTO game_references (event_id, game_type)
SELECT event_id, 'PRESEASON'
FROM nfl_preseason;

-- 4. Rebuild NFL results with the shared foreign key.

CREATE TABLE nfl_results_new (
    event_id TEXT PRIMARY KEY,
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    matchup TEXT,
    away_score INTEGER,
    home_score INTEGER,
    winning_team TEXT,
    updated_at TEXT,

    FOREIGN KEY (event_id)
        REFERENCES game_references(event_id)
);

INSERT INTO nfl_results_new (
    event_id, season_year, week, matchup,
    away_score, home_score, winning_team, updated_at
)
SELECT
    event_id, season_year, week, matchup,
    away_score, home_score, winning_team, updated_at
FROM nfl_results;

DROP TABLE nfl_results;

ALTER TABLE nfl_results_new RENAME TO nfl_results;

CREATE INDEX IF NOT EXISTS idx_nfl_results_season_week
ON nfl_results (season_year, week);

-- 5. Rebuild NFL odds history with the shared foreign key.

CREATE TABLE nfl_odds_history_new (
    odds_snapshot_id TEXT PRIMARY KEY,
    event_id TEXT NOT NULL,
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    matchup TEXT,
    game_date_time TEXT,
    pulled_at TEXT,
    odds_provider TEXT,
    spread REAL,
    over_under REAL,
    away_moneyline INTEGER,
    home_moneyline INTEGER,
    is_opening_snapshot INTEGER NOT NULL DEFAULT 0,
    is_closing_snapshot INTEGER NOT NULL DEFAULT 0,

    FOREIGN KEY (event_id)
        REFERENCES game_references(event_id)
);

INSERT INTO nfl_odds_history_new (
    odds_snapshot_id, event_id, season_year, week,
    matchup, game_date_time, pulled_at, odds_provider,
    spread, over_under, away_moneyline, home_moneyline,
    is_opening_snapshot, is_closing_snapshot
)
SELECT
    odds_snapshot_id, event_id, season_year, week,
    matchup, game_date_time, pulled_at, odds_provider,
    spread, over_under, away_moneyline, home_moneyline,
    is_opening_snapshot, is_closing_snapshot
FROM nfl_odds_history;

DROP TABLE nfl_odds_history;

ALTER TABLE nfl_odds_history_new RENAME TO nfl_odds_history;

CREATE INDEX IF NOT EXISTS idx_nfl_odds_event
ON nfl_odds_history (event_id);

CREATE INDEX IF NOT EXISTS idx_nfl_odds_season_week
ON nfl_odds_history (season_year, week);