-- Parlay League D1 Database
-- Migration 0001: Initial schema
-- All development and testing will happen locally first.

CREATE TABLE IF NOT EXISTS users (
    user_name TEXT PRIMARY KEY,
    active INTEGER NOT NULL DEFAULT 1,
    sort_order INTEGER,
    phone TEXT,
    send_text INTEGER NOT NULL DEFAULT 1,
    send_tie_info INTEGER NOT NULL DEFAULT 1,
    send_recap_info INTEGER NOT NULL DEFAULT 1,
    send_still_tied INTEGER NOT NULL DEFAULT 1,
    display_name TEXT,
    short_name TEXT,
    is_admin INTEGER NOT NULL DEFAULT 0,
    send_runback_info INTEGER NOT NULL DEFAULT 1,
    pick_reminder1_enabled INTEGER NOT NULL DEFAULT 0,
    pick_reminder1_day TEXT,
    pick_reminder1_time TEXT,
    pick_reminder2_enabled INTEGER NOT NULL DEFAULT 0,
    pick_reminder2_day TEXT,
    pick_reminder2_time TEXT,
    pick_reminder3_enabled INTEGER NOT NULL DEFAULT 0,
    pick_reminder3_day TEXT,
    pick_reminder3_time TEXT
);

-- NFL games and schedule
-- EventID is the existing identifier from ESPN.

CREATE TABLE IF NOT EXISTS nfl_games (
    event_id TEXT PRIMARY KEY,
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    row_type TEXT,
    event_date TEXT,
    event_time TEXT,
    away_team TEXT,
    away_short TEXT,
    home_team TEXT,
    home_short TEXT,
    matchup TEXT,
    venue TEXT,
    city TEXT,
    state TEXT,
    network TEXT,
    country TEXT
);

CREATE INDEX IF NOT EXISTS idx_nfl_games_season_week
ON nfl_games (season_year, week);

-- Individual user picks
-- One pick per user per NFL event.

CREATE TABLE IF NOT EXISTS picks (
    pick_id TEXT PRIMARY KEY,
    event_id TEXT NOT NULL,
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    matchup TEXT,
    user_name TEXT NOT NULL,
    picked_team TEXT NOT NULL,
    submitted_at TEXT,

    FOREIGN KEY (event_id) REFERENCES nfl_games(event_id),
    FOREIGN KEY (user_name) REFERENCES users(user_name),

    UNIQUE (event_id, user_name)
);

CREATE INDEX IF NOT EXISTS idx_picks_season_week
ON picks (season_year, week);


-- Parlay consensus picks
-- Generated from the league's individual picks.

CREATE TABLE IF NOT EXISTS parlay_picks (
    event_id TEXT PRIMARY KEY,
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    matchup TEXT,
    picked_team TEXT,
    vote_count INTEGER,
    total_votes INTEGER,
    generated_at TEXT,
    resolved_via TEXT,

    FOREIGN KEY (event_id) REFERENCES nfl_games(event_id)
);

CREATE INDEX IF NOT EXISTS idx_parlay_picks_season_week
ON parlay_picks (season_year, week);