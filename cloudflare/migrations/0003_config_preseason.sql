-- Parlay League D1 Database
-- Migration 0003: Configuration and preseason data
-- Part 1: Application configuration

CREATE TABLE IF NOT EXISTS app_config (
    key TEXT PRIMARY KEY,
    value TEXT
);

-- Part 2: NFL preseason schedule
-- Stored separately from the regular-season schedule.

CREATE TABLE IF NOT EXISTS nfl_preseason (
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
    state TEXT
);

CREATE INDEX IF NOT EXISTS idx_nfl_preseason_season_week
ON nfl_preseason (season_year, week);

-- Part 3: NFL preseason testing
-- Isolated test schedule for preseason development.

CREATE TABLE IF NOT EXISTS nfl_preseason_test (
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
    status TEXT
);

CREATE INDEX IF NOT EXISTS idx_nfl_preseason_test_season_week
ON nfl_preseason_test (season_year, week);

