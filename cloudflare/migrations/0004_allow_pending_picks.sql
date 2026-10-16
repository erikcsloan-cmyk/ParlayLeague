-- Parlay League D1 Database
-- Migration 0004: Allow pending picks
-- Preserve historical picks with no selected team.

CREATE TABLE picks_new (
    pick_id TEXT PRIMARY KEY,
    event_id TEXT NOT NULL,
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    matchup TEXT,
    user_name TEXT NOT NULL,
    picked_team TEXT,
    submitted_at TEXT,

    FOREIGN KEY (event_id) REFERENCES nfl_games(event_id),
    FOREIGN KEY (user_name) REFERENCES users(user_name),

    UNIQUE (event_id, user_name)
);

-- Copy any existing records into the updated structure.
INSERT INTO picks_new (
    pick_id,
    event_id,
    season_year,
    week,
    matchup,
    user_name,
    picked_team,
    submitted_at
)
SELECT
    pick_id,
    event_id,
    season_year,
    week,
    matchup,
    user_name,
    picked_team,
    submitted_at
FROM picks;

DROP TABLE picks;

ALTER TABLE picks_new RENAME TO picks;

CREATE INDEX idx_picks_season_week
ON picks (season_year, week);

