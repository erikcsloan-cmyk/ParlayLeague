-- Parlay League D1 Database
-- Migration 0002: League features
-- Part 1: NFL game results

CREATE TABLE IF NOT EXISTS nfl_results (
    event_id TEXT PRIMARY KEY,
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    matchup TEXT,
    away_score INTEGER,
    home_score INTEGER,
    winning_team TEXT,
    updated_at TEXT,

    FOREIGN KEY (event_id) REFERENCES nfl_games(event_id)
);

CREATE INDEX IF NOT EXISTS idx_nfl_results_season_week
ON nfl_results (season_year, week);

-- Part 2: NFL odds history
-- Multiple odds snapshots may exist for each game.

CREATE TABLE IF NOT EXISTS nfl_odds_history (
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

    FOREIGN KEY (event_id) REFERENCES nfl_games(event_id)
);

CREATE INDEX IF NOT EXISTS idx_nfl_odds_event
ON nfl_odds_history (event_id);

CREATE INDEX IF NOT EXISTS idx_nfl_odds_season_week
ON nfl_odds_history (season_year, week);

-- Part 3: Tie-break picks
-- Records individual votes used to resolve tied selections.

CREATE TABLE IF NOT EXISTS tie_break_picks (
    tie_break_id TEXT PRIMARY KEY,
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    event_id TEXT NOT NULL,
    matchup TEXT,
    user_name TEXT NOT NULL,
    picked_team TEXT NOT NULL,
    submitted_at TEXT,

    FOREIGN KEY (event_id) REFERENCES nfl_games(event_id),
    FOREIGN KEY (user_name) REFERENCES users(user_name)
);

CREATE INDEX IF NOT EXISTS idx_tie_break_picks_event
ON tie_break_picks (event_id);

CREATE INDEX IF NOT EXISTS idx_tie_break_picks_season_week
ON tie_break_picks (season_year, week);

-- Part 4: Runback votes
-- Tracks each user's vote for a weekly runback.

CREATE TABLE IF NOT EXISTS runback_votes (
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    user_name TEXT NOT NULL,
    vote TEXT,
    voted_at TEXT,

    PRIMARY KEY (season_year, week, user_name),

    FOREIGN KEY (user_name) REFERENCES users(user_name)
);

CREATE INDEX IF NOT EXISTS idx_runback_votes_season_week
ON runback_votes (season_year, week);

-- Part 5: Parlay placements
-- Tracks weekly parlay placement and runback details.

CREATE TABLE IF NOT EXISTS parlay_placements (
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    placed_by TEXT,
    placed_at TEXT,
    runback TEXT,
    runback_participants TEXT,
    runback_made_by TEXT,
    runback_made_at TEXT,

    PRIMARY KEY (season_year, week)
);

CREATE INDEX IF NOT EXISTS idx_parlay_placements_season_week
ON parlay_placements (season_year, week);

-- Part 6: League notices
-- Announcements posted by league members or administrators.

CREATE TABLE IF NOT EXISTS notices (
    notice_id TEXT PRIMARY KEY,
    posted_by TEXT,
    posted_at TEXT,
    text TEXT
);

CREATE INDEX IF NOT EXISTS idx_notices_posted_at
ON notices (posted_at);

-- Part 7: Notice dismissals
-- Tracks announcements dismissed by individual users.

CREATE TABLE IF NOT EXISTS notice_dismissals (
    notice_id TEXT NOT NULL,
    user_name TEXT NOT NULL,
    dismissed_at TEXT,

    PRIMARY KEY (notice_id, user_name),

    FOREIGN KEY (notice_id) REFERENCES notices(notice_id),
    FOREIGN KEY (user_name) REFERENCES users(user_name)
);

CREATE INDEX IF NOT EXISTS idx_notice_dismissals_user
ON notice_dismissals (user_name);

-- Part 8: User feed state
-- Tracks when each user last viewed the league activity feed.

CREATE TABLE IF NOT EXISTS user_feed_state (
    user_name TEXT PRIMARY KEY,
    last_seen_at TEXT,

    FOREIGN KEY (user_name) REFERENCES users(user_name)
);

-- Part 9: Leader overrides
-- Records leader decisions that resolve tied parlay selections.

CREATE TABLE IF NOT EXISTS leader_overrides (
    season_year INTEGER NOT NULL,
    week TEXT NOT NULL,
    event_id TEXT NOT NULL,
    matchup TEXT,
    picked_team TEXT,
    decided_by TEXT,

    PRIMARY KEY (season_year, week, event_id),

    FOREIGN KEY (event_id) REFERENCES nfl_games(event_id)
);

CREATE INDEX IF NOT EXISTS idx_leader_overrides_season_week
ON leader_overrides (season_year, week);






