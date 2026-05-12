function runMigrations(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS polls (
      id TEXT PRIMARY KEY,
      channel_id TEXT NOT NULL,
      creator_id TEXT NOT NULL,
      title TEXT NOT NULL,
      phase TEXT NOT NULL,
      vote_mode TEXT NOT NULL,
      is_open_vote INTEGER NOT NULL DEFAULT 0,
      suggestion_deadline_at TEXT,
      voting_deadline_at TEXT,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS suggestions (
      id TEXT PRIMARY KEY,
      poll_id TEXT NOT NULL,
      submitted_by TEXT NOT NULL,
      raw_text TEXT NOT NULL,
      display_name TEXT NOT NULL,
      pm_keyword TEXT,
      extra TEXT,
      created_at TEXT NOT NULL,
      FOREIGN KEY (poll_id) REFERENCES polls(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS poll_shortlist (
      poll_id TEXT NOT NULL,
      suggestion_id TEXT NOT NULL,
      PRIMARY KEY (poll_id, suggestion_id),
      FOREIGN KEY (poll_id) REFERENCES polls(id) ON DELETE CASCADE,
      FOREIGN KEY (suggestion_id) REFERENCES suggestions(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS votes_classic (
      id TEXT PRIMARY KEY,
      poll_id TEXT NOT NULL,
      suggestion_id TEXT NOT NULL,
      user_id TEXT NOT NULL,
      created_at TEXT NOT NULL,
      UNIQUE (poll_id, user_id),
      FOREIGN KEY (poll_id) REFERENCES polls(id) ON DELETE CASCADE,
      FOREIGN KEY (suggestion_id) REFERENCES suggestions(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS votes_rating (
      id TEXT PRIMARY KEY,
      poll_id TEXT NOT NULL,
      suggestion_id TEXT NOT NULL,
      user_id TEXT NOT NULL,
      rating INTEGER NOT NULL CHECK (rating BETWEEN 1 AND 5),
      created_at TEXT NOT NULL,
      UNIQUE (poll_id, suggestion_id, user_id),
      FOREIGN KEY (poll_id) REFERENCES polls(id) ON DELETE CASCADE,
      FOREIGN KEY (suggestion_id) REFERENCES suggestions(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS poll_events (
      id TEXT PRIMARY KEY,
      poll_id TEXT NOT NULL,
      actor_id TEXT NOT NULL,
      event_type TEXT NOT NULL,
      payload_json TEXT,
      created_at TEXT NOT NULL,
      FOREIGN KEY (poll_id) REFERENCES polls(id) ON DELETE CASCADE
    );
  `);
}

module.exports = {
  runMigrations,
};
