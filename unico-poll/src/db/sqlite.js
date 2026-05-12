const fs = require("fs");
const path = require("path");
const Database = require("better-sqlite3");
const env = require("../config/env");
const { runMigrations } = require("./migrations");

const dbDir = path.dirname(env.databasePath);
if (!fs.existsSync(dbDir)) {
  fs.mkdirSync(dbDir, { recursive: true });
}

const db = new Database(env.databasePath);
db.pragma("journal_mode = WAL");
db.pragma("foreign_keys = ON");

runMigrations(db);

module.exports = db;
