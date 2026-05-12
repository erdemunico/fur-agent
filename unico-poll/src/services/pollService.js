const { v4: uuidv4 } = require("uuid");
const db = require("../db/sqlite");
const { addHoursIso, addMinutesIso, nowIso, isPastIso } = require("../utils/time");
const env = require("../config/env");

const MAX_OPTIONS = 10;

function resolveDeadlineFromHours(hours) {
  if (env.fastTestMode) {
    return addMinutesIso(env.fastTestMinutes);
  }
  return addHoursIso(hours);
}

function createPoll({ channelId, creatorId, title, suggestionHours }) {
  const id = uuidv4();
  const createdAt = nowIso();
  const suggestionDeadline = resolveDeadlineFromHours(suggestionHours || env.defaultSuggestionHours);

  db.prepare(`
    INSERT INTO polls (
      id, channel_id, creator_id, title, phase, vote_mode, is_open_vote,
      suggestion_deadline_at, created_at, updated_at
    ) VALUES (
      @id, @channel_id, @creator_id, @title, 'suggestion', 'classic', 0,
      @suggestion_deadline_at, @created_at, @updated_at
    )
  `).run({
    id,
    channel_id: channelId,
    creator_id: creatorId,
    title: title || "Unico Poll",
    suggestion_deadline_at: suggestionDeadline,
    created_at: createdAt,
    updated_at: createdAt,
  });

  return getPollById(id);
}

function getPollById(pollId) {
  return db.prepare("SELECT * FROM polls WHERE id = ?").get(pollId);
}

function getActivePollInChannel(channelId) {
  return db
    .prepare("SELECT * FROM polls WHERE channel_id = ? AND phase != 'closed' ORDER BY created_at DESC LIMIT 1")
    .get(channelId);
}

function addSuggestion({ pollId, userId, parsed }) {
  const poll = getPollById(pollId);
  if (!poll || poll.phase !== "suggestion" || isPastIso(poll.suggestion_deadline_at)) {
    return { ok: false, reason: "Suggestion phase is closed." };
  }

  const exists = db
    .prepare("SELECT id FROM suggestions WHERE poll_id = ? AND lower(display_name) = lower(?)")
    .get(pollId, parsed.displayName);
  if (exists) {
    return { ok: false, reason: "This suggestion already exists." };
  }

  const id = uuidv4();
  db.prepare(`
    INSERT INTO suggestions (
      id, poll_id, submitted_by, raw_text, display_name, pm_keyword, extra, created_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `).run(id, pollId, userId, parsed.rawText, parsed.displayName, parsed.pmKeyword, parsed.extra, nowIso());

  return { ok: true, suggestionId: id };
}

function getUserSuggestionCountSince({ pollId, userId, sinceIso }) {
  return db
    .prepare(
      "SELECT COUNT(1) AS count FROM suggestions WHERE poll_id = ? AND submitted_by = ? AND created_at >= ?"
    )
    .get(pollId, userId, sinceIso).count;
}

function listSuggestions(pollId) {
  return db
    .prepare("SELECT * FROM suggestions WHERE poll_id = ? ORDER BY created_at ASC")
    .all(pollId);
}

function saveShortlist({ pollId, suggestionIds }) {
  const cleanIds = [...new Set(suggestionIds)].slice(0, MAX_OPTIONS);
  const tx = db.transaction(() => {
    db.prepare("DELETE FROM poll_shortlist WHERE poll_id = ?").run(pollId);
    const stmt = db.prepare("INSERT INTO poll_shortlist (poll_id, suggestion_id) VALUES (?, ?)");
    for (const suggestionId of cleanIds) {
      stmt.run(pollId, suggestionId);
    }
  });
  tx();
  return cleanIds;
}

function getShortlistedSuggestions(pollId) {
  const selected = db.prepare(`
    SELECT s.*
    FROM suggestions s
    INNER JOIN poll_shortlist ps ON s.id = ps.suggestion_id
    WHERE ps.poll_id = ?
    ORDER BY s.created_at ASC
  `).all(pollId);

  if (selected.length > 0) {
    return selected;
  }
  return listSuggestions(pollId).slice(0, MAX_OPTIONS);
}

function startVoting({ pollId, voteMode, isOpenVote, votingHours }) {
  const poll = getPollById(pollId);
  if (!poll || !["suggestion", "ready_for_voting"].includes(poll.phase)) {
    throw new Error("Poll is not in suggestion phase.");
  }
  const shortlisted = getShortlistedSuggestions(pollId);
  if (shortlisted.length < 2) {
    throw new Error("At least 2 suggestions are required to start voting.");
  }

  db.prepare(`
    UPDATE polls
    SET phase = 'voting',
        vote_mode = ?,
        is_open_vote = ?,
        voting_deadline_at = ?,
        updated_at = ?
    WHERE id = ?
  `).run(voteMode, isOpenVote ? 1 : 0, resolveDeadlineFromHours(votingHours || env.defaultVotingHours), nowIso(), pollId);

  return getPollById(pollId);
}

function markSuggestionClosed(pollId) {
  db.prepare(`
    UPDATE polls
    SET phase = 'ready_for_voting',
        updated_at = ?
    WHERE id = ? AND phase = 'suggestion'
  `).run(nowIso(), pollId);
  return getPollById(pollId);
}

function castClassicVote({ pollId, userId, suggestionId }) {
  const poll = getPollById(pollId);
  if (!poll || poll.phase !== "voting" || poll.vote_mode !== "classic" || isPastIso(poll.voting_deadline_at)) {
    return { ok: false, reason: "Classic voting is closed." };
  }

  const id = uuidv4();
  db.prepare(`
    INSERT INTO votes_classic (id, poll_id, suggestion_id, user_id, created_at)
    VALUES (?, ?, ?, ?, ?)
    ON CONFLICT (poll_id, user_id) DO UPDATE SET
      suggestion_id = excluded.suggestion_id,
      created_at = excluded.created_at
  `).run(id, pollId, suggestionId, userId, nowIso());

  return { ok: true };
}

function castRatingVote({ pollId, userId, suggestionId, rating }) {
  const poll = getPollById(pollId);
  if (!poll || poll.phase !== "voting" || poll.vote_mode !== "rating" || isPastIso(poll.voting_deadline_at)) {
    return { ok: false, reason: "Rating voting is closed." };
  }

  const score = Number.parseInt(rating, 10);
  if (!Number.isInteger(score) || score < 1 || score > 5) {
    return { ok: false, reason: "Rating must be between 1 and 5." };
  }

  const id = uuidv4();
  db.prepare(`
    INSERT INTO votes_rating (id, poll_id, suggestion_id, user_id, rating, created_at)
    VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT (poll_id, suggestion_id, user_id) DO UPDATE SET
      rating = excluded.rating,
      created_at = excluded.created_at
  `).run(id, pollId, suggestionId, userId, score, nowIso());

  return { ok: true };
}

function closePoll(pollId) {
  db.prepare("UPDATE polls SET phase = 'closed', updated_at = ? WHERE id = ?").run(nowIso(), pollId);
  return getPollById(pollId);
}

function getExpiredSuggestionPolls() {
  return db.prepare(`
    SELECT * FROM polls
    WHERE phase = 'suggestion' AND suggestion_deadline_at IS NOT NULL AND suggestion_deadline_at <= ?
  `).all(nowIso());
}

function getExpiredVotingPolls() {
  return db.prepare(`
    SELECT * FROM polls
    WHERE phase = 'voting' AND voting_deadline_at IS NOT NULL AND voting_deadline_at <= ?
  `).all(nowIso());
}

function buildResults(pollId) {
  const poll = getPollById(pollId);
  const suggestions = getShortlistedSuggestions(pollId);
  if (!poll) {
    return null;
  }

  let rows = [];
  if (poll.vote_mode === "classic") {
    rows = db.prepare(`
      SELECT s.id, s.display_name, s.pm_keyword, s.extra, COUNT(vc.id) AS score
      FROM suggestions s
      LEFT JOIN votes_classic vc ON vc.suggestion_id = s.id AND vc.poll_id = ?
      LEFT JOIN poll_shortlist ps ON ps.suggestion_id = s.id AND ps.poll_id = ?
      WHERE s.poll_id = ? AND (ps.suggestion_id IS NOT NULL OR NOT EXISTS (SELECT 1 FROM poll_shortlist WHERE poll_id = ?))
      GROUP BY s.id
      ORDER BY score DESC, s.display_name ASC
    `).all(pollId, pollId, pollId, pollId);
  } else {
    rows = db.prepare(`
      SELECT s.id, s.display_name, s.pm_keyword, s.extra, COALESCE(ROUND(AVG(vr.rating), 2), 0) AS score
      FROM suggestions s
      LEFT JOIN votes_rating vr ON vr.suggestion_id = s.id AND vr.poll_id = ?
      LEFT JOIN poll_shortlist ps ON ps.suggestion_id = s.id AND ps.poll_id = ?
      WHERE s.poll_id = ? AND (ps.suggestion_id IS NOT NULL OR NOT EXISTS (SELECT 1 FROM poll_shortlist WHERE poll_id = ?))
      GROUP BY s.id
      ORDER BY score DESC, s.display_name ASC
    `).all(pollId, pollId, pollId, pollId);
  }

  const filteredRows = rows.filter((row) => suggestions.some((s) => s.id === row.id));
  return { poll, results: filteredRows };
}

function isCloseResult(results) {
  if (!results || results.length < 2) {
    return false;
  }
  const top = Number(results[0].score);
  const second = Number(results[1].score);
  const diff = Math.abs(top - second);
  return diff <= (top > 10 ? 2 : 0.5);
}

function createRunoff({ sourcePollId, creatorId, channelId, suggestionIds, hours }) {
  const source = getPollById(sourcePollId);
  if (!source) {
    throw new Error("Source poll not found.");
  }
  const poll = createPoll({
    channelId,
    creatorId,
    title: `${source.title} (Run-off)`,
    suggestionHours: 1,
  });

  const sourceSuggestions = listSuggestions(sourcePollId).filter((s) => suggestionIds.includes(s.id));
  for (const suggestion of sourceSuggestions) {
    addSuggestion({
      pollId: poll.id,
      userId: creatorId,
      parsed: {
        rawText: suggestion.raw_text,
        displayName: suggestion.display_name,
        pmKeyword: suggestion.pm_keyword || "",
        extra: suggestion.extra || "",
      },
    });
  }

  const newSuggestions = listSuggestions(poll.id);
  saveShortlist({
    pollId: poll.id,
    suggestionIds: newSuggestions.map((item) => item.id).slice(0, 3),
  });

  return startVoting({
    pollId: poll.id,
    voteMode: source.vote_mode,
    isOpenVote: Boolean(source.is_open_vote),
    votingHours: hours || env.defaultRunoffHours,
  });
}

module.exports = {
  MAX_OPTIONS,
  createPoll,
  getPollById,
  getActivePollInChannel,
  addSuggestion,
  getUserSuggestionCountSince,
  listSuggestions,
  saveShortlist,
  getShortlistedSuggestions,
  startVoting,
  castClassicVote,
  castRatingVote,
  closePoll,
  getExpiredSuggestionPolls,
  getExpiredVotingPolls,
  markSuggestionClosed,
  buildResults,
  isCloseResult,
  createRunoff,
};
