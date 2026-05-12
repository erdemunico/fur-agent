const MAX_DISPLAY_NAME_LENGTH = 80;
const MAX_PM_KEYWORD_LENGTH = 64;
const MAX_EXTRA_LENGTH = 256;

function sanitizeValue(value, maxLen) {
  return String(value || "")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLen);
}

function parseSuggestionInput(rawText) {
  const raw = String(rawText || "").trim();
  if (!raw) {
    return null;
  }

  const [beforeColon, afterColon = ""] = raw.split(":");
  const displayName = sanitizeValue(beforeColon || "", MAX_DISPLAY_NAME_LENGTH);
  if (!displayName) {
    return null;
  }

  const [pmKeywordRaw = "", extraRaw = ""] = afterColon.split(";");
  return {
    rawText: raw,
    displayName,
    pmKeyword: sanitizeValue(pmKeywordRaw, MAX_PM_KEYWORD_LENGTH),
    extra: sanitizeValue(extraRaw, MAX_EXTRA_LENGTH),
  };
}

function parseHoursArg(text, fallbackHours) {
  const match = String(text || "").match(/(?:\b|^)(\d+)\s*h(?:ours?)?(?:\b|$)/i);
  if (!match) {
    return fallbackHours;
  }
  const hours = Number.parseInt(match[1], 10);
  if (!Number.isFinite(hours) || hours < 1 || hours > 336) {
    return fallbackHours;
  }
  return hours;
}

module.exports = {
  MAX_DISPLAY_NAME_LENGTH,
  parseSuggestionInput,
  parseHoursArg,
};
