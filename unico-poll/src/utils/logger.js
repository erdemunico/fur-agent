function formatMeta(meta) {
  if (!meta || typeof meta !== "object") {
    return "";
  }
  const compact = Object.entries(meta)
    .filter(([, value]) => value !== undefined && value !== null)
    .map(([key, value]) => `${key}=${String(value)}`)
    .join(" ");
  return compact ? ` ${compact}` : "";
}

function log(level, message, meta) {
  const time = new Date().toISOString();
  // eslint-disable-next-line no-console
  console.log(`[${time}] [${level}] ${message}${formatMeta(meta)}`);
}

module.exports = {
  info: (message, meta) => log("INFO", message, meta),
  warn: (message, meta) => log("WARN", message, meta),
  error: (message, meta) => log("ERROR", message, meta),
};
