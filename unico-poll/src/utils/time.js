function nowIso() {
  return new Date().toISOString();
}

function addHoursIso(hours) {
  const date = new Date();
  date.setHours(date.getHours() + hours);
  return date.toISOString();
}

function addMinutesIso(minutes) {
  const date = new Date();
  date.setMinutes(date.getMinutes() + minutes);
  return date.toISOString();
}

function isPastIso(isoValue) {
  return new Date(isoValue).getTime() <= Date.now();
}

function formatSlackDate(isoValue) {
  const timestamp = Math.floor(new Date(isoValue).getTime() / 1000);
  return `<!date^${timestamp}^{date_short_pretty} {time}|${isoValue}>`;
}

module.exports = {
  nowIso,
  addHoursIso,
  addMinutesIso,
  isPastIso,
  formatSlackDate,
};
