const cron = require("node-cron");
const pollService = require("./pollService");
const { notifySuggestionPhaseEnded } = require("../slack/commands");
const { sendCreatorResults } = require("../slack/actions");
const logger = require("../utils/logger");

function registerScheduler(app) {
  cron.schedule("*/1 * * * *", async () => {
    const expiredSuggestion = pollService.getExpiredSuggestionPolls();
    for (const poll of expiredSuggestion) {
      try {
        await notifySuggestionPhaseEnded({ app, poll });
        pollService.markSuggestionClosed(poll.id);
      } catch (err) {
        logger.error("Suggestion phase close failed", { pollId: poll.id, error: err.message });
      }
    }

    const expiredVoting = pollService.getExpiredVotingPolls();
    for (const poll of expiredVoting) {
      try {
        pollService.closePoll(poll.id);
        await sendCreatorResults(app, poll.id);
      } catch (err) {
        logger.error("Voting phase close failed", { pollId: poll.id, error: err.message });
      }
    }
  });
}

module.exports = {
  registerScheduler,
};
