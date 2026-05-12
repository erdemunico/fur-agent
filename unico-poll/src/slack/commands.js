const { parseHoursArg, parseSuggestionInput } = require("../utils/parser");
const env = require("../config/env");
const pollService = require("../services/pollService");
const logger = require("../utils/logger");
const {
  suggestionAnnouncementBlocks,
  creatorSuggestionControlBlocks,
} = require("./blocks");

function parseCommandText(text) {
  const raw = String(text || "").trim();
  if (!raw) {
    return { title: "Unico Poll", hours: env.defaultSuggestionHours };
  }

  const [titleRaw, optionsRaw = ""] = raw.split("|");
  return {
    title: titleRaw.trim() || "Unico Poll",
    hours: parseHoursArg(optionsRaw, env.defaultSuggestionHours),
  };
}

function registerCommands(app) {
  app.command("/unico-poll", async ({ ack, body, client, command }) => {
    await ack();
    const channelId = body.channel_id;
    const creatorId = body.user_id;
    try {
      if (env.allowedCreatorIds.length > 0 && !env.allowedCreatorIds.includes(creatorId)) {
        await client.chat.postEphemeral({
          channel: channelId,
          user: creatorId,
          text: "Bu komutu kullanma yetkin yok.",
        });
        logger.warn("Unauthorized poll creation attempt", { userId: creatorId, channelId });
        return;
      }

      const activePoll = pollService.getActivePollInChannel(channelId);
      if (activePoll) {
        await client.chat.postEphemeral({
          channel: channelId,
          user: creatorId,
          text: "Bu kanalda zaten aktif bir anket var.",
        });
        return;
      }

      const { title, hours } = parseCommandText(command.text);
      const poll = pollService.createPoll({ channelId, creatorId, title, suggestionHours: hours });

      await client.chat.postMessage({
        channel: channelId,
        text: `@channel Unico Poll basladi: ${title}`,
        blocks: suggestionAnnouncementBlocks(poll),
      });

      await client.chat.postEphemeral({
        channel: channelId,
        user: creatorId,
        text: "Anket olusturuldu. Oneri toplama asamasi basladi.",
      });
      logger.info("Poll created", { pollId: poll.id, channelId, userId: creatorId });
    } catch (error) {
      logger.error("Failed to handle /unico-poll", {
        userId: creatorId,
        channelId,
        error: error.message,
      });
      await client.chat.postEphemeral({
        channel: channelId,
        user: creatorId,
        text: "Anket olusturulurken bir hata olustu.",
      });
    }
  });

  app.message(async ({ message, client }) => {
    if (!message || message.subtype || !message.channel || !message.user || !message.text) {
      return;
    }
    try {
      const poll = pollService.getActivePollInChannel(message.channel);
      if (!poll || poll.phase !== "suggestion") {
        return;
      }

      const parsed = parseSuggestionInput(message.text);
      if (!parsed) {
        return;
      }

      const windowMinutes = Math.max(1, env.suggestionRateLimitWindowMinutes);
      const limitCount = Math.max(1, env.suggestionRateLimitCount);
      const since = new Date(Date.now() - windowMinutes * 60 * 1000).toISOString();
      const recentCount = pollService.getUserSuggestionCountSince({
        pollId: poll.id,
        userId: message.user,
        sinceIso: since,
      });
      if (recentCount >= limitCount) {
        logger.warn("Suggestion rate limit hit", {
          pollId: poll.id,
          userId: message.user,
          recentCount,
          limitCount,
          windowMinutes,
        });
        await client.chat.postEphemeral({
          channel: message.channel,
          user: message.user,
          text: `Cok hizli oneride bulundun. Lutfen ${windowMinutes} dakika icinde en fazla ${limitCount} oneriyi gecme.`,
        });
        return;
      }

      const result = pollService.addSuggestion({
        pollId: poll.id,
        userId: message.user,
        parsed,
      });

      if (!result.ok) {
        await client.chat.postEphemeral({
          channel: message.channel,
          user: message.user,
          text: result.reason,
        });
        return;
      }

      await client.chat.postEphemeral({
        channel: message.channel,
        user: message.user,
        text: `Onerin alindi: *${parsed.displayName}*`,
      });
    } catch (error) {
      logger.error("Suggestion processing failed", {
        userId: message.user,
        channelId: message.channel,
        error: error.message,
      });
    }
  });
}

async function notifySuggestionPhaseEnded({ app, poll }) {
  try {
    const suggestions = pollService.listSuggestions(poll.id);
    await app.client.chat.postEphemeral({
      channel: poll.channel_id,
      user: poll.creator_id,
      text: "Oneri toplama suresi doldu.",
      blocks: creatorSuggestionControlBlocks(poll, suggestions, pollService.MAX_OPTIONS),
    });
  } catch (error) {
    logger.error("Failed to notify suggestion end", { pollId: poll.id, error: error.message });
  }
}

module.exports = {
  registerCommands,
  notifySuggestionPhaseEnded,
};
