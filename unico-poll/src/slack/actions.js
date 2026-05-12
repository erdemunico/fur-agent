const pollService = require("../services/pollService");
const env = require("../config/env");
const logger = require("../utils/logger");
const {
  buildStartVotingModal,
  votingBlocks,
  buildClassicVoteModal,
  buildRatingModal,
  creatorResultsBlocks,
  channelResultsBlocks,
} = require("./blocks");

function parseMetadata(raw) {
  try {
    return JSON.parse(raw || "{}");
  } catch (err) {
    return {};
  }
}

async function sendCreatorResults(app, pollId) {
  try {
    const data = pollService.buildResults(pollId);
    if (!data) {
      return;
    }
    const close = pollService.isCloseResult(data.results);

    await app.client.chat.postEphemeral({
      channel: data.poll.channel_id,
      user: data.poll.creator_id,
      text: "Oylama bitti. Sonuclar sadece sana gonderildi.",
      blocks: creatorResultsBlocks({ poll: data.poll, results: data.results, close }),
    });
  } catch (error) {
    logger.error("Failed to send creator results", { pollId, error: error.message });
  }
}

function registerActions(app) {
  app.action("open_start_voting_modal", async ({ ack, body, client }) => {
    await ack();
    const pollId = body.actions?.[0]?.value;
    const poll = pollService.getPollById(pollId);
    if (!poll || body.user.id !== poll.creator_id) {
      return;
    }
    const suggestions = pollService.listSuggestions(pollId);

    await client.views.open({
      trigger_id: body.trigger_id,
      view: buildStartVotingModal({ poll, suggestions }),
    });
  });

  app.view("start_voting_submit", async ({ ack, body, view, client }) => {
    await ack();
    const meta = parseMetadata(view.private_metadata);
    const poll = pollService.getPollById(meta.pollId);
    if (!poll || body.user.id !== poll.creator_id) {
      return;
    }

    const state = view.state.values;
    const mode = state.vote_mode.vote_mode_select.selected_option.value;
    const privacy = state.vote_privacy.vote_privacy_select.selected_option.value;
    const hours = Number.parseInt(state.vote_duration.vote_duration_input.value, 10) || env.defaultVotingHours;
    const shortlistValues = state.shortlist?.shortlist_select?.selected_options?.map((o) => o.value) || [];

    const allSuggestions = pollService.listSuggestions(poll.id);
    const selected = shortlistValues.length > 0 ? shortlistValues : allSuggestions.map((s) => s.id).slice(0, 10);
    pollService.saveShortlist({ pollId: poll.id, suggestionIds: selected });

    try {
      const updatedPoll = pollService.startVoting({
        pollId: poll.id,
        voteMode: mode,
        isOpenVote: privacy === "open",
        votingHours: hours,
      });
      const shortlist = pollService.getShortlistedSuggestions(poll.id);

      await client.chat.postMessage({
        channel: poll.channel_id,
        text: `${poll.title} oylamasi basladi.`,
        blocks: votingBlocks({ poll: updatedPoll, suggestions: shortlist }),
      });
    } catch (error) {
      logger.error("Failed to start voting", { pollId: poll.id, userId: body.user.id, error: error.message });
      await client.chat.postEphemeral({
        channel: poll.channel_id,
        user: body.user.id,
        text: "Oylama baslatilirken hata olustu.",
      });
    }
  });

  app.action("classic_vote_click", async ({ ack, body, client }) => {
    await ack();
    const value = parseMetadata(body.actions?.[0]?.value);
    const vote = pollService.castClassicVote({
      pollId: value.pollId,
      userId: body.user.id,
      suggestionId: value.suggestionId,
    });
    await client.chat.postEphemeral({
      channel: body.channel.id,
      user: body.user.id,
      text: vote.ok ? "Oyun kaydedildi." : vote.reason,
    });
  });

  app.action("open_classic_vote_modal", async ({ ack, body, client }) => {
    await ack();
    const pollId = body.actions?.[0]?.value;
    const poll = pollService.getPollById(pollId);
    if (!poll) {
      return;
    }
    const suggestions = pollService.getShortlistedSuggestions(pollId);
    await client.views.open({
      trigger_id: body.trigger_id,
      view: buildClassicVoteModal({ poll, suggestions }),
    });
  });

  app.view("classic_vote_submit", async ({ ack, body, view, client }) => {
    await ack();
    const meta = parseMetadata(view.private_metadata);
    const suggestionId = view.state.values.classic_vote_choice.classic_vote_select.selected_option.value;
    const vote = pollService.castClassicVote({
      pollId: meta.pollId,
      userId: body.user.id,
      suggestionId,
    });
    await client.chat.postEphemeral({
      channel: pollService.getPollById(meta.pollId).channel_id,
      user: body.user.id,
      text: vote.ok ? "Oyun kaydedildi." : vote.reason,
    });
  });

  app.action("open_rating_modal", async ({ ack, body, client }) => {
    await ack();
    const pollId = body.actions?.[0]?.value;
    const poll = pollService.getPollById(pollId);
    if (!poll) {
      return;
    }
    const suggestions = pollService.getShortlistedSuggestions(pollId);
    await client.views.open({
      trigger_id: body.trigger_id,
      view: buildRatingModal({ poll, suggestions }),
    });
  });

  app.view("rating_vote_submit", async ({ ack, body, view, client }) => {
    await ack();
    const meta = parseMetadata(view.private_metadata);
    const poll = pollService.getPollById(meta.pollId);
    const stateValues = view.state.values;

    for (const [blockId, value] of Object.entries(stateValues)) {
      const suggestionId = blockId.replace("rating_", "");
      const rating = value.rating_value.selected_option.value;
      pollService.castRatingVote({
        pollId: poll.id,
        userId: body.user.id,
        suggestionId,
        rating,
      });
    }

    await client.chat.postEphemeral({
      channel: poll.channel_id,
      user: body.user.id,
      text: "Puanlarin kaydedildi.",
    });
  });

  app.action("publish_results", async ({ ack, body, client }) => {
    await ack();
    const pollId = body.actions?.[0]?.value;
    const data = pollService.buildResults(pollId);
    if (!data || body.user.id !== data.poll.creator_id) {
      return;
    }
    try {
      await client.chat.postMessage({
        channel: data.poll.channel_id,
        text: `${data.poll.title} sonuclari`,
        blocks: channelResultsBlocks(data),
      });
      logger.info("Results published", { pollId, userId: body.user.id });
    } catch (error) {
      logger.error("Failed to publish results", { pollId, error: error.message });
    }
  });

  app.action("start_runoff", async ({ ack, body, client }) => {
    await ack();
    const pollId = body.actions?.[0]?.value;
    const data = pollService.buildResults(pollId);
    if (!data || body.user.id !== data.poll.creator_id) {
      return;
    }
    try {
      const topIds = data.results.slice(0, 3).map((row) => row.id);
      const runoff = pollService.createRunoff({
        sourcePollId: pollId,
        creatorId: data.poll.creator_id,
        channelId: data.poll.channel_id,
        suggestionIds: topIds,
        hours: env.defaultRunoffHours,
      });
      const shortlist = pollService.getShortlistedSuggestions(runoff.id);

      await client.chat.postMessage({
        channel: runoff.channel_id,
        text: `${runoff.title} run-off oylamasi basladi.`,
        blocks: votingBlocks({ poll: runoff, suggestions: shortlist }),
      });
      logger.info("Runoff started", { sourcePollId: pollId, runoffPollId: runoff.id, userId: body.user.id });
    } catch (error) {
      logger.error("Failed to start runoff", { pollId, error: error.message });
      await client.chat.postEphemeral({
        channel: data.poll.channel_id,
        user: body.user.id,
        text: "Run-off baslatilirken hata olustu.",
      });
    }
  });
}

module.exports = {
  registerActions,
  sendCreatorResults,
};
