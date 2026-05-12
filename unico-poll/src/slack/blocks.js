const { formatSlackDate } = require("../utils/time");

function suggestionAnnouncementBlocks(poll) {
  return [
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: `@channel *Unico Poll* basladi: *${poll.title}*\nOneri suresi: ${formatSlackDate(
          poll.suggestion_deadline_at
        )}\nFormat: \`Oneri Ismi : PM Keyword ; Ekstra\``,
      },
    },
    {
      type: "context",
      elements: [{ type: "mrkdwn", text: "Ankette sadece ':' oncesi kisim gosterilir." }],
    },
  ];
}

function creatorSuggestionControlBlocks(poll, suggestions, maxOptions) {
  const items = suggestions.map((s, idx) => `*${idx + 1}.* ${s.display_name}`).join("\n");
  return [
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: `*${poll.title}* onerileri tamamlandi.\nToplam: *${suggestions.length}*`,
      },
    },
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: items || "_Henuz oneriler yok._",
      },
    },
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: `10'dan fazla oneride ankete girecek en fazla *${maxOptions}* secenegi modal ile secebilirsin.`,
      },
      accessory: {
        type: "button",
        text: { type: "plain_text", text: "Shortlist ve Oylamayi Baslat" },
        action_id: "open_start_voting_modal",
        value: poll.id,
      },
    },
  ];
}

function buildStartVotingModal({ poll, suggestions }) {
  const options = suggestions.map((s) => ({
    text: { type: "plain_text", text: s.display_name.slice(0, 75) },
    value: s.id,
  }));

  return {
    type: "modal",
    callback_id: "start_voting_submit",
    private_metadata: JSON.stringify({ pollId: poll.id }),
    title: { type: "plain_text", text: "Unico Poll" },
    submit: { type: "plain_text", text: "Oylamayi Baslat" },
    close: { type: "plain_text", text: "Iptal" },
    blocks: [
      {
        type: "input",
        block_id: "vote_mode",
        label: { type: "plain_text", text: "Oylama Turu" },
        element: {
          type: "static_select",
          action_id: "vote_mode_select",
          options: [
            { text: { type: "plain_text", text: "Klasik (tek oy)" }, value: "classic" },
            { text: { type: "plain_text", text: "Puanlama (1-5)" }, value: "rating" },
          ],
          initial_option: { text: { type: "plain_text", text: "Klasik (tek oy)" }, value: "classic" },
        },
      },
      {
        type: "input",
        block_id: "vote_privacy",
        label: { type: "plain_text", text: "Oy Gorunurlugu" },
        element: {
          type: "static_select",
          action_id: "vote_privacy_select",
          options: [
            { text: { type: "plain_text", text: "Kapali (anonim - varsayilan)" }, value: "closed" },
            { text: { type: "plain_text", text: "Acik (kullanici secimi gorunebilir)" }, value: "open" },
          ],
          initial_option: { text: { type: "plain_text", text: "Kapali (anonim - varsayilan)" }, value: "closed" },
        },
      },
      {
        type: "input",
        block_id: "vote_duration",
        label: { type: "plain_text", text: "Oylama Suresi (saat)" },
        element: {
          type: "plain_text_input",
          action_id: "vote_duration_input",
          initial_value: "48",
        },
      },
      {
        type: "input",
        optional: options.length <= 10,
        block_id: "shortlist",
        label: { type: "plain_text", text: "Ankete girecek secenekler (max 10)" },
        element: {
          type: "multi_static_select",
          action_id: "shortlist_select",
          placeholder: { type: "plain_text", text: "Secenekleri secin" },
          options,
          max_selected_items: 10,
        },
      },
    ],
  };
}

function votingBlocks({ poll, suggestions }) {
  const blocks = [
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: `*${poll.title}* oylamasi acik.\nBitis: ${formatSlackDate(poll.voting_deadline_at)}\nAra sonuclar gizlidir.`,
      },
    },
  ];

  if (poll.vote_mode === "classic") {
    blocks.push({
      type: "actions",
      block_id: `classic_vote_${poll.id}`,
      elements: suggestions.slice(0, 5).map((s) => ({
        type: "button",
        text: { type: "plain_text", text: s.display_name.slice(0, 75) },
        action_id: "classic_vote_click",
        value: JSON.stringify({ pollId: poll.id, suggestionId: s.id }),
      })),
    });
    if (suggestions.length > 5) {
      blocks.push({
        type: "section",
        text: { type: "mrkdwn", text: "_Tum secenekler icin acilan modal ekranini kullanin._" },
        accessory: {
          type: "button",
          text: { type: "plain_text", text: "Tum Secenekleri Ac" },
          action_id: "open_classic_vote_modal",
          value: poll.id,
        },
      });
    }
  } else {
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: "Her secenegi 1-5 arasi puanlayin.",
      },
      accessory: {
        type: "button",
        text: { type: "plain_text", text: "Puanlama Ekranini Ac" },
        action_id: "open_rating_modal",
        value: poll.id,
      },
    });
  }

  return blocks;
}

function buildClassicVoteModal({ poll, suggestions }) {
  return {
    type: "modal",
    callback_id: "classic_vote_submit",
    private_metadata: JSON.stringify({ pollId: poll.id }),
    title: { type: "plain_text", text: "Tek Oy Kullan" },
    submit: { type: "plain_text", text: "Kaydet" },
    close: { type: "plain_text", text: "Iptal" },
    blocks: [
      {
        type: "input",
        block_id: "classic_vote_choice",
        label: { type: "plain_text", text: "Secimin" },
        element: {
          type: "static_select",
          action_id: "classic_vote_select",
          options: suggestions.map((s) => ({
            text: { type: "plain_text", text: s.display_name.slice(0, 75) },
            value: s.id,
          })),
        },
      },
    ],
  };
}

function buildRatingModal({ poll, suggestions }) {
  const blocks = suggestions.map((s) => ({
    type: "input",
    block_id: `rating_${s.id}`,
    label: { type: "plain_text", text: s.display_name.slice(0, 75) },
    element: {
      type: "static_select",
      action_id: "rating_value",
      options: [1, 2, 3, 4, 5].map((score) => ({
        text: { type: "plain_text", text: `${score} / 5` },
        value: String(score),
      })),
    },
  }));

  return {
    type: "modal",
    callback_id: "rating_vote_submit",
    private_metadata: JSON.stringify({ pollId: poll.id }),
    title: { type: "plain_text", text: "Puanlama" },
    submit: { type: "plain_text", text: "Kaydet" },
    close: { type: "plain_text", text: "Iptal" },
    blocks,
  };
}

function creatorResultsBlocks({ poll, results, close }) {
  const rows = results
    .map((r, i) => `${i + 1}. *${r.display_name}* - ${r.score}`)
    .join("\n");

  const actionElements = [
    {
      type: "button",
      text: { type: "plain_text", text: "Sonuclari Kanala Yayinla" },
      style: "primary",
      action_id: "publish_results",
      value: poll.id,
    },
  ];

  if (close) {
    actionElements.push({
      type: "button",
      text: { type: "plain_text", text: "Yakin Oylari Sec ve Run-off Baslat" },
      action_id: "start_runoff",
      value: poll.id,
    });
  }

  return [
    {
      type: "section",
      text: { type: "mrkdwn", text: `*${poll.title}* sonucu hazir (creator-only).` },
    },
    { type: "section", text: { type: "mrkdwn", text: rows || "_Oy yok._" } },
    { type: "actions", elements: actionElements },
  ];
}

function channelResultsBlocks({ poll, results }) {
  const winner = results[0];
  const lines = results.map((r, idx) => {
    const suffix = idx === 0 && r.pm_keyword ? ` _(PM: ${r.pm_keyword}${r.extra ? ` | ${r.extra}` : ""})_` : "";
    return `*${idx + 1}.* ${r.display_name} - ${r.score}${suffix}`;
  });
  return [
    {
      type: "section",
      text: { type: "mrkdwn", text: `*${poll.title}* sonuclari yayinlandi.` },
    },
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: lines.join("\n"),
      },
    },
    {
      type: "context",
      elements: [{ type: "mrkdwn", text: `Kazanan: *${winner ? winner.display_name : "Yok"}*` }],
    },
  ];
}

module.exports = {
  suggestionAnnouncementBlocks,
  creatorSuggestionControlBlocks,
  buildStartVotingModal,
  votingBlocks,
  buildClassicVoteModal,
  buildRatingModal,
  creatorResultsBlocks,
  channelResultsBlocks,
};
