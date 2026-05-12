const { App } = require("@slack/bolt");
const env = require("./config/env");
require("./db/sqlite");
const { registerCommands } = require("./slack/commands");
const { registerActions } = require("./slack/actions");
const { registerScheduler } = require("./services/scheduler");

const app = new App({
  token: env.slackBotToken,
  signingSecret: env.slackSigningSecret,
  socketMode: true,
  appToken: env.slackAppToken,
});

registerCommands(app);
registerActions(app);
registerScheduler(app);

(async () => {
  await app.start(env.port);
  // eslint-disable-next-line no-console
  console.log(`Unico Poll is running on port ${env.port}`);
})();
