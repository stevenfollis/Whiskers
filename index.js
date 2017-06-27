const builder = require('botbuilder');
const restify = require('restify');
const expressSession = require('express-session');
const querystring = require('querystring');
const emoji = require('node-emoji');
const refresh = require('./helpers/token');
const passport = require('passport-restify');

// Configure Application Insights
const telemetryModule = require('./helpers/telemetry-module');
const appInsights = require('applicationinsights');

appInsights.setup(process.env.APPINSIGHTS_INSTRUMENTATIONKEY).start();
const appInsightsClient = appInsights.getClient();

//= ========================================================
// Bot Setup
//= ========================================================

// Setup Restify Server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log('%s listening to %s', server.name, server.url);
});

// Create chat bot
const connector = new builder.ChatConnector({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
  gzipData: true,
});
const bot = new builder.UniversalBot(connector, {
  persistConversationData: true,
});

bot.use(builder.Middleware.sendTyping());

server.post('/api/messages', connector.listen());
server.get('/', restify.serveStatic({
  directory: __dirname,
  default: 'index.html',
}));
server.get('/ping', (req, res) => {
  res.statusCode = 200;
  res.send('Howdy world');
});

//= ========================================================
// Bots Global Actions
//= ========================================================

bot.endConversationAction('goodbye', 'Goodbye :)', { matches: /^goodbye/i });
bot.beginDialogAction('help', '/help', { matches: /^help/i });

//= ========================================================
// Auth Setup
//= ========================================================

server.use(restify.queryParser());
server.use(restify.bodyParser());
server.use(expressSession({
  secret: process.env.SESSION_SECRET,
  resave: true,
  saveUninitialized: false,
}));
server.use(passport.initialize());

// Passport Setup
require('./helpers/passport')(server, bot);

//= ========================================================
// Bots Dialogs
//= ========================================================

function login(session) {
  const telemetryData = telemetryModule.createTelemetry(session);
  appInsightsClient.trackEvent('botLaunched', telemetryData);

  // Generate signin link
  const address = session.message.address;

  // TODO: Encrypt the address string
  const link = `${process.env.AUTHBOT_CALLBACKHOST}/login?address=${querystring.escape(JSON.stringify(address))}`;

  const msg = new builder.Message(session)
    .attachments([
      new builder.HeroCard(session)
        .text(`Let's get started! ${emoji.get('smiley')} Please sign-in below...`)
        .buttons([
          builder.CardAction.openUrl(session, link, 'Sign In'),
        ]),
    ]);
  session.send(msg);

  builder.Prompts.text(session, 'You must first sign into your account.');
}

// Setup dialogs
require('./dialogs/engagement/engagement')(bot);
require('./dialogs/engagement/my/my')(bot);
require('./dialogs/engagement/status/create')(bot);
require('./dialogs/logout/logout')(bot);
require('./dialogs/search/search')(bot);
require('./dialogs/first-run/first-run')(bot);

bot.dialog('signin', [
  (session) => {
    session.endDialog();
  },
]);

bot.dialog('/', [
  (session, args, next) => {
    if (!session.userData.userName) {
      session.beginDialog('signinPrompt');
    } else {
      next();
    }
  },
  (session) => {
    if (session.userData.userName) {  // They're logged in
      refresh(session.userData.refreshToken, (err, body) => {
        if (err || body.error) {
          session.send(`Something happened ${emoji.get('thunder_cloud_and_rain')}`);
          session.endDialog();
        }

        // Stash access token in persisted userData
        session.userData.accessTokenCRM = body.accessToken;

        // Say howdy like a genteel bot
        if (!session.userData.howdy) {
          session.send(`Howdy ${session.userData.userName}! ${emoji.get('smiley')}`);
          session.userData.howdy = true;
        }

        session.beginDialog('/engagement');
      });
    } else {
      session.endConversation(`Goodbye! ${emoji.get('wave')}`);
    }
  },
  (session) => {
    if (!session.userData.userName) {
      session.endConversation(`Goodbye! ${emoji.get('wave')} You have been logged out.`);
    }
  },
]);

bot.dialog('signinPrompt', [
  (session, args) => {
    if (args && args.invalid) {
      // Re-prompt the user to click the link
      builder.Prompts.text(session, 'Please click the signin link.');
    } else if (session.userData.refreshToken) {
      // TODO: Authorization
      // get access token from refresh token
    } else {
      login(session);
    }
  },
  (session, results) => {
    // resuming

    session.userData.loginData = JSON.parse(results.response);
    if (session.userData.loginData && session.userData.loginData.magicCode) {
      session.beginDialog('validateCode');
    } else {
      session.replaceDialog('signinPrompt', { invalid: true });
    }
  },
  (session, results) => {
    if (results.response) {
      // code validated
      session.userData.userName = session.userData.loginData.name;
      session.userData.upn = session.userData.loginData.upn;
      session.endDialogWithResult({ response: true });
    } else {
      session.endDialogWithResult({ response: false });
    }
  },
]);

bot.dialog('validateCode', [
  (session) => {
    builder.Prompts.text(session, "Please enter the code you received or type 'quit' to end. ");
  },
  (session, results) => {
    const code = results.response;
    if (code === 'quit') {
      session.endDialogWithResult({ response: false });
    } else if (code === session.userData.loginData.magicCode) {
      const telemetryData = telemetryModule.createTelemetry(session);
      appInsightsClient.trackEvent('userLoggedIn');

      // Authenticated, save
      session.userData.accessToken = session.userData.loginData.accessToken;
      session.userData.refreshToken = session.userData.loginData.refreshToken;

      // getting access token for CRM
      refresh(session.userData.refreshToken, (err, body) => {
        if (err || body.error) {
          session.send(`Something happened ${emoji.get('thunder_cloud_and_rain')}`);
          session.endDialog();
        }

        session.userData.accessTokenCRM = body.accessToken;

        session.endDialogWithResult({ response: true });
      });
    } else {
      const telemetryData = telemetryModule.createTelemetry(session);
      appInsightsClient.trackEvent('invalidCode', telemetryData);

      session.send('Hmm... Looks like that was an invalid code. Please try again.');
      session.replaceDialog('validateCode');
    }
  },
]);
