require('./helpers/telemetry-module');
const builder = require('botbuilder');
const restify = require('restify');
const expressSession = require('express-session');
const emoji = require('node-emoji');
const refresh = require('./helpers/token');
const passport = require('passport-restify');
const stateHelper = require('./helpers/state');

//= ========================================================
// Bot Setup
//= ========================================================

// State Configuration
const tableStorage = stateHelper.getClient();

// Setup Restify Server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log('%s listening to %s', server.name, server.url);
});

// Create Bot
const connector = new builder.ChatConnector({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD,
  gzipData: true,
});
const bot = new builder.UniversalBot(connector, {
  persistConversationData: true,
}).set('storage', tableStorage);

// Setup Bot API routes
server.post('/api/messages', connector.listen());
server.get('/', restify.serveStatic({
  directory: __dirname,
  default: 'index.html',
}));

// Setup Healthcheck
server.get('/ping', (req, res) => {
  res.statusCode = 200;
  res.send('Howdy world');
});

//= ========================================================
// Bots Middleware
//= ========================================================
bot.use(builder.Middleware.sendTyping());

//= ========================================================
// Bots Global Actions
//= ========================================================
bot.endConversationAction('goodbye', 'Goodbye :)', { matches: /^goodbye/i });

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

// Setup dialogs
require('./dialogs/engagement/engagement')(bot);
require('./dialogs/engagement/my/my')(bot);
require('./dialogs/engagement/status/create')(bot);
require('./dialogs/logout/logout')(bot);
require('./dialogs/search/search')(bot);
require('./dialogs/first-run/first-run')(bot);
require('./dialogs/help/help')(bot);
require('./dialogs/login/login')(bot);

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
