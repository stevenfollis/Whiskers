const builder = require('botbuilder');
const refresh = require('../../helpers/token');
const querystring = require('querystring');
const emoji = require('node-emoji');

function login(session) {
  // const telemetryData = telemetryModule.createTelemetry(session);
  // appInsightsClient.trackEvent('botLaunched', telemetryData);

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

module.exports = (bot) => {
  bot.dialog('signin', [
    (session) => {
      session.endDialog();
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
        // const telemetryData = telemetryModule.createTelemetry(session);
        // appInsightsClient.trackEvent('userLoggedIn');

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
        // const telemetryData = telemetryModule.createTelemetry(session);
        // appInsightsClient.trackEvent('invalidCode', telemetryData);

        session.send('Hmm... Looks like that was an invalid code. Please try again.');
        session.replaceDialog('validateCode');
      }
    },
  ]);
};
