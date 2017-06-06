const builder = require('botbuilder');
const request = require('request-promise-native');
const emoji = require('node-emoji');
const moment = require('moment');
moment.locale('en');

module.exports = {
  Label: 'Engagement Information',
  Dialog: [
    (session) => {
      // Create message with card
      const msg = new builder.Message(session)
        .attachmentLayout(builder.AttachmentLayout.list)
        .attachments([
          new builder.HeroCard(session)
            .title('What would you like to do?')
            .buttons([
              builder.CardAction.imBack(session, 'My Engagements', 'My Engagements'),
              builder.CardAction.imBack(session, 'Search', 'Search'),
              builder.CardAction.imBack(session, 'Logout', 'Logout'),
            ]),
        ]);

      // Send message with choice options
      builder.Prompts.choice(
        session,
        msg,
        ['My Engagements', 'Search', 'Logout'],
        {
          maxRetries: 3,
          retryPrompt: 'Not a valid option',
        });
    },
    async (session, results) => {
      switch (results.response.entity) {

        case 'My Engagements': {
          session.replaceDialog('/engagementmy');
          break;
        }

        case 'Search': {
          session.replaceDialog('/search');
          break;
        }

        case 'Logout': {
          // Run logout dialog
          session.beginDialog('/logout');

          // End project dialog
          session.endDialog();

          break;
        }

        default: {
          session.replaceDialog('/');
          break;
        }
      }
    },
  ],
};