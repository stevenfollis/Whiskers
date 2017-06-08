const builder = require('botbuilder');
const request = require('request-promise-native');
const emoji = require('node-emoji');
const moment = require('moment');

module.exports = (bot) => {
  bot.dialog('/engagementstatuscreate', [
    (session, args) => {
      // Create message with card
      const msg = new builder.Message(session)
        .attachmentLayout(builder.AttachmentLayout.list)
        .attachments([
          new builder.HeroCard(session)
            .title('The engagement is in which phase?')
            .buttons([
              builder.CardAction.imBack(session, 'New', 'New'),
              builder.CardAction.imBack(session, 'Active', 'Active'),
              builder.CardAction.imBack(session, 'Complete', 'Complete'),
              builder.CardAction.imBack(session, 'Engaged', 'Engaged'),
              builder.CardAction.imBack(session, 'Disengaged', 'Disengaged'),
              builder.CardAction.imBack(session, 'Transitioned', 'Transitioned'),
              builder.CardAction.imBack(session, 'Queued', 'Queued'),
            ]),
        ]);

      // Send message with choice options
      builder.Prompts.choice(
        session,
        msg,
        ['New', 'Active', 'Complete', 'Engaged', 'Disengaged', 'Transitioned', 'Queued'],
        {
          maxRetries: 3,
          retryPrompt: 'Not a valid option',
        });
    },
    (session, results, next) => {
      // Define variable to hold phase
      let phase;

      // Store engagement status
      switch (results.response.entity) {
        case 'New':
          phase = 100000003;
          break;
        case 'Active':
          phase = 100000002;
          break;
        case 'Complete':
          phase = 100000004;
          break;
        case 'Engaged':
          phase = 100000005;
          break;
        case 'Disengaged':
          phase = 100000007;
          break;
        case 'Transitioned':
          phase = 100000009;
          break;
        case 'Queued':
          phase = 100000003;
          break;
        default:
          phase = 100000002;
          break;
      }

      // Store status in dialog data
      session.dialogData.phase = phase;

      next();
    },
    (session) => {
      // Prompt user for status
      builder.Prompts.text(session, 'Thanks. Now, what is the current status of this project?');
    },
    (session, results, next) => {
      // Store status
      session.dialogData.status = results.response;
      next();
    },
    (session) => {
      // Prompt user for next steps
      builder.Prompts.text(session, 'Got it. Finally, what are the next steps for this project?');
    },
    (session, results, next) => {
      // Store next steps
      session.dialogData.nextSteps = results.response;
      next();
    },
    (session) => {
      // Create message with card
      const msg = new builder.Message(session)
        .attachmentLayout(builder.AttachmentLayout.list)
        .attachments([
          new builder.HeroCard(session)
            .title('Would you like to submit this update?')
            .buttons([
              builder.CardAction.imBack(session, 'Submit', 'Submit'),
              builder.CardAction.imBack(session, 'Cancel', 'Cancel'),
            ]),
        ]);

      // Send message with choice options
      builder.Prompts.choice(
        session,
        msg,
        ['Submit', 'Cancel'],
        {
          maxRetries: 3,
          retryPrompt: 'Not a valid option',
        });
    },
    (session, results) => {
      // Upload message, else end dialog
      switch (results.response.entity) {
        case 'Submit':
          session.send('Thanks for providing this status update. (DEMO for now)');
          session.replaceDialog('/');
          break;
        case 'Cancel':
          session.send('Cancelled your status update');
          session.replaceDialog('/');
          break;
        default:
          break;
      }
    },
  ]);
};
