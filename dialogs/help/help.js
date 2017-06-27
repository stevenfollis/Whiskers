const builder = require('botbuilder');
const emoji = require('node-emoji');

module.exports = (bot) => {
  // Add first run dialog
  bot.dialog('help', (session) => {
    // Create Message
    const msg = new builder.Message(session);
    msg
      .attachmentLayout(builder.AttachmentLayout.list)
      .attachments([
        new builder.HeroCard(session)
          .title(`Help is on the way! ${emoji.get('question')}`)
          .text('Say "menu" at any time to see a list of activities that I can help with'),
        new builder.HeroCard(session)
          .text('Say "reset" at any time for a new beginning'),
        new builder.HeroCard(session)
          .text('Say "cancel" or "back" to back up a level in a dialog'),
      ]);

    // Send help message
    session.send(msg);

    // Redirect to first run for an overview of Whiskers
    session.replaceDialog('firstRun');
  }).triggerAction({
    matches: /help/i,
  });
};
