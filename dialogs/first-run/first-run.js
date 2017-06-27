const builder = require('botbuilder');

module.exports = (bot) => {
  // Add first run dialog
  bot.dialog('firstRun', (session) => {
    // Update userProfile
    session.userData.firstRun = true;

    // Create Message
    const msg = new builder.Message(session);
    msg.attachmentLayout(builder.AttachmentLayout.carousel);

    // Create Carousel
    msg.attachments([
      new builder.AnimationCard(session)
        .title('Hi there! I\'m Whiskers!')
        .subtitle('I\'m a cat here to help AzureCAT')
        .text('I know how to do several tricks! After you sign in, say "menu" at any time to see all options. I also respond to "back", "help", and "reset"')
        .media([
          { url: 'https://media.giphy.com/media/VOPK1BqsMEJRS/giphy.gif' },
        ]),
      new builder.HeroCard(session)
        .title('Get Your Engagements')
        .text('I can grab your active engagmenets from CRM, even popping them into a browser if desired'),
      new builder.HeroCard(session)
        .title('Provide Status Update')
        .text('Want to provide an update? Answer a few questions and I\'ll handle everything'),
      new builder.HeroCard(session)
        .title('Search All Engagements')
        .text('Looking for a particular engagmenet? Let me query CRM for you'),
    ]);

    session.send(msg);
    session.replaceDialog('/');
  }).triggerAction({
    onFindAction: (context, callback) => {
      // Only trigger if we've never seen user before
      if (!context.userData.firstRun) {
        // Return a score of 1.1 to ensure the first run dialog wins
        callback(null, 1.1);
      } else {
        callback(null, 0.0);
      }
    },
  });
};
