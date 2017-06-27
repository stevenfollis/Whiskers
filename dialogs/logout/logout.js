const emoji = require('node-emoji');

module.exports = (bot) => {
  bot.dialog('/logout', [
    (session) => {
      // Clear all userData
      session.userData = {};

      // Give a heartfelt goodbye
      session.endConversation(`Catch you later, alligator. ${emoji.get('crocodile')}`);
      session.beginDialog('/');
    },
  ]).triggerAction({
    matches: /^(logout|reset)$/,
  });
};
