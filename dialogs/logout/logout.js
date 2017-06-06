const emoji = require('node-emoji');

module.exports = {
  Label: 'Logout',
  Dialog: [
    (session) => {
      // Clear all userData
      session.userData.userName = null;
      session.userData.accessToken = null;
      session.userData.accessTokenCRM = null;
      session.userData.refreshToken = null;
      session.userData.loginData = null;
      session.userData.upn = null;

      // Give a heartfelt goodbye
      session.endConversation(`Catch you later alligator. ${emoji.get('crocodile')}`);
      session.beginDialog('/');
    },
  ],
};
