module.exports = {
  Label: 'Logout',
  Dialog: [
    (session) => {
      session.userData.userName = null;
      session.userData.accessToken = null;
      session.userData.accessTokenCRM = null;
      session.userData.refreshToken = null;

      session.endDialog();
    },
  ],
};
