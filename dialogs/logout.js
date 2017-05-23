'use strict';

module.exports = {
    Label: 'Logout',
    Dialog: [
        function (session) {
            session.userData.userName = null;
            session.userData.accessToken = null;
            session.userData.accessTokenCRM = null;
            session.userData.refreshToken = null;
            session.userData.crmId = null;
            
            session.endDialog();
        }
    ]
};