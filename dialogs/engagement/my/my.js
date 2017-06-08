const builder = require('botbuilder');
const request = require('request-promise-native');
const emoji = require('node-emoji');
const moment = require('moment');

async function getCrmId(userData) {
  return new Promise((resolve, reject) => {
    const options = {
      url: `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.1/systemusers?$filter=contains(domainname,%27${userData.upn}%27)&$select=systemuserid`,
      headers: { Authorization: userData.accessTokenCRM },
      json: true,
    };

    // Query the API
    request
      .get(options)
      .then((result) => {
        resolve(result.value[0].systemuserid);
      })
      .catch((error) => {
        reject(error);
      });
  });
}

async function getEngagements(userData) {
  // Create OData URL for the user's active engagements
  const url = `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.1/ee_projects%20?$select=ee_projectname,ee_engagementcurrentphase,modifiedon%20&$filter=(%20_ownerid_value%20eq%20${userData.crmId}%20or%20_ee_coowner1_value%20eq%20${userData.crmId}%20or%20_ee_coowner2_value%20eq%20${userData.crmId}%20or%20_ee_coowner3_value%20eq%20${userData.crmId}%20or%20_ee_coowner4_value%20eq%20${userData.crmId}%20or%20_ee_coowner5_value%20eq%20${userData.crmId}%20)%20and%20ee_engagementcurrentphase%20eq%20100000002`;

  const options = {
    url,
    headers: { Authorization: userData.accessTokenCRM },
    json: true,
  };

  return request.get(options);
}

module.exports = (bot) => {
  bot.dialog('/engagementmy', [
    async (session) => {
      // Get CRM ID
      session.userData.crmId = await getCrmId(session.userData);

      // Get Engagements
      const engagements = await getEngagements(session.userData);
      session.dialogData.engagements = engagements;

      // Create Carousel
      const msg = new builder.Message(session);
      msg.attachmentLayout(builder.AttachmentLayout.carousel);

      // Create Cards
      const attachments = engagements.value.map((engagement) => {
        return new builder.HeroCard(session)
          .title(engagement.ee_projectname)
          .subtitle(`Last updated ${moment(engagement.modifiedon).fromNow()}`)
          .text('Select this engagement with the button')
          .buttons([
            builder.CardAction.openUrl(session, `${process.env.MICROSOFT_RESOURCE_CRM}/main.aspx?etc=10096&extraqs=formid=33679b2b-bfd5-4e84-a63d-13aa63146ebb&pagetype=entityrecord&id={${engagement.ee_projectid}}`, 'Browser'),
            builder.CardAction.imBack(session, 'Status Update', 'Status Update'),
          ]);
      });

      // Attach cards to message
      msg.attachments(attachments);

      // Send message with choice options
      session.send('Here are your active projects:');
      builder.Prompts.choice(
        session,
        msg,
        ['Status Update'],
        {
          maxRetries: 3,
          retryPrompt: 'Not a valid option',
        });
    },
    (session, results) => {
      session.replaceDialog('/engagementstatuscreate');
    },
  ]);
};
