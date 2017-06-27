const builder = require('botbuilder');
const request = require('request-promise-native');
const moment = require('moment');

async function getCrmId(userData) {
  return new Promise((resolve, reject) => {
    const options = {
      url: `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.2/systemusers?$filter=contains(domainname,%27${userData.upn}%27)&$select=systemuserid`,
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
  const url = `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.2/ee_projects%20?$select=ee_projectname,ee_engagementcurrentphase,modifiedon%20&$filter=(%20_ownerid_value%20eq%20${userData.crmId}%20or%20_ee_coowner1_value%20eq%20${userData.crmId}%20or%20_ee_coowner2_value%20eq%20${userData.crmId}%20or%20_ee_coowner3_value%20eq%20${userData.crmId}%20or%20_ee_coowner4_value%20eq%20${userData.crmId}%20or%20_ee_coowner5_value%20eq%20${userData.crmId}%20)%20and%20ee_engagementcurrentphase%20eq%20100000002`;

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

      // Ensure user actually has active project(s)
      if (engagements.value.length === 0) {
        session.send('Sorry but I cannot find any active projects for you.');
        session.replaceDialog('/');
      }

      // Persist engagements
      session.dialogData.engagements = engagements;

      // Create Carousel
      const msg = new builder.Message(session);
      msg.attachmentLayout(builder.AttachmentLayout.carousel);

      // Create Cards
      const attachments = engagements.value.map((engagement) => {
        // Abbreviate the overview
        const overview = engagement.ee_engagementoverview ? `${engagement.ee_engagementoverview.substr(0, 100)}...` : null;

        return new builder.HeroCard(session)
          .title(engagement.ee_projectname)
          .subtitle(`Last updated ${moment(engagement.modifiedon).fromNow()}`)
          .text(overview)
          .buttons([
            builder.CardAction.openUrl(session, `${process.env.MICROSOFT_RESOURCE_CRM}/main.aspx?etc=10096&extraqs=formid%3d33679b2b-bfd5-4e84-a63d-13aa63146ebb&id=%7b${engagement.ee_projectid}%7d&pagetype=entityrecord`, 'Browser'),
            builder.CardAction.imBack(session, `Status Update for ${engagement.ee_projectname}`, 'Status Update'),
          ]);
      });

      // Create list for options dialog
      const choices = engagements.value.map(engagement => `Status Update for ${engagement.ee_projectname}`);

      // Generate a mapping of project names to ids
      session.dialogData.projectIds = {};
      engagements.value.forEach((engagement) => {
        session.dialogData.projectIds[engagement.ee_projectname] = engagement.ee_projectid;
      });

      // Attach cards to message
      msg.attachments(attachments);

      // Send message with choice options
      session.send('Here are your active projects:');
      builder.Prompts.choice(
        session,
        msg,
        choices,
        {
          maxRetries: 3,
          retryPrompt: 'Not a valid option',
        });
    },
    (session, results) => {
      // Map project name to project id
      const projectId = session.dialogData.projectIds[results.response.entity.split('Status Update for ')[1]];

      // Pass ID as an arg to status creation dialog
      session.replaceDialog('/engagementstatuscreate', projectId);
    },
  ]).cancelAction('cancelSearch', 'Understood', {
    matches: /(cancel|back)/i,
  }).triggerAction({
    matches: /(my projects|update)/i,
  });
};
