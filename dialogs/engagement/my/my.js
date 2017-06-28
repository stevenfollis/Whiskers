const builder = require('botbuilder');
const request = require('request-promise-native');
const moment = require('moment');

async function getCrmId(userData) {
  const options = {
    method: 'GET',
    url: `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.2/systemusers?$filter=contains(domainname, '${userData.upn}')&$select=ownerid`,
    headers: { Authorization: userData.accessTokenCRM },
    json: true,
  };

  // Query the API
  return request.get(options);
}

async function getEngagements(userData) {
  // Create OData URL for the user's active engagements
  const url = [
    `${process.env.MICROSOFT_RESOURCE_CRM}`,
    '/api/data/v8.2/ee_projects',
    '?$select=ee_projectname,ee_engagementcurrentphase,modifiedon',
    `&$filter=(_ownerid_value eq ${userData.crmId} or _ee_coowner1_value eq ${userData.crmId} or _ee_coowner2_value eq ${userData.crmId} or _ee_coowner3_value eq ${userData.crmId} or _ee_coowner4_value eq ${userData.crmId} or _ee_coowner5_value eq ${userData.crmId} ) and ee_engagementcurrentphase eq 100000002`,
    '&$orderby=modifiedon desc'].join('');

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
      try {
        // Get CRM ID
        const crmId = await getCrmId(session.userData);
        session.userData.crmId = crmId.value[0].ownerid;

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

          // Create Hero Card
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
      } catch (error) {
        console.log(error);
        session.send('Uh oh...had a problem getting your engagements. Sorry!');
        session.replaceDialog('/');
      }
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
