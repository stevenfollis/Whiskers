const builder = require('botbuilder');
const request = require('request-promise-native');
const emoji = require('node-emoji');
const moment = require('moment');

/**
 * Search the CRM API with a filter for a given query
 * Returns up to 10 projects sorted by recently modified
 * @param {object} userData
 * @param {string} query
 */
async function searchEngagements(userData, query) {
  const options = {
    method: 'GET',
    url: `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.2/ee_projects?$filter=contains(ee_projectname,'${query}')&$orderby=modifiedon desc&$top=10`,
    headers: { Authorization: userData.accessTokenCRM },
    json: true,
  };

  return request(options);
}

module.exports = (bot) => {
  bot.dialog('/search', [
    (session) => {
      builder.Prompts.text(session, `What would you like to search for? ${emoji.get('mag')}`);
    },
    async (session, results) => {
      // Get query string
      const query = results.response;

      try {
        // Get engagements
        const engagements = await searchEngagements(session.userData, query);

        if (engagements.value.length === 0) {
          session.send(`Sorry, but no engagements matched your query. ${emoji.get('hushed')} `);
          session.replaceDialog('/');
        } else {
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
              ]);
          });

          // Attach cards to message
          msg.attachments(attachments);

          // Respond
          session.send(`Found the following ${engagements.value.length} engagements:`);
          session.send(msg);
          session.replaceDialog('/');
        }
      } catch (error) {
        console.log(error);
        session.endDialog('Oops looks like there was an error completing your search');
      }
    },
  ]).cancelAction('cancelSearch', 'Search Canceled', {
    matches: /(cancel|back)/i,
  }).triggerAction({
    matches: /search/i,
  });
};
