const builder = require('botbuilder');
const request = require('request-promise-native');
const emoji = require('node-emoji');
const moment = require('moment');

async function searchEngagements(userData, query) {
  return new Promise((resolve, reject) => {
    const options = {
      url: `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.1/ee_projects?$filter=contains(ee_projectname,'${query}')&$orderby=modifiedon desc&$top=10`,
      headers: { Authorization: userData.accessTokenCRM },
      json: true,
    };

    request
      .get(options)
      .then((result) => {
        resolve(result.value);
      })
      .catch((error) => {
        reject(error);
      });
  });
}

module.exports = {
  Label: 'Search',
  Dialog: [
    (session) => {
      builder.Prompts.text(session, 'What query should I use?');
    },
    async (session, results) => {
      // Get query string
      const query = results.response;

      // Get engagements
      const engagements = await searchEngagements(session.userData, query);

      if (engagements.length === 0) {
        session.send(`Sorry but no engagements matched your query. ${emoji.get('hushed')} `);
        session.replaceDialog('/');
      } else {
        // Create Carousel
        const msg = new builder.Message(session);
        msg.attachmentLayout(builder.AttachmentLayout.carousel);

        // Create Cards
        const attachments = engagements.map((engagement) => {
          return new builder.HeroCard(session)
            .title(engagement.ee_projectname)
            .subtitle(`Last updated ${moment(engagement.modifiedon).fromNow()}`)
            .text('Lorem ipsum')
            .buttons([
              builder.CardAction.openUrl(session, `${process.env.MICROSOFT_RESOURCE_CRM}/main.aspx?etc=10096&extraqs=formid=33679b2b-bfd5-4e84-a63d-13aa63146ebb&pagetype=entityrecord&id={${engagement.ee_projectid}}`, 'Browser'),
            ]);
        });

        // Attach cards to message
        msg.attachments(attachments);

        // Respond
        session.send(`Found the following ${engagements.length} engagements:`);
        session.endDialog(msg);
      }
    },
  ],
};
