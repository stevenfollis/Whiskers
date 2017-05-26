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

async function getProjects(userData) {
  // Create OData URL for the user's active projects
  const url = `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.1/ee_projects%20?$select=ee_projectname,ee_engagementcurrentphase,modifiedon%20&$filter=(%20_ownerid_value%20eq%20${userData.crmId}%20or%20_ee_coowner1_value%20eq%20${userData.crmId}%20or%20_ee_coowner2_value%20eq%20${userData.crmId}%20or%20_ee_coowner3_value%20eq%20${userData.crmId}%20or%20_ee_coowner4_value%20eq%20${userData.crmId}%20or%20_ee_coowner5_value%20eq%20${userData.crmId}%20)%20and%20ee_engagementcurrentphase%20eq%20100000002`;

  const options = {
    url,
    headers: { Authorization: userData.accessTokenCRM },
    json: true,
  };

  return request.get(options);
}

module.exports = {
  Label: 'Project Information',
  Dialog: [
    (session) => {
      // Create message with card
      const msg = new builder.Message(session)
        .attachmentLayout(builder.AttachmentLayout.list)
        .attachments([
          new builder.HeroCard(session)
            .title('What would you like to do?')
            .buttons([
              builder.CardAction.imBack(session, 'My Projects', 'My Projects'),
              builder.CardAction.imBack(session, 'Search', 'Search'),
              builder.CardAction.imBack(session, 'Logout', 'Logout'),
            ]),
        ]);

      // Send message with choice options
      builder.Prompts.choice(
        session,
        msg,
        ['My Projects', 'Search', 'Logout'],
        {
          maxRetries: 3,
          retryPrompt: 'Not a valid option',
        });
    },
    async (session, results) => {
      switch (results.response.entity) {

        case 'My Projects': {
          // Get CRM ID
          session.userData.crmId = await getCrmId(session.userData);

          // Get Projects
          const projects = await getProjects(session.userData);

          // Create Carousel
          const msg = new builder.Message(session);
          msg.attachmentLayout(builder.AttachmentLayout.carousel);

          // Create Cards
          const attachments = projects.value.map((project) => {
            return new builder.HeroCard(session)
              .title(project.ee_projectname)
              .subtitle(`Last updated ${moment(project.modifiedon).fromNow()}`)
              .text('Select this project with the button')
              .buttons([
                builder.CardAction.imBack(session, 'Status Update', 'Status Update'),
              ]);
          });

          // Attach cards to message
          msg.attachments(attachments);

          // Send message with choice options
          builder.Prompts.choice(
            session,
            msg,
            ['Status Update'],
            {
              maxRetries: 3,
              retryPrompt: 'Not a valid option',
            });

          break;
        }

        case 'Search': {
          // session.replaceDialog("/search")
          session.send('Search coming soon!');
          session.replaceDialog('/');
          break;
        }

        case 'Logout': {
          // Run logout dialog
          session.beginDialog('/logout');

          // End project dialog
          session.endDialog();
          break;
        }

        default: {
          session.replaceDialog('/');
          break;
        }
      }
    },
    (session, result) => {
      switch (result.response.entity) {
        case 'Status Update': {
          session.replaceDialog('/projectstatuscreate');
          break;
        }
        default: {
          session.endDialog();
          break;
        }
      }
    },
  ],
};
