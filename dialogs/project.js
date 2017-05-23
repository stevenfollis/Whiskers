'use strict';

const builder = require('botbuilder');
const request = require('request-promise-native');
const emoji = require('node-emoji');
const moment = require('moment');

module.exports = {
    Label: 'Project Information',
    Dialog: [
        (session) => {

            // Create message with card
            var msg = new builder.Message(session)
                .attachmentLayout(builder.AttachmentLayout.list)
                .attachments([
                    new builder.HeroCard(session)
                        .title("What would you like to do?")
                        .buttons([
                            builder.CardAction.imBack(session, 'My Projects', 'My Projects'),
                            builder.CardAction.imBack(session, 'Search', 'Search')
                        ]),
                ]);

            // Send message with choice options
            builder.Prompts.choice(
                session,
                msg,
                ['My Projects', 'Search'],
                {
                    maxRetries: 3,
                    retryPrompt: 'Not a valid option'
                });

        },
        async (session, results, next) => {
            switch (results.response.entity) {
                case "My Projects":

                    // Get CRM ID
                    session.userData.crmId = await getCrmId(session.userData.accessTokenCRM);

                    // Get projects
                    const projects = await getProjects(session.userData);

                    // Create carousel
                    var msg = new builder.Message(session);
                    msg.attachmentLayout(builder.AttachmentLayout.carousel)

                    // Create Cards
                    let attachments = projects.value.map((project, i) => {
                        return new builder.HeroCard(session)
                            .title(project.ee_projectname)
                            .subtitle(`Last updated ${moment(project.modifiedon).fromNow()}`)
                            .text("Select this project with the button")
                            .buttons([
                                builder.CardAction.imBack(session, "New project update", "Status Update")
                            ]);
                    });

                    // Attach cards to message
                    msg.attachments(attachments);

                    // Send message ending the dialog
                    session.send(msg).endDialog();

                    break;

                case "Search":
                    // session.replaceDialog("/search")
                    session.endDialog('Search coming soon!');
                default:
                    session.replaceDialog("/");
                    break;
            }
        }
    ]
}

async function getCrmId(accessTokenCRM) {
    return new Promise((resolve, reject) => {

        var options = {
            url: `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.1/systemusers?$filter=contains(domainname,%27stfollis@microsoft.com%27)&$select=systemuserid`,
            headers: { 'Authorization': accessTokenCRM },
            json: true
        };

        request.get(options).then((result) => {
            console.log(`Retrieved CRM User ID`);
            resolve(result.value[0].systemuserid);
        });;

    });
}

async function getProjects(userData) {

    // Create OData URL for the user's active projects
    const url = `${process.env.MICROSOFT_RESOURCE_CRM}/api/data/v8.1/ee_projects%20?$select=ee_projectname,ee_engagementcurrentphase,modifiedon%20&$filter=(%20_ownerid_value%20eq%20${userData.crmId}%20or%20_ee_coowner1_value%20eq%20${userData.crmId}%20or%20_ee_coowner2_value%20eq%20${userData.crmId}%20or%20_ee_coowner3_value%20eq%20${userData.crmId}%20or%20_ee_coowner4_value%20eq%20${userData.crmId}%20or%20_ee_coowner5_value%20eq%20${userData.crmId}%20)%20and%20ee_engagementcurrentphase%20eq%20100000002`;

    var options = {
        url: url,
        headers: { 'Authorization': userData.accessTokenCRM },
        json: true
    };

    return await request.get(options);

}