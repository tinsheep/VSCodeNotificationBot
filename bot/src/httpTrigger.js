
const notificationTemplate = require("./adaptiveCards/notification-default.json");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { bot } = require("./internal/initialize");
//try this:
const { AppCredential, createMicrosoftGraphClientWithCredential } = require("@microsoft/teamsfx");

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
module.exports = async function (context, req) {
  for (const target of await bot.notification.installations()) {

    //try this: **********************************************
    const appAuthConfig = {
      authorityHost: process.env.M365_AUTHORITY_HOST,
      clientId: process.env.M365_CLIENT_ID,
      tenantId: process.env.M365_TENANT_ID,
      clientSecret: process.env.M365_CLIENT_SECRET
    }
    const appCredential = new AppCredential(appAuthConfig)
    //const token = appCredential.getToken()
    const graphClient = createMicrosoftGraphClientWithCredential(appCredential);

    //create Incident response team from a teamsTemplate

    const teamTemplate = {
      'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'' + process.env.TEAMS_TEMPLATE_ID  +  '\')',
      displayName: 'My Incident A',
      description: 'My Incident A description',
      members:[
          {
             '@odata.type': '#microsoft.graph.aadUserConversationMember',
             roles:[
                'owner'
             ],
             'user@odata.bind': 'https://graph.microsoft.com/v1.0/users/' + process.env.MOD_ID
          }
      ]
    }

    const team = await graphClient.api('/teams').post(teamTemplate);
    

    //********************************************************/

    await target.sendAdaptiveCard(
      AdaptiveCards.declare(notificationTemplate).render({
        title: "New Incident Occurred!",
        appName: "Disaster Tech",
        description: `Welcome to the new incident team. Here is the Incident Action Plan:  ${target.type}`,
        notificationUrl: "https://m365x501367.sharepoint.com/sites/MyIncident4/Shared%20Documents/General/469308021927-5522126379-ticket.pdf",
      })
    );
  }

  /****** To distinguish different target types ******/
  /** "Channel" means this bot is installed to a Team (default to notify General channel)
  if (target.type === NotificationTargetType.Channel) {
    // Directly notify the Team (to the default General channel)
    await target.sendAdaptiveCard(...);
    // List all channels in the Team then notify each channel
    const channels = await target.channels();
    for (const channel of channels) {
      await channel.sendAdaptiveCard(...);
    }
    // List all members in the Team then notify each member
    const members = await target.members();
    for (const member of members) {
      await member.sendAdaptiveCard(...);
    }
  }
  **/

  /** "Group" means this bot is installed to a Group Chat
  if (target.type === NotificationTargetType.Group) {
    // Directly notify the Group Chat
    await target.sendAdaptiveCard(...);
    // List all members in the Group Chat then notify each member
    const members = await target.members();
    for (const member of members) {
      await member.sendAdaptiveCard(...);
    }
  }
  **/

  /** "Person" means this bot is installed as a Personal app
  if (target.type === NotificationTargetType.Person) {
    // Directly notify the individual person
    await target.sendAdaptiveCard(...);
  }
  **/

  context.res = {};
};
