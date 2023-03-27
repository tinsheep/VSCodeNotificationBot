
const notificationTemplate = require("./adaptiveCards/notification-default.json");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { bot } = require("./internal/initialize");
const { AppCredential, createMicrosoftGraphClientWithCredential } = require("@microsoft/teamsfx");
const { ResponseType } = require('@microsoft/microsoft-graph-client');

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
    const graphClient = createMicrosoftGraphClientWithCredential(appCredential);

       
    //create Incident response team from a teamsTemplate

    const teamTemplate = {
      'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'' + process.env.TEAMS_TEMPLATE_ID  +  '\')',
      displayName: 'My Incident 327-1',
      description: 'My Incident 327-1 description',
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

    const team = await graphClient
        .api('/teams')
        .responseType(ResponseType.RAW)
        .post(teamTemplate);
    console.log(team.headers.get('client-request-id'));

    // get the URL value where we can make the call to check if the asynchonous operation to create the Team is complete.
    // this can take a minute or so
    const location = team.headers.get('Location')

    let teamStatus = "inProgress";
    while (teamStatus != "succeeded") {
      const checkStatusResponse = await graphClient.api(location).get();
      teamStatus = checkStatusResponse.status;
      await new Promise((resolve) => setTimeout(resolve, 5000));
    }
    console.log("Team created successfully!");


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
