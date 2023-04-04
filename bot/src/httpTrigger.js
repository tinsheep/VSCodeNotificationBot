
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
      displayName: 'My Incident 43-6',
      description: 'My Incident 43-6 description',
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
    // this can take a couple of minutes to fully complete.
    const location = team.headers.get('Location');
    // also get the teamId out of the location URL
    const teamId = location.match(/'([^']+)'/)[1];

    let teamStatus = "inProgress";
    while (teamStatus != "succeeded") {
      const checkStatusResponse = await graphClient.api(location).get();
      teamStatus = checkStatusResponse.status;
      await new Promise((resolve) => setTimeout(resolve, 5000));
    }
    console.log("Team created successfully!");

    // I need the driveId so i can post the incident report to the channel
    const {driveId, generalChannelId} = await getGeneralChannelDriveId(teamId, graphClient);
    console.log(`The driveId of the General channel for team ${teamId} is ${driveId}`);
    console.log(`The id of the General channel for team ${teamId} is ${generalChannelId}`);

    const filePath = 'C:\\OneNoteLocal\\Incident Report.pdf';

    //upload the incident report to the general channel
    const incidentReportUrl = await uploadFileToTeamsChannel(driveId, teamId, graphClient, filePath);
    console.log(`File uploaded to General channel: ${incidentReportUrl}`);


    //********************************************************/

    await target.sendAdaptiveCard(
      AdaptiveCards.declare(notificationTemplate).render({
        title: "New Incident Occurred!",
        appName: "Disaster Tech",
        description: `Welcome to the new incident team. Here is the Incident Action Plan:  ${target.type}`,
        notificationUrl: incidentReportUrl,
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

  async function getGeneralChannelDriveId(teamId, client) {
    const channels = await client.api(`/teams/${teamId}/channels`).filter(`displayName eq 'General'`).select('id').get();
    const generalChannelId = channels.value[0].id;
    const drive = await client.api(`/teams/${teamId}/channels/${generalChannelId}/filesFolder`).get();
    const driveId = drive.parentReference.driveId;
    return {driveId, generalChannelId};
  }


  async function uploadFileToTeamsChannel(driveId, teamId, graphClient, filePath) {
    const fs = require('fs');
    const path = require('path');
    const fileName = path.basename(filePath);
    const fileSize = fs.statSync(filePath).size;
    const fileContent = fs.createReadStream(filePath);

    //A Microsoft Graph API call to upload a local file to the Teams General channel
  
    try {
      const uploadSession = await graphClient.api(`/drives/${driveId}/root:/General/${fileName}:/createUploadSession`)
        .post({
          item: {
            '@microsoft.graph.conflictBehavior': 'rename',
            name: fileName,
          },
        });
   
      const uploadUrl = uploadSession.uploadUrl;
      const maxChunkSize = 320 * 1024; // 320 KB
      let start = 0;
      let end = maxChunkSize;
      let fileSlice;
     
      while (start < fileSize) {
        if (fileSize - end < 0) {
          end = fileSize;
        }
      
        fileSlice = Buffer.alloc(maxChunkSize);
        const bytesRead = fileContent.read(fileSlice, 0, maxChunkSize, start);
      
        const response = await fetch(uploadUrl, {
          method: 'PUT',
          headers: {
            'Content-Range': `bytes ${start}-${start + bytesRead - 1}/${fileSize}`,
          },
          body: fileSlice.slice(0, bytesRead),
        });
      
        start += bytesRead;
        end += maxChunkSize;
      }

      //Complete the upload
      const response = await fetch(uploadUrl, {
        method: 'POST',
        headers: {
          'Content-Length': 0,
        },
      });

      //Get the URL to the uploaded file


      const deepLink = `https://teams.microsoft.com/l/file/${encodeURIComponent(fileName)}/preview?groupId=${teamId}&tenantId=${appAuthConfig.tenantId}&channelId=${generalChannelId}`;
      const uploadedFileUrl = `https://teams.microsoft.com/_#/files/tab/${teamId}/${driveId}/${encodeURIComponent(fileName)}`;
      console.log(`File uploaded successfully to ${uploadedFileUrl}`);
      return deepLink;
    } catch (error) {
      console.error(`Error uploading file: ${error}`);
      throw error;
    }
  }

};
