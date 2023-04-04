"use strict";

var notificationTemplate = require("./adaptiveCards/notification-default.json");

var _require = require("@microsoft/adaptivecards-tools"),
    AdaptiveCards = _require.AdaptiveCards;

var _require2 = require("./internal/initialize"),
    bot = _require2.bot;

var _require3 = require("@microsoft/teamsfx"),
    AppCredential = _require3.AppCredential,
    createMicrosoftGraphClientWithCredential = _require3.createMicrosoftGraphClientWithCredential;

var _require4 = require('@microsoft/microsoft-graph-client'),
    ResponseType = _require4.ResponseType; // HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.


module.exports = function _callee(context, req) {
  var _iteratorNormalCompletion, _didIteratorError, _iteratorError, _iterator, _step, target, _appAuthConfig, appCredential, graphClient, teamTemplate, team, location, teamId, teamStatus, checkStatusResponse, _ref, driveId, _generalChannelId, filePath, incidentReportUrl, getGeneralChannelDriveId, uploadFileToTeamsChannel;

  return regeneratorRuntime.async(function _callee$(_context3) {
    while (1) {
      switch (_context3.prev = _context3.next) {
        case 0:
          uploadFileToTeamsChannel = function _ref3(driveId, teamId, graphClient, filePath) {
            var fs, path, fileName, fileSize, fileContent, uploadSession, uploadUrl, maxChunkSize, start, end, fileSlice, bytesRead, _response, response, deepLink, uploadedFileUrl;

            return regeneratorRuntime.async(function uploadFileToTeamsChannel$(_context2) {
              while (1) {
                switch (_context2.prev = _context2.next) {
                  case 0:
                    fs = require('fs');
                    path = require('path');
                    fileName = path.basename(filePath);
                    fileSize = fs.statSync(filePath).size;
                    fileContent = fs.createReadStream(filePath); //A Microsoft Graph API call to upload a local file to the Teams General channel

                    _context2.prev = 5;
                    _context2.next = 8;
                    return regeneratorRuntime.awrap(graphClient.api("/drives/".concat(driveId, "/root:/General/").concat(fileName, ":/createUploadSession")).post({
                      item: {
                        '@microsoft.graph.conflictBehavior': 'rename',
                        name: fileName
                      }
                    }));

                  case 8:
                    uploadSession = _context2.sent;
                    uploadUrl = uploadSession.uploadUrl;
                    maxChunkSize = 320 * 1024; // 320 KB

                    start = 0;
                    end = maxChunkSize;

                  case 13:
                    if (!(start < fileSize)) {
                      _context2.next = 24;
                      break;
                    }

                    if (fileSize - end < 0) {
                      end = fileSize;
                    }

                    fileSlice = Buffer.alloc(maxChunkSize);
                    bytesRead = fileContent.read(fileSlice, 0, maxChunkSize, start);
                    _context2.next = 19;
                    return regeneratorRuntime.awrap(fetch(uploadUrl, {
                      method: 'PUT',
                      headers: {
                        'Content-Range': "bytes ".concat(start, "-").concat(start + bytesRead - 1, "/").concat(fileSize)
                      },
                      body: fileSlice.slice(0, bytesRead)
                    }));

                  case 19:
                    _response = _context2.sent;
                    start += bytesRead;
                    end += maxChunkSize;
                    _context2.next = 13;
                    break;

                  case 24:
                    _context2.next = 26;
                    return regeneratorRuntime.awrap(fetch(uploadUrl, {
                      method: 'POST',
                      headers: {
                        'Content-Length': 0
                      }
                    }));

                  case 26:
                    response = _context2.sent;
                    //Get the URL to the uploaded file
                    deepLink = "https://teams.microsoft.com/l/file/".concat(encodeURIComponent(fileName), "/preview?groupId=").concat(teamId, "&tenantId=").concat(appAuthConfig.tenantId, "&channelId=").concat(generalChannelId);
                    uploadedFileUrl = "https://teams.microsoft.com/_#/files/tab/".concat(teamId, "/").concat(driveId, "/").concat(encodeURIComponent(fileName));
                    console.log("File uploaded successfully to ".concat(uploadedFileUrl));
                    return _context2.abrupt("return", deepLink);

                  case 33:
                    _context2.prev = 33;
                    _context2.t0 = _context2["catch"](5);
                    console.error("Error uploading file: ".concat(_context2.t0));
                    throw _context2.t0;

                  case 37:
                  case "end":
                    return _context2.stop();
                }
              }
            }, null, null, [[5, 33]]);
          };

          getGeneralChannelDriveId = function _ref2(teamId, client) {
            var channels, generalChannelId, drive, driveId;
            return regeneratorRuntime.async(function getGeneralChannelDriveId$(_context) {
              while (1) {
                switch (_context.prev = _context.next) {
                  case 0:
                    _context.next = 2;
                    return regeneratorRuntime.awrap(client.api("/teams/".concat(teamId, "/channels")).filter("displayName eq 'General'").select('id').get());

                  case 2:
                    channels = _context.sent;
                    generalChannelId = channels.value[0].id;
                    _context.next = 6;
                    return regeneratorRuntime.awrap(client.api("/teams/".concat(teamId, "/channels/").concat(generalChannelId, "/filesFolder")).get());

                  case 6:
                    drive = _context.sent;
                    driveId = drive.parentReference.driveId;
                    return _context.abrupt("return", {
                      driveId: driveId,
                      generalChannelId: generalChannelId
                    });

                  case 9:
                  case "end":
                    return _context.stop();
                }
              }
            });
          };

          _iteratorNormalCompletion = true;
          _didIteratorError = false;
          _iteratorError = undefined;
          _context3.prev = 5;
          _context3.next = 8;
          return regeneratorRuntime.awrap(bot.notification.installations());

        case 8:
          _context3.t0 = Symbol.iterator;
          _iterator = _context3.sent[_context3.t0]();

        case 10:
          if (_iteratorNormalCompletion = (_step = _iterator.next()).done) {
            _context3.next = 50;
            break;
          }

          target = _step.value;
          //try this: **********************************************
          _appAuthConfig = {
            authorityHost: process.env.M365_AUTHORITY_HOST,
            clientId: process.env.M365_CLIENT_ID,
            tenantId: process.env.M365_TENANT_ID,
            clientSecret: process.env.M365_CLIENT_SECRET
          };
          appCredential = new AppCredential(_appAuthConfig);
          graphClient = createMicrosoftGraphClientWithCredential(appCredential); //create Incident response team from a teamsTemplate

          teamTemplate = {
            'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'' + process.env.TEAMS_TEMPLATE_ID + '\')',
            displayName: 'My Incident 43-6',
            description: 'My Incident 43-6 description',
            members: [{
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': 'https://graph.microsoft.com/v1.0/users/' + process.env.MOD_ID
            }]
          };
          _context3.next = 18;
          return regeneratorRuntime.awrap(graphClient.api('/teams').responseType(ResponseType.RAW).post(teamTemplate));

        case 18:
          team = _context3.sent;
          console.log(team.headers.get('client-request-id')); // get the URL value where we can make the call to check if the asynchonous operation to create the Team is complete.
          // this can take a couple of minutes to fully complete.

          location = team.headers.get('Location'); // also get the teamId out of the location URL

          teamId = location.match(/'([^']+)'/)[1];
          teamStatus = "inProgress";

        case 23:
          if (!(teamStatus != "succeeded")) {
            _context3.next = 32;
            break;
          }

          _context3.next = 26;
          return regeneratorRuntime.awrap(graphClient.api(location).get());

        case 26:
          checkStatusResponse = _context3.sent;
          teamStatus = checkStatusResponse.status;
          _context3.next = 30;
          return regeneratorRuntime.awrap(new Promise(function (resolve) {
            return setTimeout(resolve, 5000);
          }));

        case 30:
          _context3.next = 23;
          break;

        case 32:
          console.log("Team created successfully!"); // I need the driveId so i can post the incident report to the channel

          _context3.next = 35;
          return regeneratorRuntime.awrap(getGeneralChannelDriveId(teamId, graphClient));

        case 35:
          _ref = _context3.sent;
          driveId = _ref.driveId;
          _generalChannelId = _ref.generalChannelId;
          console.log("The driveId of the General channel for team ".concat(teamId, " is ").concat(driveId));
          console.log("The id of the General channel for team ".concat(teamId, " is ").concat(_generalChannelId));
          filePath = 'C:\\OneNoteLocal\\Incident Report.pdf'; //upload the incident report to the general channel

          _context3.next = 43;
          return regeneratorRuntime.awrap(uploadFileToTeamsChannel(driveId, teamId, graphClient, filePath));

        case 43:
          incidentReportUrl = _context3.sent;
          console.log("File uploaded to General channel: ".concat(incidentReportUrl)); //********************************************************/

          _context3.next = 47;
          return regeneratorRuntime.awrap(target.sendAdaptiveCard(AdaptiveCards.declare(notificationTemplate).render({
            title: "New Incident Occurred!",
            appName: "Disaster Tech",
            description: "Welcome to the new incident team. Here is the Incident Action Plan:  ".concat(target.type),
            notificationUrl: incidentReportUrl
          })));

        case 47:
          _iteratorNormalCompletion = true;
          _context3.next = 10;
          break;

        case 50:
          _context3.next = 56;
          break;

        case 52:
          _context3.prev = 52;
          _context3.t1 = _context3["catch"](5);
          _didIteratorError = true;
          _iteratorError = _context3.t1;

        case 56:
          _context3.prev = 56;
          _context3.prev = 57;

          if (!_iteratorNormalCompletion && _iterator["return"] != null) {
            _iterator["return"]();
          }

        case 59:
          _context3.prev = 59;

          if (!_didIteratorError) {
            _context3.next = 62;
            break;
          }

          throw _iteratorError;

        case 62:
          return _context3.finish(59);

        case 63:
          return _context3.finish(56);

        case 64:
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

        case 65:
        case "end":
          return _context3.stop();
      }
    }
  }, null, null, [[5, 52, 56, 64], [57,, 59, 63]]);
};
//# sourceMappingURL=httpTrigger.dev.js.map
