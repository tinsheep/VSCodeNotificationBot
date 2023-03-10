"use strict";

var notificationTemplate = require("./adaptiveCards/notification-default.json");

var _require = require("@microsoft/adaptivecards-tools"),
    AdaptiveCards = _require.AdaptiveCards;

var _require2 = require("./internal/initialize"),
    bot = _require2.bot; //try this:


var _require3 = require("@microsoft/teamsfx"),
    AppCredential = _require3.AppCredential,
    createMicrosoftGraphClientWithCredential = _require3.createMicrosoftGraphClientWithCredential; // HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.


module.exports = function _callee(context, req) {
  var _iteratorNormalCompletion, _didIteratorError, _iteratorError, _iterator, _step, target, appAuthConfig, appCredential, graphClient, teamTemplate, team;

  return regeneratorRuntime.async(function _callee$(_context) {
    while (1) {
      switch (_context.prev = _context.next) {
        case 0:
          _iteratorNormalCompletion = true;
          _didIteratorError = false;
          _iteratorError = undefined;
          _context.prev = 3;
          _context.next = 6;
          return regeneratorRuntime.awrap(bot.notification.installations());

        case 6:
          _context.t0 = Symbol.iterator;
          _iterator = _context.sent[_context.t0]();

        case 8:
          if (_iteratorNormalCompletion = (_step = _iterator.next()).done) {
            _context.next = 22;
            break;
          }

          target = _step.value;
          //try this: **********************************************
          appAuthConfig = {
            authorityHost: process.env.M365_AUTHORITY_HOST,
            clientId: process.env.M365_CLIENT_ID,
            tenantId: process.env.M365_TENANT_ID,
            clientSecret: process.env.M365_CLIENT_SECRET
          };
          appCredential = new AppCredential(appAuthConfig); //const token = appCredential.getToken()

          graphClient = createMicrosoftGraphClientWithCredential(appCredential); //create Incident response team from a teamsTemplate

          teamTemplate = {
            'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'' + process.env.TEAMS_TEMPLATE_ID + '\')',
            displayName: 'My Incident A',
            description: 'My Incident A description',
            members: [{
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': 'https://graph.microsoft.com/v1.0/users/' + process.env.MOD_ID
            }]
          };
          _context.next = 16;
          return regeneratorRuntime.awrap(graphClient.api('/teams').post(teamTemplate));

        case 16:
          team = _context.sent;
          _context.next = 19;
          return regeneratorRuntime.awrap(target.sendAdaptiveCard(AdaptiveCards.declare(notificationTemplate).render({
            title: "New Incident Occurred!",
            appName: "Disaster Tech",
            description: "Welcome to the new incident team. Here is the Incident Action Plan:  ".concat(target.type),
            notificationUrl: "https://m365x501367.sharepoint.com/sites/MyIncident4/Shared%20Documents/General/469308021927-5522126379-ticket.pdf"
          })));

        case 19:
          _iteratorNormalCompletion = true;
          _context.next = 8;
          break;

        case 22:
          _context.next = 28;
          break;

        case 24:
          _context.prev = 24;
          _context.t1 = _context["catch"](3);
          _didIteratorError = true;
          _iteratorError = _context.t1;

        case 28:
          _context.prev = 28;
          _context.prev = 29;

          if (!_iteratorNormalCompletion && _iterator["return"] != null) {
            _iterator["return"]();
          }

        case 31:
          _context.prev = 31;

          if (!_didIteratorError) {
            _context.next = 34;
            break;
          }

          throw _iteratorError;

        case 34:
          return _context.finish(31);

        case 35:
          return _context.finish(28);

        case 36:
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

        case 37:
        case "end":
          return _context.stop();
      }
    }
  }, null, null, [[3, 24, 28, 36], [29,, 31, 35]]);
};
//# sourceMappingURL=httpTrigger.dev.js.map
