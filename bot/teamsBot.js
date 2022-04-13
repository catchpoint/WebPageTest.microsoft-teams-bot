const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const WebPageTest = require("webpagetest");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const resultCard = require("./adaptiveCards/resultData.json");
const keyCard = require("./adaptiveCards/webpagetest-key.json");
const cardTools = require("@microsoft/adaptivecards-tools");
const { runTest, getLocations } = require("./utils/wptHelpers");

let url;
let key;
let isRunning = false;

let options = {
  firstViewOnly: true,
  runs: 1,
  pollResults: 5,
};

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    // record the likeCount
    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      //receving the values and running the test
      if (context._activity.value && context._activity.value.key) {
        key = context._activity.value.key;
      }
      if (context._activity.value && context._activity.value.url && isRunning === false) {
        url = context._activity.value.url;

        options = Object.assign(options, {
          emulateMobile: context._activity.value.isMobile == "true" ? true : false,
        });

        options = Object.assign(options, {
          location: context._activity.value.location,
        });
        options = Object.assign(options, {
          connectivity: context._activity.value.connectivity,
        });
        const wpt = new WebPageTest("www.webpagetest.org", key); // Your WPT API Key
        isRunning = true;
        await runTest(wpt, url, options)
          .then(async (test) => {
            if (test) {
              isRunning = false;
              this.resultDataObj = {
                testedurl: test.result.data.url,
                result: `https://webpagetest.org/result/${test.result.data.id}`,
                image: test.result.data.median.firstView.images.waterfall,
              };
              const card = cardTools.AdaptiveCards.declare(resultCard).render(this.resultDataObj);

              await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
              await next();
            }
          })
          .catch(async () => {
            await context.sendActivity("Please update your key, Use command 'key'");
            await next();
          });
      }

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "key": {
          const card = cardTools.AdaptiveCards.declare(keyCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "test": {
          if (key === undefined) {
            const card = cardTools.AdaptiveCards.declare(keyCard).render();
            await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          } else {
            const locationsResult = await getLocations(
              new WebPageTest("www.webpagetest.org"),
              options
            );
            const allLocations = locationsResult.result.response.data.location;

            ///////////////////////////// Creating Location Json Card ////////////////////////////////

            const arrayLoc = [];

            allLocations.forEach((loc) => {
              const titals = loc.Browsers.split(",").map(
                (item) => "Location: " + loc.Label + ", Browser: " + item
              );
              const values = loc.Browsers.split(",").map((item) => loc.location + ":" + item);

              for (let i = 0; i < titals.length; i++) {
                let tit = titals[i];
                let val = values[i];

                arrayLoc.push({ title: tit, value: val });
              }
            });

            const locationsJson = {
              $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
              type: "AdaptiveCard",
              version: "1.0",
              body: [
                {
                  type: "Input.Text",
                  id: "url",
                  style: "text",
                  label: "Enter URL",
                  isRequired: true,
                  errorMessage: "Required input",
                },
                {
                  type: "TextBlock",
                  text: "Connection Type",
                },
                {
                  type: "Input.ChoiceSet",
                  id: "connectivity",
                  style: "compact",
                  isMultiSelect: false,
                  value: "",
                  choices: [
                    {
                      title: "DSL - 1.5 Mbps down, 384 Kbps up, 50 ms first-hop RTT",
                      value: "DSL",
                    },
                    {
                      title: "Cable - 5 Mbps down, 1 Mbps up, 28ms first-hop RTT",
                      value: "Cable",
                    },
                    {
                      title: "FIOS - 20 Mbps down, 5 Mbps up, 4 ms first-hop RTT",
                      value: "FIOS",
                    },
                    {
                      title: "Dial - 49 Kbps down, 30 Kbps up, 120 ms first-hop RTT",
                      value: "Dial",
                    },
                    {
                      title: "Edge - 240 Kbps down, 200 Kbps up, 840 ms first-hop RTT",
                      value: "Edge",
                    },
                    {
                      title: "2G - 280 Kbps down, 256 Kbps up, 800 ms first-hop RTT",
                      value: "2G",
                    },
                    {
                      title: "3GSlow - 400 Kbps down and up, 400 ms first-hop RTT",
                      value: "3GSlow",
                    },
                    {
                      title: "3G - 1.6 Mbps down, 768 Kbps up, 300 ms first-hop RTT",
                      value: "3G",
                    },
                    {
                      title: "3GFast - 1.6 Mbps down, 768 Kbps up, 150 ms first-hop RTT",
                      value: "3GFast",
                    },
                    {
                      title: "4G - 9 Mbps down and up, 170 ms first-hop RTT",
                      value: "4G",
                    },
                    {
                      title: "LTE - 12 Mbps down and up, 70 ms first-hop RTT",
                      value: "LTE",
                    },
                    {
                      title: "Native - No synthetic traffic shaping applied",
                      value: "Native",
                    },
                  ],
                },
                {
                  type: "TextBlock",
                  text: "Select Location & Browser",
                },
                {
                  type: "Input.ChoiceSet",
                  id: "location",
                  style: "compact",
                  isMultiSelect: false,
                  value: "",
                  choices: arrayLoc,
                },
                {
                  type: "TextBlock",
                  text: "Emulate Mobile?",
                },
                {
                  type: "Input.ChoiceSet",
                  id: "isMobile",
                  style: "expanded",
                  value: "false",
                  choices: [
                    {
                      title: "Yes",
                      value: "true",
                    },
                    {
                      title: "No",
                      value: "false",
                    },
                  ],
                },
              ],
              actions: [
                {
                  type: "Action.Submit",
                  title: "OK",
                  data: {
                    msteams: {
                      type: "messageBack",
                      displayText: "Test Submitted!! Result will be posted here soon",
                    },
                  },
                },
              ],
            };

            /////////////////////////////////////////////////////////////////////////////////////

            const card = cardTools.AdaptiveCards.declare(locationsJson).render();
            await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          }

          break;
        }
      }

      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(context, invokeValue) {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200 };
    }
  }

  // Messaging extension Code
  // Action.
  handleTeamsMessagingExtensionSubmitAction(context, action) {
    switch (action.commandId) {
      case "createCard":
        return createCardCommand(context, action);
      case "shareMessage":
        return shareMessageCommand(context, action);
      default:
        throw new Error("NotImplemented");
    }
  }

  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const heroCard = CardFactory.heroCard(obj.package.name);
      const preview = CardFactory.heroCard(obj.package.name);
      preview.content.tap = {
        type: "invoke",
        value: { name: obj.package.name, description: obj.package.description },
      };
      const attachment = { ...heroCard, preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  async handleTeamsMessagingExtensionSelectItem(context, obj) {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }

  // Link Unfurling.
  handleTeamsAppBasedLinkQuery(context, query) {
    const attachment = CardFactory.thumbnailCard("Thumbnail Card", query.url, [query.url]);

    const result = {
      attachmentLayout: "list",
      type: "result",
      attachments: [attachment],
    };

    const response = {
      composeExtension: result,
    };
    return response;
  }
}

function createCardCommand(context, action) {
  // The user has chosen to create a card by choosing the 'Create Card' context menu command.
  const data = action.data;
  const heroCard = CardFactory.heroCard(data.title, data.text);
  heroCard.content.subtitle = data.subTitle;
  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

function shareMessageCommand(context, action) {
  // The user has chosen to share a message by choosing the 'Share Message' context menu command.
  let userName = "unknown";
  if (
    action.messagePayload &&
    action.messagePayload.from &&
    action.messagePayload.from.user &&
    action.messagePayload.from.user.displayName
  ) {
    userName = action.messagePayload.from.user.displayName;
  }

  // This Messaging Extension example allows the user to check a box to include an image with the
  // shared message.  This demonstrates sending custom parameters along with the message payload.
  let images = [];
  const includeImage = action.data.includeImage;
  if (includeImage === "true") {
    images = [
      "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU",
    ];
  }
  const heroCard = CardFactory.heroCard(
    `${userName} originally sent this message:`,
    action.messagePayload.body.content,
    images
  );

  if (
    action.messagePayload &&
    action.messagePayload.attachment &&
    action.messagePayload.attachments.length > 0
  ) {
    // This sample does not add the MessagePayload Attachments.  This is left as an
    // exercise for the user.
    heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
  }

  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

module.exports.TeamsBot = TeamsBot;
