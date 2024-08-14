import {
  CardFactory,
  MessageFactory,
  TeamsActivityHandler
} from "botbuilder";
import config from "./config";

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context) => {
      console.log("Running with Message Activity.");

      const signInLink = `https://${config.botDomain}/auth-start.html`;
      const oauthCard = CardFactory.adaptiveCard({
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.0",
        body: [
          {
            type: "TextBlock",
            text: "Sign in to your account",
          },
        ],
        actions: [
          {
            type: "Action.Submit",
            title: "Sign in",
            data: {
              "msteams": {
                "type": "signin",
                "value": signInLink,
              }
            }
          },
        ],
      })
      const message = MessageFactory.attachment(oauthCard);
      await context.sendActivity(message);
    });
  }
}
