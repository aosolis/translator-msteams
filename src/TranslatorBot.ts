import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import { TranslatorApi } from "./TranslatorApi";

// =========================================================
// Bot Setup
// =========================================================

export class TranslatorBot extends builder.UniversalBot {

    private loadSessionAsync: {(address: builder.IAddress): Promise<builder.Session>};
    private translator: TranslatorApi;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);
        this.set("persistConversationData", true);

        this.translator = botSettings.translator as TranslatorApi;

        // Handle invoke events
        this.loadSessionAsync = (address) => {
            return new Promise((resolve, reject) => {
                this.loadSession(address, (err: any, session: builder.Session) => {
                    if (err) {
                        reject(err);
                    } else {
                        resolve(session);
                    }
                });
            });
        };

        // Handle compose extension invokes
        let teamsConnector = this._connector as msteams.TeamsChatConnector;
        if (teamsConnector.onQuery) {
            teamsConnector.onQuery("search", (event, query, cb) => { this.onComposeExtensionQuery(event, query, cb); });
        }

        // Register default dialog
        this.dialog("/", async (session) => {
            let text = await this.translator.translateText(session.message.text, "it");
            session.endDialog(text);
        });
    }

    // Handle compose extension query invocation
    private async onComposeExtensionQuery(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event.address);
        if (session) {
            cb(new Error(), null, 500);
        }
    }
}
