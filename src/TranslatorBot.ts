import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import { TranslatorApi, TranslationResult } from "./TranslatorApi";
import { Strings } from "./locale/locale";

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
            teamsConnector.onQuery("translate", (event, query, cb) => { this.onComposeExtensionQuery(event, query, cb); });
        }

        // Register default dialog
        this.dialog("/", async (session) => {
            let result = await this.translator.translateText(session.message.text, "it");
            session.endDialog(result[0].translatedText);
        });
    }

    // Handle compose extension query invocation
    private async onComposeExtensionQuery(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event.address);
        if (session) {
            let text = (query.parameters[0].name === "text") ? query.parameters[0].value : "";
            if (text) {
                try {
                    let translations = await this.translator.translateText(text, this.translator.getDefaultLanguages());
                    let response = msteams.ComposeExtensionResponse.result("list")
                        .attachments(translations
                            .filter(translation => translation.from !== translation.to)
                            .map(translation => this.createResult(session, translation)));
                    cb(null, response.toResponse());
                } catch (e) {
                    cb(null, this.createMessageResponse(session, Strings.error_translation));
                }
            } else {
                cb(null, this.createMessageResponse(session, Strings.error_notext));
            }
        } else {
            cb(new Error(), null, 500);
        }
    }

    private createMessageResponse(session: builder.Session, text: string): msteams.IComposeExtensionResponse {
        let response = new msteams.ComposeExtensionResponse("message");
        (response as any).data.composeExtension.text = session.gettext(text);
        return response.toResponse();
    }

    private createResult(session: builder.Session, translation: TranslationResult): msteams.ComposeExtensionAttachment {
        let card: msteams.ComposeExtensionAttachment = new builder.ThumbnailCard()
            .title(translation.translatedText)
            .text(translation.text)
            .toAttachment();
        card.preview = new builder.ThumbnailCard()
            .title(translation.translatedText)
            .text(session.gettext(translation.to))
            .toAttachment();
        return card;
    }
}
