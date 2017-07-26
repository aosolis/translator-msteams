import * as config from "config";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
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
                        winston.error("Failed to load session", { error: err, address: address });
                        reject(err);
                    } else if (!session) {
                        winston.error("Loaded null session", { address: address });
                        reject(new Error("Failed to load session"));
                    } else {
                        resolve(session);
                    }
                });
            });
        };

        // Handle compose extension invokes
        let teamsConnector = this._connector as msteams.TeamsChatConnector;
        if (teamsConnector.onQuery) {
            teamsConnector.onQuery("translate", async (event, query, cb) => {
                try {
                    await this.handleTranslateQuery(event, query, cb);
                } catch (e) {
                    winston.error("Translate handler failed", e);
                    cb(e, null, 500);
                }
            });
        }
        if (teamsConnector.onQuerySettingsUrl) {
            teamsConnector.onQuerySettingsUrl(async (event, query, cb) => {
                try {
                    await this.handleQuerySettingsUrl(event, query, cb);
                } catch (e) {
                    winston.error("Query settings url handler failed", e);
                    cb(e, null, 500);
                }
            });
        }
        if (teamsConnector.onSettingsUpdate) {
            teamsConnector.onSettingsUpdate(async (event, query, cb) => {
                try {
                    await this.handleSettingsUpdate(event, query, cb);
                } catch (e) {
                    winston.error("Settings update handler failed", e);
                    cb(e, null, 500);
                }
            });
        }

        // Handle generic invokes
        teamsConnector.onInvoke((event, cb) => { this.onInvoke(event, cb); });

        // Register default dialog
        this.dialog("/", async (session) => {
            let result = await this.translator.translateText(session.message.text, "it");
            session.endDialog(result[0].translatedText);
        });
    }

    // Handle compose extension query invocation
    private async handleTranslateQuery(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event.address);

        let text = (query.parameters[0].name === "text") ? query.parameters[0].value : "";

        // Handle setting state
        let incomingSettings = query.state;
        if (incomingSettings) {
            this.updateSettings(session, incomingSettings);
            text = "";
        }

        let translationLanguages = this.getTranslationLanguages(session);

        if ((text === "settings") && config.get("features.allowConfigurationViaQuery")) {
            // Provide a way to get to settings for client versions that don't support canUpdateConfiguration
            cb(null, this.getConfigurationResponse(session, translationLanguages));
        } else if (text) {
            try {
                let translations = await this.translator.translateText(text, translationLanguages);
                let response = msteams.ComposeExtensionResponse.result("list")
                    .attachments(translations
                        .filter(translation => translation.from !== translation.to)
                        .map(translation => this.createResult(session, translation)));
                cb(null, response.toResponse());
            } catch (e) {
                winston.error("Failed to get translations", e);
                cb(null, this.createMessageResponse(session, Strings.error_translation));
            }
        } else {
            cb(null, this.createMessageResponse(session, Strings.error_notext));
        }
    }

    // Handle compose extension query settings url
    private async handleQuerySettingsUrl(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event.address);
        cb(null, this.getConfigurationResponse(session));
    }

    // Handle compose extension query settings url
    private async handleSettingsUpdate(event: builder.IEvent, query: msteams.ComposeExtensionQuery, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event.address);
        let incomingSettings = query.state;
        if (incomingSettings) {
            this.updateSettings(session, incomingSettings);
        }
    }

    // Handle other invokes
    private async onInvoke(event: builder.IEvent, cb: (err: Error, result: msteams.IComposeExtensionResponse, statusCode?: number) => void): Promise<void> {
        let session = await this.loadSessionAsync(event.address);
        if (session) {
            let invokeEvent = event as msteams.IInvokeEvent;
            let eventName = invokeEvent.name;
            // let eventValue = invokeEvent.value;

            switch (eventName) {
                default:
                    let unrecognizedEvent = `Unrecognized event name: ${eventName}`;
                    winston.error(unrecognizedEvent);
                    cb(new Error(unrecognizedEvent), null, 500);
                    break;
            }
        } else {
            cb(new Error(), null, 500);
        }
    }

    private getConfigurationResponse(session: builder.Session, translationLanguages?: string[]): msteams.IComposeExtensionResponse {
        translationLanguages = translationLanguages || this.getTranslationLanguages(session);
        let baseUri = config.get("app.baseUri");
        let languages = translationLanguages.join(",");
        let response = msteams.ComposeExtensionResponse.config().actions([
            builder.CardAction.openUrl(session, `${baseUri}/html/compose-config.html?languages=${languages}`, Strings.configure_text),
        ]);
        return response.toResponse();
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

    private updateSettings(session: builder.Session, state: string): void {
        // State is a comma-separated list of languages
        state = state || "";

        let supportedLangs = this.translator.getSupportedLanguages();
        let langs = state.split(",")
            .filter(lang => supportedLangs.find(i => i === lang));
        if (langs.length === 0) {
            langs = this.translator.getDefaultLanguages();
        }

        session.userData.languages = langs;
        session.sendBatch();
    }

    private getTranslationLanguages(session: builder.Session): string[] {
        return session.userData.languages || this.translator.getDefaultLanguages();
    }
}
