import * as request from "request";
import * as xml2js from "xml2js";

// =========================================================
// Translator API
// =========================================================

// Access tokens last 10 minutes, but refresh every 9 minutes to be safe
const accessTokenLifetimeMs = 9 * 60 * 1000;

export class TranslatorApi {

    private accessToken: string;
    private accessTokenExpiryTime: number;

    constructor(
        private accessKey: string,
    )
    {
    }

    public async translateText(text: string, to: string): Promise<string> {
        // Default to valid parameters
        text = text || "";
        to = to || "en";

        let url = `https://api.microsofttranslator.com/v2/http.svc/Translate?text=${encodeURIComponent(text)}&to=${encodeURIComponent(to)}`;
        let authHeader = await this.getAuthorizationHeader();
        let options: request.Options = {
            url: url,
            headers: {
                "Authorization": authHeader,
            },
        };

        return new Promise<string>((resolve, reject) => {
            request.get(options, (error, response, body) => {
                if (error) {
                    reject(error);
                } else if (response.statusCode !== 200) {
                    reject(new Error(response.statusMessage));
                } else {
                    xml2js.parseString(body as string, (parseError, result) => {
                        if (parseError) {
                            reject(parseError);
                        } else {
                            resolve(result.string._);
                        }
                    });
                }
            });
        });
    }

    private async getAuthorizationHeader(): Promise<string> {
        if (!this.accessToken || this.isAccessTokenExpired()) {
            await this.refreshAccessToken();
        }

        return "Bearer " + this.accessToken;
    }

    private isAccessTokenExpired(): boolean {
        return new Date().valueOf() > this.accessTokenExpiryTime;
    }

    private async refreshAccessToken(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            this.accessToken = null;
            this.accessTokenExpiryTime = 0;

            let options: request.Options = {
                url: "https://api.cognitive.microsoft.com/sts/v1.0/issueToken",
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey,
                },
                body: "",
            };

            request.post(options, (error, response, body) => {
                if (error) {
                    reject(error);
                } else if (response.statusCode !== 200) {
                    reject(new Error(response.statusMessage));
                } else {
                    this.accessToken = body as string;
                    this.accessTokenExpiryTime = new Date().valueOf() + accessTokenLifetimeMs;
                    resolve();
                }
            });
        });
    }

}
