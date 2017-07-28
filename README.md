# Translator
This contains the source for the Microsoft Translator compose extension for Microsoft Teams.

### Compile
To compile the Typescript files, run `gulp build`.
To package up the app for Azure deployment, run `gulp package`.

### Debugging
Visual Studio Code is recommended for running/debugging the code.

#### Prerequisites
1. Set up a Microsoft Translator account, following the instructions at [Getting started using the Translator API](https://www.microsoft.com/en-us/translator/getstarted.aspx).
2. Download ngrok from https://ngrok.com/. Run the following command to setup a tunnel to localhost:3978
 `ngrok http 3978`
 Note the ngrok address, which looks something like `https://013e0d3f.ngrok.io`.
3. Register a bot at https://dev.botframework.com. Note the app id and password for your bot.
4. Set the messaging endpoint for the bot to `https://[ngrok_https_url]/api/messages`.

#### Launch configuration
Add the following configuration to `launch.json` (or define the environment variables).
```
    {
        "type": "node",
        "request": "launch",
        "name": "Launch Program",
        "program": "${workspaceRoot}\\build\\src\\app.js",
        "cwd": "${workspaceRoot}\\build\\src",
        "sourceMaps": true,
        "outFiles": [ "${workspaceRoot}/build/**/*.js" ],
        "env": {
            "APP_BASE_URI": "[ngrok_https_url]",
            "MICROSOFT_APP_ID": "[your_bot_app_id]",
            "MICROSOFT_APP_PASSWORD": "[your_bot_secret]",
            "TRANSLATOR_ACCESS_KEY": "[your translator access key]"
        }
    },
```

#### Fiddler
To use Fiddler, add the following lines to the `env` section:
```
    "http_proxy": "http://localhost:8888",
    "no_proxy": "login.microsoftonline.com",
    "NODE_TLS_REJECT_UNAUTHORIZED": "0"
```
Change `http_proxy` to the Fiddler endpoint.

### Bot state
Per-user and per-conversation state goes to the Bot Framework state store (https://docs.botframework.com/en-us/core-concepts/userdata/).

### Configuration 
 - Default configuration is in `config\default.json`
 - Production overrides are in `config\production.json`
 - Environment variables overrides are in `config\custom-environment-variables.json`.
 - Specify local overrides in `config\local.json`. **Do not check in this file.**
