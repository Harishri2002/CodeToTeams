const { ConfidentialClientApplication } = require('@azure/msal-node');
const vscode = require('vscode');
const http = require('http');
const url = require('url');

const msalConfig = {
    auth: {
    }
};

const cca = new ConfidentialClientApplication(msalConfig);

/**
 * Get an access token for Microsoft Graph API using the authorization code flow
 * @returns {Promise<string|null>} Access token or null if authentication fails
 */
async function getAccessToken() {
    try {
        // Step 1: Generate auth URL
        const authCodeUrlParameters = {
            scopes: ['Chat.ReadWrite', 'User.Read'],
            redirectUri: msalConfig.auth.redirectUri,
        };
        const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);

        // Step 2: Open the auth URL in the user's browser
        await vscode.env.openExternal(vscode.Uri.parse(authUrl));

        // Step 3: Set up a local server to capture the authorization code
        const authCode = await new Promise((resolve, reject) => {
            const server = http.createServer((req, res) => {
                const requestUrl = url.parse(req.url, true);
                if (requestUrl.pathname === '/auth/callback') {
                    const code = requestUrl.query.code;
                    if (code) {
                        res.writeHead(200, { 'Content-Type': 'text/plain' });
                        res.end('Authentication successful! You can close this window.');
                        resolve(code);
                    } else {
                        res.writeHead(400, { 'Content-Type': 'text/plain' });
                        res.end('Authentication failed: No code received.');
                        reject(new Error('No authorization code received'));
                    }
                    server.close();
                }
            });

            server.listen(3000, () => {
                console.log('Local server listening on port 3000 for auth callback...');
            });

            server.on('error', (err) => {
                reject(new Error(`Server error: ${err.message}`));
            });
        });

        // Step 4: Exchange the auth code for an access token
        const tokenResponse = await cca.acquireTokenByCode({
            code: authCode,
            scopes: ['Chat.ReadWrite', 'User.Read'],
            redirectUri: msalConfig.auth.redirectUri,
        });

        if (tokenResponse && tokenResponse.accessToken) {
            console.log('Access token retrieved successfully');
            return tokenResponse.accessToken;
        }

        throw new Error('Failed to retrieve access token');
    } catch (error) {
        console.error('Authentication error:', error);
        vscode.window.showErrorMessage(`Authentication failed: ${error.message}`);
        return null;
    }
}

module.exports = {
    getAccessToken
};