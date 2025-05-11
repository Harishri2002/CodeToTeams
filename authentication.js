const { ConfidentialClientApplication } = require('@azure/msal-node');
const vscode = require('vscode');
const http = require('http');
const url = require('url');
const path = require('path');
const fs = require('fs');

// Configuration for Microsoft Authentication Library
const msalConfig = {
    auth: {
  
    },
    cache: {
        cacheLocation: 'localStorage',
    }
};

const cca = new ConfidentialClientApplication(msalConfig);
let tokenCachePath;

// Required scopes for Microsoft Graph
const requiredScopes = ['Chat.ReadWrite', 'User.Read', 'ChatMessage.Send'];

/**
 * Initialize the authentication module
 * @param {vscode.ExtensionContext} context - Extension context for storage
 */
function initialize(context) {
    tokenCachePath = path.join(context.globalStoragePath, 'msal-token-cache.json');
    if (!fs.existsSync(context.globalStoragePath)) {
        fs.mkdirSync(context.globalStoragePath, { recursive: true });
    }
    try {
        if (fs.existsSync(tokenCachePath)) {
            const cacheData = fs.readFileSync(tokenCachePath, 'utf8');
            if (cacheData && JSON.parse(cacheData)) {
                cca.getTokenCache().deserialize(cacheData);
                console.log('Loaded existing token cache');
            } else {
                console.log('Token cache is empty or invalid, starting fresh');
            }
        }
    } catch (error) {
        console.error('Error loading token cache, clearing cache:', error);
        if (fs.existsSync(tokenCachePath)) {
            fs.unlinkSync(tokenCachePath);
        }
    }
}

/**
 * Save the token cache to disk
 */
async function saveTokenCache() {
    try {
        const cacheData = cca.getTokenCache().serialize();
        fs.writeFileSync(tokenCachePath, cacheData);
        console.log('Token cache saved');
    } catch (error) {
        console.error('Error saving token cache:', error);
    }
}

/**
 * Try to get token silently from cache
 * @returns {Promise<string|null>} Access token or null if not available
 */
async function getSilentToken() {
    try {
        const accounts = await cca.getTokenCache().getAllAccounts();
        if (accounts && accounts.length > 0) {
            const silentRequest = {
                account: accounts[0],
                scopes: requiredScopes,
            };
            console.log('Requesting silent token with scopes:', silentRequest.scopes);
            const tokenResponse = await cca.acquireTokenSilent(silentRequest);
            if (tokenResponse && tokenResponse.accessToken) {
                const tokenScopes = tokenResponse.scopes || [];
                const missingScopes = requiredScopes.filter(scope => !tokenScopes.includes(scope));
                if (missingScopes.length === 0) {
                    console.log('Got token silently from cache with scopes:', tokenScopes);
                    return tokenResponse.accessToken;
                } else {
                    console.log('Cached token missing required scopes:', missingScopes);
                    return null;
                }
            }
        }
        return null;
    } catch (error) {
        console.log('Silent token acquisition failed:', error);
        if (error.errorCode === 'invalid_grant' || error.errorCode === 'no_tokens_found') {
            console.log('Clearing invalid token cache');
            if (fs.existsSync(tokenCachePath)) {
                fs.unlinkSync(tokenCachePath);
            }
        }
        return null;
    }
}

/**
 * Get an access token for Microsoft Graph API
 * @returns {Promise<string|null>} Access token or null if authentication fails
 */
async function getAccessToken() {
    try {
        const silentToken = await getSilentToken();
        if (silentToken) {
            return silentToken;
        }
        console.log('No valid cached token found, proceeding with interactive login');
        const authCodeUrlParameters = {
            scopes: requiredScopes,
            redirectUri: msalConfig.auth.redirectUri,
            prompt: 'consent', // Ensure all scopes are granted
        };
        console.log('Requesting authorization code with scopes:', authCodeUrlParameters.scopes);
        const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);
        console.log('Opening authentication URL:', authUrl);
        await vscode.env.openExternal(vscode.Uri.parse(authUrl));
        const authCode = await new Promise((resolve, reject) => {
            const server = http.createServer((req, res) => {
                const requestUrl = url.parse(req.url, true);
                if (requestUrl.pathname === '/auth/callback') {
                    const code = requestUrl.query.code;
                    if (code) {
                        const successHtml = `
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <title>Authentication Successful</title>
                            <style>
                                body { font-family: 'Segoe UI', sans-serif; background-color: #f9f9f9; color: #333; text-align: center; padding: 50px; max-width: 600px; margin: 0 auto; }
                                .container { background-color: white; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); padding: 30px; }
                                h1 { color: #107C10; margin-bottom: 20px; }
                                p { margin-bottom: 30px; line-height: 1.5; }
                                .button { background-color: #0078d4; color: white; border: none; padding: 12px 24px; border-radius: 4px; cursor: pointer; font-size: 16px; text-decoration: none; display: inline-block; }
                                .button:hover { background-color: #106ebe; }
                                .icon { font-size: 48px; color: #107C10; margin-bottom: 20px; }
                            </style>
                        </head>
                        <body>
                            <div class="container">
                                <div class="icon">✓</div>
                                <h1>Authentication Successful!</h1>
                                <p>You've successfully authenticated with Microsoft Teams.</p>
                                <a href="vscode://Harishri.CodeToTeams" class="button">Return to VS Code</a>
                            </div>
                            <script>
                                setTimeout(() => { window.location.href = "vscode://Harishri.CodeToTeams"; }, 3000);
                            </script>
                        </body>
                        </html>`;
                        res.writeHead(200, { 'Content-Type': 'text/html' });
                        res.end(successHtml);
                        resolve(code);
                    } else {
                        const errorHtml = `
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <title>Authentication Failed</title>
                            <style>
                                body { font-family: 'Segoe UI', sans-serif; background-color: #f9f9f9; color: #333; text-align: center; padding: 50px; max-width: 600px; margin: 0 auto; }
                                .container { background-color: white; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); padding: 30px; }
                                h1 { color: #d83b01; margin-bottom: 20px; }
                                p { margin-bottom: 30px; line-height: 1.5; }
                                .button { background-color: #0078d4; color: white; border: none; padding: 12px 24px; border-radius: 4px; cursor: pointer; font-size: 16px; text-decoration: none; display: inline-block; }
                                .button:hover { background-color: #106ebe; }
                                .icon { font-size: 48px; color: #d83b01; margin-bottom: 20px; }
                            </style>
                        </head>
                        <body>
                            <div class="container">
                                <div class="icon">✗</div>
                                <h1>Authentication Failed</h1>
                                <p>No authorization code received. Please try again.</p>
                                <a href="vscode://Harishri.CodeToTeams" class="button">Return to VS Code</a>
                            </div>
                            <script>
                                setTimeout(() => { window.location.href = "vscode://Harishri.CodeToTeams"; }, 3000);
                            </script>
                        </body>
                        </html>`;
                        res.writeHead(400, { 'Content-Type': 'text/html' });
                        res.end(errorHtml);
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
        console.log('Acquiring token with authorization code, requesting scopes:', requiredScopes);
        const tokenResponse = await cca.acquireTokenByCode({
            code: authCode,
            scopes: requiredScopes,
            redirectUri: msalConfig.auth.redirectUri,
        });
        if (tokenResponse && tokenResponse.accessToken) {
            console.log('Access token retrieved successfully with scopes:', tokenResponse.scopes);
            console.log('Full token response for debugging:', {
                scopes: tokenResponse.scopes,
                tenantId: tokenResponse.tenantId,
                expiresOn: tokenResponse.expiresOn,
                error: tokenResponse.error,
                errorDescription: tokenResponse.errorDescription
            });
            await saveTokenCache();
            return tokenResponse.accessToken;
        }
        throw new Error('Failed to retrieve access token');
    } catch (error) {
        console.error('Authentication error:', error);
        let errorMessage = `Authentication failed: ${error.message}`;
        if (error.errorCode === 'access_denied') {
            errorMessage = 'Authentication failed: Access denied. Please check your Microsoft Teams permissions or ensure admin consent is granted for Chat.ReadWrite and ChatMessage.Send.';
        } else if (error.errorCode === 'invalid_grant') {
            errorMessage = 'Authentication failed: Invalid token. Please sign out and try again.';
        } else if (error.errorCode === 'consent_required') {
            errorMessage = 'Authentication failed: Consent required for Chat.ReadWrite and ChatMessage.Send. Please sign out, re-authenticate, and grant consent.';
        }
        vscode.window.showErrorMessage(errorMessage);
        return null;
    }
}

/**
 * Sign out the user and clear the token cache
 */
async function signOut() {
    try {
        const accounts = await cca.getTokenCache().getAllAccounts();
        for (const account of accounts) {
            await cca.getTokenCache().removeAccount(account);
        }
        if (fs.existsSync(tokenCachePath)) {
            fs.unlinkSync(tokenCachePath);
        }
        console.log('Signed out successfully');
        vscode.window.showInformationMessage('Signed out from Microsoft Teams');
    } catch (error) {
        console.error('Error signing out:', error);
        vscode.window.showErrorMessage(`Failed to sign out: ${error.message}`);
    }
}

module.exports = {
    initialize,
    getAccessToken,
    signOut
};