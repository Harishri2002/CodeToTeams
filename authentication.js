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
        cacheLocation: 'localStorage', // This enables caching the token
    }
};

const cca = new ConfidentialClientApplication(msalConfig);

// Store token cache path
let tokenCachePath;

/**
 * Initialize the authentication module
 * @param {vscode.ExtensionContext} context - Extension context for storage
 */
function initialize(context) {
    tokenCachePath = path.join(context.globalStoragePath, 'msal-token-cache.json');
    
    // Create directory if it doesn't exist
    if (!fs.existsSync(context.globalStoragePath)) {
        fs.mkdirSync(context.globalStoragePath, { recursive: true });
    }
    
    // Try to load existing token cache
    try {
        if (fs.existsSync(tokenCachePath)) {
            const cacheData = fs.readFileSync(tokenCachePath, 'utf8');
            cca.getTokenCache().deserialize(cacheData);
            console.log('Loaded existing token cache');
        }
    } catch (error) {
        console.error('Error loading token cache:', error);
        // Continue without cache if there's an error
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
        // Check if we have accounts in the cache
        const accounts = await cca.getTokenCache().getAllAccounts();
        
        if (accounts && accounts.length > 0) {
            // Use the first account
            const silentRequest = {
                account: accounts[0],
                scopes: ['Chat.ReadWrite', 'User.Read'],
            };
            
            const tokenResponse = await cca.acquireTokenSilent(silentRequest);
            if (tokenResponse && tokenResponse.accessToken) {
                console.log('Got token silently from cache');
                return tokenResponse.accessToken;
            }
        }
        return null;
    } catch (error) {
        console.log('Silent token acquisition failed:', error);
        return null;
    }
}

/**
 * Get an access token for Microsoft Graph API using the authorization code flow
 * @returns {Promise<string|null>} Access token or null if authentication fails
 */
async function getAccessToken() {
    try {
        // First try to get a token silently
        const silentToken = await getSilentToken();
        if (silentToken) {
            return silentToken;
        }
        
        // If silent token acquisition fails, proceed with interactive login
        console.log('No cached token found, proceeding with interactive login');
        
        // Step 1: Generate auth URL
        const authCodeUrlParameters = {
            scopes: ['Chat.ReadWrite', 'User.Read', 'ChatMessage.Send'],
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
                        // Prettier success page with redirect button
                        const successHtml = `
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <title>Authentication Successful</title>
                            <style>
                                body {
                                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                                    background-color: #f9f9f9;
                                    color: #333;
                                    text-align: center;
                                    padding: 50px;
                                    max-width: 600px;
                                    margin: 0 auto;
                                }
                                .container {
                                    background-color: white;
                                    border-radius: 8px;
                                    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
                                    padding: 30px;
                                }
                                h1 {
                                    color: #107C10;
                                    margin-bottom: 20px;
                                }
                                p {
                                    margin-bottom: 30px;
                                    line-height: 1.5;
                                }
                                .button {
                                    background-color: #0078d4;
                                    color: white;
                                    border: none;
                                    padding: 12px 24px;
                                    border-radius: 4px;
                                    cursor: pointer;
                                    font-size: 16px;
                                    text-decoration: none;
                                    display: inline-block;
                                    transition: background-color 0.3s;
                                }
                                .button:hover {
                                    background-color: #106ebe;
                                }
                                .icon {
                                    font-size: 48px;
                                    color: #107C10;
                                    margin-bottom: 20px;
                                }
                            </style>
                        </head>
                        <body>
                            <div class="container">
                                <div class="icon">✓</div>
                                <h1>Authentication Successful!</h1>
                                <p>You've successfully authenticated with Microsoft Teams. You can now share code snippets directly to Teams from VS Code.</p>
                                <p>You can close this window and return to VS Code.</p>
                                <a href="vscode://Harishri.CodeToTeams" class="button">Return to VS Code</a>
                            </div>
                            <script>
                                // Auto-redirect after 3 seconds
                                setTimeout(() => {
                                    window.location.href = "vscode://Harishri.CodeToTeams";
                                }, 3000);
                            </script>
                        </body>
                        </html>`;
                        
                        res.writeHead(200, { 'Content-Type': 'text/html' });
                        res.end(successHtml);
                        resolve(code);
                    } else {
                        // Error page
                        const errorHtml = `
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <title>Authentication Failed</title>
                            <style>
                                body {
                                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                                    background-color: #f9f9f9;
                                    color: #333;
                                    text-align: center;
                                    padding: 50px;
                                    max-width: 600px;
                                    margin: 0 auto;
                                }
                                .container {
                                    background-color: white;
                                    border-radius: 8px;
                                    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
                                    padding: 30px;
                                }
                                h1 {
                                    color: #d83b01;
                                    margin-bottom: 20px;
                                }
                                p {
                                    margin-bottom: 30px;
                                    line-height: 1.5;
                                }
                                .button {
                                    background-color: #0078d4;
                                    color: white;
                                    border: none;
                                    padding: 12px 24px;
                                    border-radius: 4px;
                                    cursor: pointer;
                                    font-size: 16px;
                                    text-decoration: none;
                                    display: inline-block;
                                    transition: background-color 0.3s;
                                }
                                .button:hover {
                                    background-color: #106ebe;
                                }
                                .icon {
                                    font-size: 48px;
                                    color: #d83b01;
                                    margin-bottom: 20px;
                                }
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
                                // Auto-redirect after 3 seconds
                                setTimeout(() => {
                                    window.location.href = "vscode://Harishri.CodeToTeams";
                                }, 3000);
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

        // Step 4: Exchange the auth code for an access token
        const tokenResponse = await cca.acquireTokenByCode({
            code: authCode,
            scopes: ['Chat.ReadWrite', 'User.Read', 'ChatMessage.Send'],
            redirectUri: msalConfig.auth.redirectUri,
        });

        if (tokenResponse && tokenResponse.accessToken) {
            console.log('Access token retrieved successfully');
            
            // Save token cache
            await saveTokenCache();
            
            return tokenResponse.accessToken;
        }

        throw new Error('Failed to retrieve access token');
    } catch (error) {
        console.error('Authentication error:', error);
        vscode.window.showErrorMessage(`Authentication failed: ${error.message}`);
        return null;
    }
}

/**
 * Sign out the user and clear the token cache
 */
async function signOut() {
    try {
        // Get all accounts and remove them
        const accounts = await cca.getTokenCache().getAllAccounts();
        
        for (const account of accounts) {
            await cca.getTokenCache().removeAccount(account);
        }
        
        // Clear the cache file
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