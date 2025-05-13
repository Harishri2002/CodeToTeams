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
// Changed from User.ReadBasic.All to Contacts.Read
// Also includes offline_access to get refresh tokens
const requiredScopes = ['User.Read', 'Contacts.Read', 'People.Read', 'offline_access'];

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
                console.log('Got token silently from cache with scopes:', tokenResponse.scopes);
                return tokenResponse.accessToken;
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
 * Get access token supporting both organization and personal accounts
 * @returns {Promise<string|null>} Access token or null if authentication fails
 */
async function getAccessToken() {
    try {
        // First try with existing token
        const silentToken = await getSilentToken();
        if (silentToken) {
            return silentToken;
        }
        
        console.log('No valid cached token found, proceeding with interactive login');
        
        // Use common endpoint to support both personal and work accounts
        vscode.window.showInformationMessage('Signing in with your Microsoft account...');
        const token = await loginWithAuthority('https://login.microsoftonline.com/common');
        if (token) return token;
        
        throw new Error('Authentication failed. Please try again or check your account permissions.');
        
    } catch (error) {
        console.error('Authentication error:', error);
        let errorMessage = `Authentication failed: ${error.message}`;
        
        if (error.errorCode === 'access_denied') {
            errorMessage = 'Authentication failed: Access denied. Please check your Microsoft account permissions.';
        } else if (error.errorCode === 'invalid_grant') {
            errorMessage = 'Authentication failed: Invalid token. Please sign out and try again.';
        } else if (error.errorCode === 'consent_required') {
            errorMessage = 'Authentication failed: Consent required. Please sign out, re-authenticate, and grant consent.';
        }
        
        vscode.window.showErrorMessage(errorMessage);
        return null;
    }
}

/**
 * Login with a specific authority
 * @param {string} authority - The authority URL to use
 * @returns {Promise<string|null>} - Access token if successful
 */
async function loginWithAuthority(authority) {
    // Create a new client with the specified authority
    const tempConfig = {
        ...msalConfig,
        auth: {
            ...msalConfig.auth,
            authority: authority
        }
    };
    
    const tempCca = new ConfidentialClientApplication(tempConfig);
    
    const authCodeUrlParameters = {
        scopes: requiredScopes,
        redirectUri: msalConfig.auth.redirectUri,
        prompt: 'select_account', // Allow user to choose account
    };
    
    console.log(`Requesting authorization code with ${authority} and scopes:`, authCodeUrlParameters.scopes);
    const authUrl = await tempCca.getAuthCodeUrl(authCodeUrlParameters);
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
                            <p>You've successfully authenticated with Microsoft.</p>
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
    
    console.log(`Acquiring token with authorization code using ${authority}, requesting scopes:`, requiredScopes);
    const tokenResponse = await tempCca.acquireTokenByCode({
        code: authCode,
        scopes: requiredScopes,
        redirectUri: msalConfig.auth.redirectUri,
    });
    
    if (tokenResponse && tokenResponse.accessToken) {
        console.log('Access token retrieved successfully with scopes:', tokenResponse.scopes);
        
        // Update the global cache with this token
        const serializedCache = tempCca.getTokenCache().serialize();
        cca.getTokenCache().deserialize(serializedCache);
        await saveTokenCache();
        
        return tokenResponse.accessToken;
    }
    
    throw new Error('Failed to retrieve access token');
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
        vscode.window.showInformationMessage('Signed out from Microsoft account');
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