const vscode = require('vscode');

/**
 * Get an access token for Microsoft Graph API
 * @returns {Promise<string|null>} Access token or null if authentication fails
 */
async function getAccessToken() {
    try {
        // Use VS Code's built-in authentication provider
        const session = await vscode.authentication.getSession(
            'microsoft',
            ['Chat.ReadWrite'],
            { createIfNone: true }
        );
        
        if (session && session.accessToken) {
            return session.accessToken;
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