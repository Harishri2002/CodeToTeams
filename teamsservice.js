const { Client } = require('@microsoft/microsoft-graph-client');
const vscode = require('vscode');

/**
 * Get the user's Teams chats
 * @param {string} accessToken - Microsoft Graph API access token
 * @returns {Promise<Array>} - List of chats
 */
async function getChats(accessToken) {
    try {
        const client = Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        // Get chats with expanded members to show user names
        const response = await client
            .api('/me/chats')
            .expand('members')
            .top(50) // Adjust based on expected number of chats
            .get();

        return response.value;
    } catch (error) {
        console.error('Error getting Teams chats:', error);
        const errorMessage = error.message || 'Unknown error';
        const errorCode = error.statusCode || '';
        
        // Handle common error cases
        if (errorCode === 401 || errorCode === 403) {
            throw new Error('Not authorized to access Teams chats. Please check permissions.');
        } else if (errorCode === 404) {
            throw new Error('Teams chat API not found. Are you using Microsoft Teams?');
        } else {
            throw new Error(`Failed to get Teams chats: ${errorMessage}`);
        }
    }
}

/**
 * Send a message to a Teams chat
 * @param {string} accessToken - Microsoft Graph API access token
 * @param {string} chatId - ID of the chat to send the message to
 * @param {string} content - Message content
 * @returns {Promise<Object>} - API response
 */
async function sendMessage(accessToken, chatId, content) {
    try {
        const client = Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        // Create message payload
        const message = {
            body: {
                contentType: 'text',
                content: content
            }
        };

        // Send the message
        const response = await client
            .api(`/chats/${chatId}/messages`)
            .post(message);

        return response;
    } catch (error) {
        console.error('Error sending message to Teams:', error);
        const errorMessage = error.message || 'Unknown error';
        const errorCode = error.statusCode || '';
        
        // Handle common error cases
        if (errorCode === 401 || errorCode === 403) {
            throw new Error('Not authorized to send messages. Please check permissions.');
        } else if (errorCode === 404) {
            throw new Error('Chat not found. The chat ID may be invalid.');
        } else {
            throw new Error(`Failed to send message: ${errorMessage}`);
        }
    }
}

module.exports = {
    getChats,
    sendMessage
};