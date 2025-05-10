const { Client } = require('@microsoft/microsoft-graph-client');
const vscode = require('vscode');
const axios = require('axios');

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

        // First, check if the token is valid by getting the user profile
        await client.api('/me').get();

        // Get chats with expanded members to show user names
        const response = await client
            .api('/me/chats')
            .expand('members')
            .top(50) // Adjust based on expected number of chats
            .get();

        return response.value;
    } catch (error) {
        console.error('Error getting Teams chats:', error);
        
        // If we get a 401 or 403, the token is invalid or doesn't have the right permissions
        if (error.statusCode === 401 || error.statusCode === 403) {
            throw new Error('Not authorized to access Teams chats. Please check permissions or sign out and try again.');
        }
        
        // Try a direct approach using axios as a fallback
        try {
            const response = await axios.get('https://graph.microsoft.com/v1.0/me/chats?$expand=members', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (response.data && response.data.value) {
                return response.data.value;
            }
        } catch (axiosError) {
            console.error('Axios fallback also failed:', axiosError);
        }
        
        // If we reach here, both approaches failed
        throw new Error('Failed to get Teams chats. Please ensure you have Microsoft Teams installed and are logged in.');
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
        
        // Try a direct approach using axios as a fallback
        try {
            const message = {
                body: {
                    contentType: 'text',
                    content: content
                }
            };
            
            const response = await axios.post(`https://graph.microsoft.com/v1.0/chats/${chatId}/messages`, message, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (response.data) {
                return response.data;
            }
        } catch (axiosError) {
            console.error('Axios fallback also failed:', axiosError);
        }
        
        // If we reach here, both approaches failed
        throw new Error('Failed to send message to Teams. Please try the fallback method.');
    }
}

module.exports = {
    getChats,
    sendMessage
};