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

        // Verify token by getting user profile
        console.log('Validating token by fetching user profile...');
        const userProfile = await client.api('/me').get();
        console.log('User profile fetched successfully:', userProfile.id);

        // Get chats with expanded members
        console.log('Fetching Teams chats with expanded members...');
        const response = await client
            .api('/me/chats')
            .expand('members')
            .top(50)
            .get();

        console.log('Teams chats fetched successfully:', response.value.length, 'chats found');
        return response.value;
    } catch (error) {
        console.error('Error getting Teams chats:', error);
        if (error.statusCode === 401 || error.statusCode === 403) {
            console.error('Authorization error details:', error.body);
            throw new Error('Not authorized to access Teams chats. Please sign out and try again or check permissions in Azure AD.');
        }
        try {
            console.log('Falling back to axios for fetching chats...');
            const response = await axios.get('https://graph.microsoft.com/v1.0/me/chats?$expand=members', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            if (response.data && response.data.value) {
                console.log('Teams chats fetched via axios fallback:', response.data.value.length, 'chats found');
                return response.data.value;
            }
        } catch (axiosError) {
            console.error('Axios fallback also failed:', axiosError);
        }
        throw new Error('Failed to get Teams chats. Ensure Microsoft Teams is installed and you are logged in.');
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

        const message = {
            body: {
                contentType: 'text',
                content: content
            }
        };

        console.log('Sending message to chat ID:', chatId);
        const response = await client
            .api(`/chats/${chatId}/messages`)
            .post(message);

        console.log('Message sent successfully:', response.id);
        return response;
    } catch (error) {
        console.error('Error sending message to Teams:', error);
        try {
            const message = {
                body: {
                    contentType: 'text',
                    content: content
                }
            };
            console.log('Falling back to axios for sending message...');
            const response = await axios.post(`https://graph.microsoft.com/v1.0/chats/${chatId}/messages`, message, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            if (response.data) {
                console.log('Message sent via axios fallback:', response.data.id);
                return response.data;
            }
        } catch (axiosError) {
            console.error('Axios fallback also failed:', axiosError);
        }
        throw new Error('Failed to send message to Teams. Try the fallback method.');
    }
}

module.exports = {
    getChats,
    sendMessage
};