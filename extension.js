const vscode = require('vscode');
const auth = require('./authentication');
const teamsService = require('./teamsService');

/**
 * Activates the extension
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
    console.log('Activating "Share Code to Teams" extension');

    // Initialize authentication module with context
    auth.initialize(context);

    // Register the command to share code to Teams
    let shareCommand = vscode.commands.registerCommand('extension.shareToTeams', async function () {
        try {
            // Check if text is selected
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showErrorMessage('No active editor found');
                return;
            }

            const selection = editor.selection;
            if (selection.isEmpty) {
                vscode.window.showErrorMessage('No text selected');
                return;
            }

            // Get the selected text
            const selectedText = editor.document.getText(selection);
            if (!selectedText) {
                vscode.window.showErrorMessage('Selected text is empty');
                return;
            }

            // Format the code snippet
            const languageId = editor.document.languageId;
            const includeLanguage = vscode.workspace.getConfiguration('shareToTeams').get('includeLanguage');
            const formattedText = formatCodeSnippet(selectedText, includeLanguage ? languageId : null);

            // Show progress indicator
            await vscode.window.withProgress({
                location: vscode.ProgressLocation.Notification,
                title: "Sharing code to Teams",
                cancellable: false
            }, async (progress) => {
                progress.report({ message: "Authenticating..." });
                
                try {
                    // First try direct API integration
                    const directShareResult = await shareViaGraphApi(formattedText, context, progress);
                    
                    if (directShareResult) {
                        // Success - API share worked
                        return;
                    }
                    
                    // If direct API fails, use deep linking as fallback
                    progress.report({ message: "Using Teams deep link..." });
                    shareViaDeepLink(formattedText);
                    
                } catch (error) {
                    console.error('Error sharing:', error);
                    
                    // Show error and offer direct Teams link as fallback
                    const fallback = await vscode.window.showErrorMessage(
                        `Error: ${error.message}. Use Teams deep link instead?`,
                        'Yes', 'No'
                    );
                    
                    if (fallback === 'Yes') {
                        shareViaDeepLink(formattedText);
                    }
                }
            });
        } catch (error) {
            console.error('Error sharing code to Teams:', error);
            vscode.window.showErrorMessage(`Error sharing code to Teams: ${error.message}`);
        }
    });

    // Register a command to sign out
    let signOutCommand = vscode.commands.registerCommand('extension.teamsSignOut', async function() {
        await auth.signOut();
    });

    context.subscriptions.push(shareCommand, signOutCommand);
}

/**
 * Format the selected code as a code block
 * @param {string} text - Selected text
 * @param {string|null} languageId - Language identifier
 * @returns {string} - Formatted code block
 */
function formatCodeSnippet(text, languageId) {
    const language = languageId || '';
    return `\`\`\`${language}\n${text}\n\`\`\``;
}

/**
 * Share code via Microsoft Graph API
 * @param {string} formattedText - Formatted code block
 * @param {vscode.ExtensionContext} context - Extension context
 * @param {vscode.Progress} progress - Progress reporter
 * @returns {Promise<boolean>} - True if successful, false otherwise
 */
async function shareViaGraphApi(formattedText, context, progress) {
    // Get access token
    progress.report({ message: "Getting access token..." });
    const accessToken = await auth.getAccessToken();
    
    if (!accessToken) {
        throw new Error('Authentication failed');
    }

    // Get list of chats
    progress.report({ message: "Fetching Teams chats..." });
    const chats = await teamsService.getChats(accessToken);
    
    if (!chats || chats.length === 0) {
        throw new Error('No Teams chats found');
    }

    // Prepare chat items for selection
    const chatItems = chats.map(chat => ({
        label: getChatLabel(chat),
        id: chat.id,
        description: getChatDescription(chat)
    }));

    // Show quick pick to select a chat
    const selectedChat = await vscode.window.showQuickPick(chatItems, {
        placeHolder: 'Select a Teams chat to share the code',
        matchOnDescription: true
    });

    if (!selectedChat) {
        // User cancelled the selection
        return false;
    }

    // Send the message
    progress.report({ message: `Sending message to ${selectedChat.label}...` });
    await teamsService.sendMessage(accessToken, selectedChat.id, formattedText);
    
    // Show success message
    vscode.window.showInformationMessage(`Code shared to Teams chat "${selectedChat.label}" successfully!`);
    return true;
}

/**
 * Get a display label for a chat
 * @param {Object} chat - Chat object from Graph API
 * @returns {string} - Display label
 */
function getChatLabel(chat) {
    if (chat.topic) {
        return chat.topic; // Group chat with topic
    }
    
    // Try to find the other person in a one-on-one chat
    if (chat.members && chat.members.length > 0) {
        const otherMembers = chat.members.filter(m => !m.displayName.includes('(You)'));
        if (otherMembers.length > 0) {
            return otherMembers.map(m => m.displayName).join(', ');
        }
    }
    
    return 'Chat';
}

/**
 * Get a description for a chat
 * @param {Object} chat - Chat object from Graph API
 * @returns {string} - Chat description
 */
function getChatDescription(chat) {
    if (chat.topic) {
        // For group chats, show the number of members
        return chat.members ? `${chat.members.length} members` : 'Group chat';
    }
    
    return 'Direct message';
}

/**
 * Share code via Teams deep link
 * @param {string} formattedText - Formatted code block
 */
function shareViaDeepLink(formattedText) {
    const encodedMessage = encodeURIComponent(formattedText);
    const deepLink = `msteams://teams.microsoft.com/l/chat/0/0?message=${encodedMessage}`;
    
    vscode.env.openExternal(vscode.Uri.parse(deepLink));
    vscode.window.showInformationMessage('Teams opened. Please select recipients to share your code.');
}

function deactivate() {}

module.exports = {
    activate,
    deactivate,
    formatCodeSnippet // Export for testing
};