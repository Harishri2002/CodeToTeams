const vscode = require('vscode');
const auth = require('./authentication');
const teamsService = require('./teamsService');

/**
 * Activates the extension
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
    console.log('Activating "Share Code to Teams" extension');
    auth.initialize(context);

    let shareCommand = vscode.commands.registerCommand('extension.shareToTeams', async function () {
        try {
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

            const selectedText = editor.document.getText(selection);
            if (!selectedText) {
                vscode.window.showErrorMessage('Selected text is empty');
                return;
            }

            const languageId = editor.document.languageId;
            const includeLanguage = vscode.workspace.getConfiguration('shareToTeams').get('includeLanguage', true);
            const formattedText = formatCodeSnippet(selectedText, includeLanguage ? languageId : null);

            await vscode.window.withProgress({
                location: vscode.ProgressLocation.Notification,
                title: "Sharing code to Teams",
                cancellable: false
            }, async (progress) => {
                progress.report({ message: "Authenticating..." });
                try {
                    const directShareResult = await shareViaGraphApi(formattedText, context, progress);
                    if (directShareResult) {
                        return;
                    }
                    progress.report({ message: "Using Teams deep link..." });
                    await shareViaDeepLink(formattedText);
                } catch (error) {
                    console.error('Error sharing:', error);
                    const fallback = await vscode.window.showErrorMessage(
                        `Error: ${error.message}. Use Teams deep link instead?`,
                        'Yes', 'No'
                    );
                    if (fallback === 'Yes') {
                        await shareViaDeepLink(formattedText);
                    }
                }
            });
        } catch (error) {
            console.error('Error sharing code to Teams:', error);
            vscode.window.showErrorMessage(`Error sharing code to Teams: ${error.message}`);
        }
    });

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
    progress.report({ message: "Getting access token..." });
    const accessToken = await auth.getAccessToken();
    
    if (!accessToken) {
        throw new Error('Authentication failed');
    }

    progress.report({ message: "Fetching Teams chats..." });
    const chats = await teamsService.getChats(accessToken);
    
    if (!chats || chats.length === 0) {
        throw new Error('No Teams chats found. Ensure you have active chats in Microsoft Teams.');
    }

    const chatItems = chats.map(chat => ({
        label: getChatLabel(chat),
        id: chat.id,
        description: getChatDescription(chat)
    }));

    const selectedChat = await vscode.window.showQuickPick(chatItems, {
        placeHolder: 'Select a Teams chat to share the code',
        matchOnDescription: true
    });

    if (!selectedChat) {
        return false; // User cancelled
    }

    progress.report({ message: `Sending message to ${selectedChat.label}...` });
    await teamsService.sendMessage(accessToken, selectedChat.id, formattedText);
    
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
        return chat.topic; // Group chat
    }
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
        return chat.members ? `${chat.members.length} members` : 'Group chat';
    }
    return 'Direct message';
}

/**
 * Share code via Teams deep link
 * @param {string} formattedText - Formatted code block
 */
async function shareViaDeepLink(formattedText) {
    const encodedMessage = encodeURIComponent(formattedText);
    const deepLink = `msteams://teams.microsoft.com/l/chat/0/0?message=${encodedMessage}`;
    await vscode.env.openExternal(vscode.Uri.parse(deepLink));
    vscode.window.showInformationMessage('Microsoft Teams opened. Please select recipients to share your code.');
}

/**
 * Deactivates the extension
 */
function deactivate() {}

module.exports = {
    activate,
    deactivate,
    formatCodeSnippet
};