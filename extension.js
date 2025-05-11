const vscode = require('vscode');
const auth = require('./authentication');
const contactsService = require('./contactsService');

// Cache for recipients to avoid fetching again in a session
let recipientsCache = null;
let lastFetchTime = null;
const CACHE_EXPIRY_MS = 10 * 60 * 1000; // 10 minutes

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

            // Format the code snippet with language information
            const languageId = editor.document.languageId;
            const includeLanguage = vscode.workspace.getConfiguration('shareToTeams').get('includeLanguage', true);
            const formattedText = formatCodeSnippet(selectedText, includeLanguage ? languageId : null);

            await vscode.window.withProgress({
                location: vscode.ProgressLocation.Notification,
                title: "Share code snippet to Teams",
                cancellable: false
            }, async (progress) => {
                progress.report({ message: "Authenticating..." });
                
                // Get access token
                const accessToken = await auth.getAccessToken();
                if (!accessToken) {
                    throw new Error('Authentication failed');
                }

                // Get recipients
                progress.report({ message: "Fetching contacts..." });
                const recipients = await getRecipients(accessToken);
                
                if (!recipients || recipients.length === 0) {
                    throw new Error('No contacts found. Please ensure you have contacts in your Microsoft account.');
                }

                // Show QuickPick UI for selecting recipients
                progress.report({ message: "Select contacts to share with..." });
                const selectedRecipients = await showRecipientSelector(recipients);
                
                if (!selectedRecipients || selectedRecipients.length === 0) {
                    return; // User cancelled
                }

                // Create deep link and open it
                progress.report({ message: "Opening Teams..." });
                const emails = selectedRecipients.map(r => r.email);
                const deepLink = contactsService.createTeamsDeepLink(emails, formattedText);
                await vscode.env.openExternal(vscode.Uri.parse(deepLink));
                
                vscode.window.showInformationMessage(`Code shared to Teams for ${selectedRecipients.length} recipient(s).`);
            });
        } catch (error) {
            console.error('Error sharing code to Teams:', error);
            vscode.window.showErrorMessage(`Error sharing code to Teams: ${error.message}`);
        }
    });

    let signOutCommand = vscode.commands.registerCommand('extension.teamsSignOut', async function() {
        await auth.signOut();
        // Clear cache when signing out
        recipientsCache = null;
        lastFetchTime = null;
    });

    let addManualRecipientCommand = vscode.commands.registerCommand('extension.addManualTeamsRecipient', async function() {
        const email = await vscode.window.showInputBox({
            prompt: 'Enter email address of the Teams user',
            placeHolder: 'example@company.com',
            validateInput: (text) => {
                // Simple email validation
                return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text) ? null : 'Please enter a valid email address';
            }
        });
        
        if (email) {
            // Store in settings
            const config = vscode.workspace.getConfiguration('shareToTeams');
            const manualRecipients = config.get('manualRecipients', []);
            
            // Check if already exists
            if (!manualRecipients.includes(email)) {
                manualRecipients.push(email);
                await config.update('manualRecipients', manualRecipients, vscode.ConfigurationTarget.Global);
                vscode.window.showInformationMessage(`Added ${email} to your Teams recipients`);
            } else {
                vscode.window.showInformationMessage(`${email} is already in your recipients list`);
            }
            
            // Clear cache to include the new manual recipient
            recipientsCache = null;
        }
    });

    context.subscriptions.push(shareCommand, signOutCommand, addManualRecipientCommand);
}

/**
 * Get recipients from cache or fetch fresh
 * @param {string} accessToken - Microsoft Graph API access token
 * @returns {Promise<Array>} - List of recipients
 */
async function getRecipients(accessToken) {
    // Check if cache is valid
    const now = Date.now();
    if (recipientsCache && lastFetchTime && (now - lastFetchTime < CACHE_EXPIRY_MS)) {
        console.log('Using cached recipients:', recipientsCache.length);
        return recipientsCache;
    }
    
    // Fetch fresh data
    const apiRecipients = await contactsService.getPotentialRecipients(accessToken);
    
    // Add manual recipients from settings
    const config = vscode.workspace.getConfiguration('shareToTeams');
    const manualRecipients = config.get('manualRecipients', []);
    
    // Combine API recipients with manual recipients
    const emailMap = new Map();
    
    // Add API recipients
    apiRecipients.forEach(recipient => {
        emailMap.set(recipient.email.toLowerCase(), recipient);
    });
    
    // Add manual recipients
    manualRecipients.forEach(email => {
        const lowerEmail = email.toLowerCase();
        if (!emailMap.has(lowerEmail)) {
            emailMap.set(lowerEmail, {
                id: `manual-${lowerEmail}`,
                displayName: email,
                email: email,
                isManual: true
            });
        }
    });
    
    // Update cache
    recipientsCache = Array.from(emailMap.values());
    lastFetchTime = now;
    
    return recipientsCache;
}

/**
 * Show UI for selecting recipients
 * @param {Array} recipients - List of recipients
 * @returns {Promise<Array>} - Selected recipients
 */
async function showRecipientSelector(recipients) {
    // Create QuickPick items with detailed info
    const items = recipients.map(r => ({
        label: r.displayName,
        description: r.email,
        detail: getRecipientDetail(r),
        recipient: r
    }));
    
    // Create multi-select QuickPick
    const quickPick = vscode.window.createQuickPick();
    quickPick.title = "Select Teams Recipients";
    quickPick.placeholder = "Search by name or email";
    quickPick.items = items;
    quickPick.canSelectMany = true;
    
    // Add buttons
    quickPick.buttons = [
        {
            iconPath: new vscode.ThemeIcon('add'),
            tooltip: 'Add new recipient'
        }
    ];
    
    // Return promise that resolves when selection is made
    return new Promise((resolve) => {
        quickPick.onDidAccept(() => {
            const selected = quickPick.selectedItems.map(item => item.recipient);
            quickPick.hide();
            resolve(selected);
        });
        
        quickPick.onDidHide(() => {
            resolve([]);
        });
        
        quickPick.onDidTriggerButton(() => {
            // This will execute the command asynchronously
            vscode.commands.executeCommand('extension.addManualTeamsRecipient');
            // Keep the QuickPick open
            quickPick.busy = true;
            setTimeout(() => { quickPick.busy = false; }, 500);
        });
        
        quickPick.show();
    });
}

/**
 * Get detailed information for a recipient
 * @param {Object} recipient - Recipient object
 * @returns {string} - Formatted details
 */
function getRecipientDetail(recipient) {
    const parts = [];
    
    if (recipient.isManual) {
        parts.push('Custom recipient');
    } else {
        if (recipient.jobTitle) parts.push(recipient.jobTitle);
        if (recipient.department) parts.push(recipient.department);
        if (recipient.company) parts.push(recipient.company);
    }
    
    return parts.join(' • ');
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
 * Deactivates the extension
 */
function deactivate() {}

module.exports = {
    activate,
    deactivate,
    formatCodeSnippet
};