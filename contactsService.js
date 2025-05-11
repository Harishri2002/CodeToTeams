const { Client } = require('@microsoft/microsoft-graph-client');
const vscode = require('vscode');
const axios = require('axios');

/**
 * Get the user's profile and contacts
 * @param {string} accessToken - Microsoft Graph API access token
 * @returns {Promise<Array>} - List of contacts
 */
async function getContacts(accessToken) {
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

        // Get people (contacts and colleagues)
        console.log('Fetching people...');
        const response = await client
            .api('/me/people')
            .filter("personType/class eq 'Person'")
            .top(100)
            .get();

        // Store the main user email for reference
        const myEmail = userProfile.mail || userProfile.userPrincipalName;
        console.log('Current user email:', myEmail);

        // Filter and map contacts
        const contacts = response.value
            .filter(person => person.scoredEmailAddresses && person.scoredEmailAddresses.length > 0)
            .map(person => ({
                id: person.id,
                displayName: person.displayName,
                email: person.scoredEmailAddresses[0].address,
                department: person.department || '',
                company: person.companyName || '',
                jobTitle: person.jobTitle || '',
                userPrincipalName: person.userPrincipalName || person.scoredEmailAddresses[0].address
            }));

        console.log('People/contacts fetched successfully:', contacts.length, 'contacts found');
        return contacts;
    } catch (error) {
        console.error('Error getting contacts:', error);
        
        try {
            console.log('Falling back to axios for fetching contacts...');
            const profileResponse = await axios.get('https://graph.microsoft.com/v1.0/me', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            const contactsResponse = await axios.get('https://graph.microsoft.com/v1.0/me/people?$filter=personType/class eq \'Person\'&$top=100', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (contactsResponse.data && contactsResponse.data.value) {
                const myEmail = profileResponse.data.mail || profileResponse.data.userPrincipalName;
                
                const contacts = contactsResponse.data.value
                    .filter(person => person.scoredEmailAddresses && person.scoredEmailAddresses.length > 0)
                    .map(person => ({
                        id: person.id,
                        displayName: person.displayName,
                        email: person.scoredEmailAddresses[0].address,
                        department: person.department || '',
                        company: person.companyName || '',
                        jobTitle: person.jobTitle || '',
                        userPrincipalName: person.userPrincipalName || person.scoredEmailAddresses[0].address
                    }));
                
                console.log('Contacts fetched via axios fallback:', contacts.length, 'contacts found');
                return contacts;
            }
        } catch (axiosError) {
            console.error('Axios fallback also failed:', axiosError);
        }
        
        throw new Error('Failed to get contacts. Please ensure you have granted the necessary permissions.');
    }
}

/**
 * Get colleagues from the same organization
 * @param {string} accessToken - Microsoft Graph API access token
 * @returns {Promise<Array>} - List of colleagues
 */
async function getColleagues(accessToken) {
    try {
        const client = Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        // Get colleagues by searching for users in the organization
        console.log('Fetching colleagues...');
        const response = await client
            .api('/users')
            .top(50)
            .select('id,displayName,mail,userPrincipalName,department,jobTitle')
            .get();

        const colleagues = response.value
            .filter(user => user.mail || user.userPrincipalName)
            .map(user => ({
                id: user.id,
                displayName: user.displayName,
                email: user.mail || user.userPrincipalName,
                department: user.department || '',
                jobTitle: user.jobTitle || '',
                userPrincipalName: user.userPrincipalName
            }));

        console.log('Colleagues fetched successfully:', colleagues.length, 'colleagues found');
        return colleagues;
    } catch (error) {
        console.error('Error getting colleagues:', error);
        
        // If we don't have permission to get users, just return empty array
        // We can still use the contacts from people endpoint
        console.log('Unable to fetch colleagues - this is okay, will use contacts only');
        return [];
    }
}

/**
 * Get a combined list of unique contacts and colleagues
 * @param {string} accessToken - Microsoft Graph API access token
 * @returns {Promise<Array>} - List of unique recipients
 */
async function getPotentialRecipients(accessToken) {
    try {
        // Get both contacts and colleagues
        const [contacts, colleagues] = await Promise.all([
            getContacts(accessToken),
            getColleagues(accessToken).catch(error => {
                console.log('Failed to get colleagues, continuing with contacts only:', error.message);
                return [];
            })
        ]);
        
        // Combine and deduplicate based on email
        const emailMap = new Map();
        
        // Add contacts first
        contacts.forEach(contact => {
            if (contact.email) {
                emailMap.set(contact.email.toLowerCase(), contact);
            }
        });
        
        // Add colleagues (prioritizing colleagues if there's a duplicate)
        colleagues.forEach(colleague => {
            if (colleague.email) {
                emailMap.set(colleague.email.toLowerCase(), colleague);
            }
        });
        
        // Convert map back to array
        const uniqueRecipients = Array.from(emailMap.values());
        
        // Sort by display name
        uniqueRecipients.sort((a, b) => a.displayName.localeCompare(b.displayName));
        
        console.log('Combined unique recipients:', uniqueRecipients.length);
        return uniqueRecipients;
    } catch (error) {
        console.error('Error getting potential recipients:', error);
        throw new Error(`Failed to get contacts: ${error.message}`);
    }
}

/**
 * Create a deep link URL for sharing to Teams
 * @param {Array<string>} emails - Email addresses of recipients
 * @param {string} content - Message content
 * @returns {string} - Teams deep link URL
 */
function createTeamsDeepLink(emails, content) {
    const encodedUsers = encodeURIComponent(emails.join(','));
    const encodedMessage = encodeURIComponent(content);
    return `https://teams.microsoft.com/l/chat/0/0?users=${encodedUsers}&message=${encodedMessage}`;
}

module.exports = {
    getContacts,
    getColleagues,
    getPotentialRecipients,
    createTeamsDeepLink
};