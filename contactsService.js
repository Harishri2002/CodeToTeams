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
        // First verify we can access the user's profile
        let userProfile;
        try {
            const profileResponse = await axios.get('https://graph.microsoft.com/v1.0/me', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            userProfile = profileResponse.data;
            console.log('User profile fetched successfully via axios:', userProfile.id);
        } catch (profileError) {
            console.error('Error fetching user profile:', profileError.response?.status, profileError.response?.data?.error);
            throw new Error(`Unable to access your profile. Status: ${profileError.response?.status}. You may need to sign out and sign in again.`);
        }

        // Try the people API endpoint
        try {
            const contactsResponse = await axios.get('https://graph.microsoft.com/v1.0/me/people?$filter=personType/class eq \'Person\'&$top=100', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (contactsResponse.data && contactsResponse.data.value) {
                const myEmail = userProfile.mail || userProfile.userPrincipalName;
                
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
                
                console.log('Contacts fetched via people API:', contacts.length, 'contacts found');
                return contacts;
            }
        } catch (peopleError) {
            console.log('Unable to fetch contacts via people API:', peopleError.response?.status, peopleError.response?.data?.error);
            // Continue to try other endpoints
        }

        // If people API fails, try the users API as a fallback
        try {
            const usersResponse = await axios.get('https://graph.microsoft.com/v1.0/users?$top=50&$select=id,displayName,mail,userPrincipalName,department,jobTitle', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (usersResponse.data && usersResponse.data.value) {
                const users = usersResponse.data.value
                    .filter(user => user.mail || user.userPrincipalName)
                    .map(user => ({
                        id: user.id,
                        displayName: user.displayName,
                        email: user.mail || user.userPrincipalName,
                        department: user.department || '',
                        jobTitle: user.jobTitle || '',
                        userPrincipalName: user.userPrincipalName
                    }));
                
                console.log('Users fetched as fallback for contacts:', users.length, 'users found');
                return users;
            }
        } catch (usersError) {
            console.log('Unable to fetch users list:', usersError.response?.status, usersError.response?.data?.error);
            // Last fallback option before failing
        }

        // If we can't get contacts or users, at least include the current user
        if (userProfile && (userProfile.mail || userProfile.userPrincipalName)) {
            const currentUser = {
                id: userProfile.id,
                displayName: userProfile.displayName || userProfile.mail || userProfile.userPrincipalName,
                email: userProfile.mail || userProfile.userPrincipalName,
                department: userProfile.department || '',
                jobTitle: userProfile.jobTitle || '',
                userPrincipalName: userProfile.userPrincipalName
            };
            
            console.log('Using only current user as contact');
            return [currentUser];
        }
        
        throw new Error('Unable to fetch any contacts or users from Microsoft Graph API');
    } catch (error) {
        console.error('Error getting contacts:', error);
        
        // Provide more specific error message based on the error
        if (error.response?.status === 401) {
            throw new Error('Authentication failed. Please sign out and sign in again.');
        } else if (error.response?.status === 403) {
            throw new Error('Permission denied. Your account may not have the necessary permissions.');
        }
        
        throw new Error(`Failed to get contacts: ${error.message}`);
    }
}

/**
 * Get colleagues from the same organization
 * @param {string} accessToken - Microsoft Graph API access token
 * @returns {Promise<Array>} - List of colleagues
 */
async function getColleagues(accessToken) {
    try {
        const response = await axios.get('https://graph.microsoft.com/v1.0/users?$top=50&$select=id,displayName,mail,userPrincipalName,department,jobTitle', {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });

        if (response.data && response.data.value) {
            const colleagues = response.data.value
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
        }
        
        return [];
    } catch (error) {
        console.log('Error getting colleagues:', error.response?.status, error.response?.data?.error);
        // Don't throw, just return empty array if colleagues can't be fetched
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
        const contactsPromise = getContacts(accessToken).catch(error => {
            console.log('Failed to get contacts:', error.message);
            return [];
        });
        
        const colleaguesPromise = getColleagues(accessToken).catch(error => {
            console.log('Failed to get colleagues:', error.message);
            return [];
        });
        
        const [contacts, colleagues] = await Promise.all([contactsPromise, colleaguesPromise]);
        
        // If we have no contacts from either source, throw error
        if (contacts.length === 0 && colleagues.length === 0) {
            throw new Error('Could not find any contacts. Please check your permissions or try adding contacts manually.');
        }
        
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