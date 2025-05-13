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

        // Get account type to handle differently for personal vs. work accounts
        const accountType = userProfile.userPrincipalName?.includes('#EXT#') || 
                           userProfile.userPrincipalName?.includes('@') && 
                           !userProfile.userPrincipalName?.includes('outlook.com') && 
                           !userProfile.userPrincipalName?.includes('hotmail.com') && 
                           !userProfile.userPrincipalName?.includes('live.com') ? 
                           'work' : 'personal';
        
        console.log(`Account type detected: ${accountType}`);
        
        let contacts = [];
        
        // Try to get contacts directly (works for both personal and work accounts with Contacts.Read permission)
        try {
            const contactsResponse = await axios.get('https://graph.microsoft.com/v1.0/me/contacts?$top=100', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (contactsResponse.data && contactsResponse.data.value) {
                const directContacts = contactsResponse.data.value
                    .filter(contact => contact.emailAddresses && contact.emailAddresses.length > 0)
                    .map(contact => ({
                        id: contact.id,
                        displayName: contact.displayName || contact.emailAddresses[0].address,
                        email: contact.emailAddresses[0].address,
                        department: contact.department || '',
                        company: contact.companyName || '',
                        jobTitle: contact.jobTitle || '',
                        userPrincipalName: contact.emailAddresses[0].address
                    }));
                
                console.log('Contacts fetched via contacts API:', directContacts.length, 'contacts found');
                contacts = [...contacts, ...directContacts];
            }
        } catch (contactsError) {
            console.log('Unable to fetch direct contacts:', contactsError.response?.status, contactsError.response?.data?.error);
            // Continue to try other endpoints
        }

        // Try the people API endpoint as well
        try {
            const peopleResponse = await axios.get('https://graph.microsoft.com/v1.0/me/people?$filter=personType/class eq \'Person\'&$top=100', {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });
            
            if (peopleResponse.data && peopleResponse.data.value) {
                const myEmail = userProfile.mail || userProfile.userPrincipalName;
                
                const peopleContacts = peopleResponse.data.value
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
                
                console.log('Contacts fetched via people API:', peopleContacts.length, 'contacts found');
                
                // Add people contacts but avoid duplicates based on email
                const emailSet = new Set(contacts.map(c => c.email.toLowerCase()));
                const uniquePeopleContacts = peopleContacts.filter(p => !emailSet.has(p.email.toLowerCase()));
                
                contacts = [...contacts, ...uniquePeopleContacts];
            }
        } catch (peopleError) {
            console.log('Unable to fetch contacts via people API:', peopleError.response?.status, peopleError.response?.data?.error);
            // Continue to try other endpoints
        }

        // For work accounts, try the users API as a fallback
        if (accountType === 'work' && contacts.length < 5) {
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
                    
                    // Add users but avoid duplicates based on email
                    const emailSet = new Set(contacts.map(c => c.email.toLowerCase()));
                    const uniqueUsers = users.filter(u => !emailSet.has(u.email.toLowerCase()));
                    
                    contacts = [...contacts, ...uniqueUsers];
                }
            } catch (usersError) {
                console.log('Unable to fetch users list:', usersError.response?.status, usersError.response?.data?.error);
                // Last fallback option before failing
            }
        }

        // If we haven't found any contacts, at least include the current user
        if (contacts.length === 0 && userProfile && (userProfile.mail || userProfile.userPrincipalName)) {
            const currentUser = {
                id: userProfile.id,
                displayName: userProfile.displayName || userProfile.mail || userProfile.userPrincipalName,
                email: userProfile.mail || userProfile.userPrincipalName,
                department: userProfile.department || '',
                jobTitle: userProfile.jobTitle || '',
                userPrincipalName: userProfile.userPrincipalName
            };
            
            console.log('Using only current user as contact');
            contacts = [currentUser];
        }
        
        if (contacts.length === 0) {
            throw new Error('Unable to fetch any contacts or users from Microsoft Graph API');
        }
        
        return contacts;
    } catch (error) {
        console.error('Error getting contacts:', error);
        
        // Provide more specific error message based on the error
        if (error.response?.status === 401) {
            throw new Error('Authentication failed. Please sign out and sign in again.');
        } else if (error.response?.status === 403) {
            throw new Error('Permission denied. Your account may not have the necessary permissions. Make sure you authorized Contacts.Read permissions.');
        }
        
        throw new Error(`Failed to get contacts: ${error.message}`);
    }
}

/**
 * Get colleagues from the same organization (only for work accounts)
 * @param {string} accessToken - Microsoft Graph API access token
 * @returns {Promise<Array>} - List of colleagues
 */
async function getColleagues(accessToken) {
    try {
        // First check if this is a work account by verifying the organization name
        const orgResponse = await axios.get('https://graph.microsoft.com/v1.0/organization', {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!orgResponse.data || !orgResponse.data.value || orgResponse.data.value.length === 0) {
            console.log('No organization found, likely a personal account');
            return [];
        }

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
        console.log('Error getting colleagues or not a work account:', error.response?.status, error.response?.data?.error);
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