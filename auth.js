/* SharePoint Authentication Module */
/* Handles certificate-based authentication with Microsoft Graph API */

class SharePointAuth {
    constructor(clientId, tenantId, certificatePath) {
        this.clientId = clientId;
        this.tenantId = tenantId;
        this.certificatePath = certificatePath;
        this.tokenCache = new Map();
        this.tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    }

    /**
     * Get access token using certificate-based authentication
     * @returns {Promise<string|null>} Access token or null if authentication fails
     */
    async getAccessToken() {
        const cacheKey = 'sharepoint_token';
        const cached = this.tokenCache.get(cacheKey);
        
        // Check if we have a valid cached token
        if (cached && cached.expires > Date.now() + 60000) { // 1 minute buffer
            return cached.token;
        }

        try {
            // In a real implementation, you would:
            // 1. Load your certificate from the certificate store
            // 2. Create a JWT assertion signed with the certificate
            // 3. Exchange the JWT for an access token
            
            // For now, this is a placeholder that shows the structure
            const token = await this.requestTokenWithCertificate();
            
            if (token) {
                // Cache the token with expiration
                this.tokenCache.set(cacheKey, {
                    token: token.access_token,
                    expires: Date.now() + (token.expires_in * 1000)
                });
                
                return token.access_token;
            }
            
            return null;
            
        } catch (error) {
            console.error('Authentication failed:', error);
            return null;
        }
    }

    /**
     * Request token using certificate-based client credentials flow
     * @private
     */
    async requestTokenWithCertificate() {
        // This is where you would implement the actual certificate authentication
        // The process involves:
        
        // 1. Create JWT assertion
        const assertion = await this.createJWTAssertion();
        
        // 2. Request token
        const response = await fetch(this.tokenUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: new URLSearchParams({
                'client_id': this.clientId,
                'scope': 'https://graph.microsoft.com/.default',
                'client_assertion_type': 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
                'client_assertion': assertion,
                'grant_type': 'client_credentials'
            })
        });

        if (!response.ok) {
            const error = await response.text();
            throw new Error(`Token request failed: ${response.status} - ${error}`);
        }

        return await response.json();
    }

    /**
     * Create JWT assertion signed with certificate
     * @private
     */
    async createJWTAssertion() {
        // This would require a JWT library and certificate handling
        // For now, this is a placeholder
        
        // In a real implementation:
        // 1. Load certificate from Windows Certificate Store or file
        // 2. Create JWT with proper claims
        // 3. Sign with certificate private key
        
        throw new Error('Certificate-based authentication not yet implemented. Please configure your SharePoint app with certificate authentication.');
    }

    /**
     * Clear cached tokens
     */
    clearCache() {
        this.tokenCache.clear();
    }
}

// Alternative: Simplified authentication for development/testing
class SharePointAuthDev {
    constructor(clientId, tenantId, clientSecret) {
        this.clientId = clientId;
        this.tenantId = tenantId;
        this.clientSecret = clientSecret;
        this.tokenCache = new Map();
        this.tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    }

    /**
     * Get access token using client secret (for development only)
     * @returns {Promise<string|null>}
     */
    async getAccessToken() {
        const cacheKey = 'sharepoint_token_dev';
        const cached = this.tokenCache.get(cacheKey);
        
        if (cached && cached.expires > Date.now() + 60000) {
            return cached.token;
        }

        try {
            const response = await fetch(this.tokenUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
                body: new URLSearchParams({
                    'client_id': this.clientId,
                    'client_secret': this.clientSecret,
                    'scope': 'https://graph.microsoft.com/.default',
                    'grant_type': 'client_credentials'
                })
            });

            if (!response.ok) {
                const error = await response.text();
                throw new Error(`Token request failed: ${response.status} - ${error}`);
            }

            const token = await response.json();
            
            this.tokenCache.set(cacheKey, {
                token: token.access_token,
                expires: Date.now() + (token.expires_in * 1000)
            });
            
            return token.access_token;
            
        } catch (error) {
            console.error('Dev authentication failed:', error);
            return null;
        }
    }
}

// Export for use in main application
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { SharePointAuth, SharePointAuthDev };
} else {
    window.SharePointAuth = SharePointAuth;
    window.SharePointAuthDev = SharePointAuthDev;
}
