/* Direct-to-Matter Save - Outlook Add-in */
/* Integrates with sp-bridge infrastructure for matter search and SharePoint upload */

(function() {
    'use strict';

    // Configuration - Update these URLs to match your deployment
    const CONFIG = {
        API_BASE: 'https://workplace-continuity-nelson-manner.trycloudflare.com',  // sp-bridge via Cloudflare Tunnel
        API_KEY: '1a3f307b2ba7f56a80fc9c1898644c76440b0cb3b69c5fa1cf70fe72c40a2828',  // shared secret for matter_search.py
        GRAPH_BASE: 'https://graph.microsoft.com/v1.0',
        CLIENT_ID: '6b7c22e5-e8bc-416c-97c8-7f1b0cba335a',
        TENANT_ID: '612c697b-b95c-4ad9-a9df-58d52ed8eff1',
        SEARCH_DEBOUNCE_MS: 350
    };

    // State management
    let state = {
        attachments: [],
        selectedAttachments: new Set(),
        searchResults: [],
        selectedMatter: null,
        selectedFolder: null,
        isLoading: false,
        accessToken: null
    };

    // DOM elements
    let elements = {};

    // Initialize the add-in
    Office.onReady((info) => {
        if (info.host === Office.HostType.Outlook) {
            initializeApp();
        }
    });

    function initializeApp() {
        cacheElements();
        bindEvents();
        loadAttachments();
        updateStatus('Ready to save attachments', 'pending');
    }

    function cacheElements() {
        elements = {
            statusContent: document.querySelector('.status-content'),
            statusIcon: document.querySelector('.status-icon'),
            statusText: document.querySelector('.status-text'),
            attachmentList: document.getElementById('attachmentList'),
            attachmentSection: document.getElementById('attachmentSection'),
            searchSection: document.getElementById('searchSection'),
            searchInput: document.getElementById('searchInput'),
            resultsList: document.getElementById('resultsList'),
            folderSelection: document.getElementById('folderSelection'),
            customFolderInput: document.getElementById('customFolderInput'),
            cancelBtn: document.getElementById('cancelBtn'),
            saveBtn: document.getElementById('saveBtn'),
            progressBar: document.getElementById('progressBar'),
            progressFill: document.getElementById('progressFill')
        };
    }

    function bindEvents() {
        // Search input with debouncing
        let searchTimeout;
        elements.searchInput.addEventListener('input', (e) => {
            clearTimeout(searchTimeout);
            searchTimeout = setTimeout(() => {
                performSearch(e.target.value.trim());
            }, CONFIG.SEARCH_DEBOUNCE_MS);
        });

        // Folder selection
        document.querySelectorAll('.folder-option').forEach(option => {
            option.addEventListener('click', () => selectFolder(option.dataset.folder));
        });

        // Custom folder input
        elements.customFolderInput.addEventListener('input', (e) => {
            if (e.target.value.trim()) {
                clearFolderSelection();
                state.selectedFolder = e.target.value.trim();
                updateSaveButton();
            }
        });

        // Action buttons
        elements.cancelBtn.addEventListener('click', closeAddIn);
        elements.saveBtn.addEventListener('click', saveAttachments);
    }

    async function loadAttachments() {
        try {
            updateStatus('Loading attachments...', 'pending');
            
            // Get current email item
            const item = Office.context.mailbox.item;
            
            if (!item.attachments || item.attachments.length === 0) {
                updateStatus('No attachments found in this email', 'error');
                elements.attachmentSection.classList.add('hidden');
                return;
            }

            // Process attachments
            state.attachments = item.attachments.map(attachment => ({
                id: attachment.id,
                name: attachment.name,
                size: attachment.size || 0,
                type: attachment.attachmentType,
                selected: true  // Default to all selected
            }));

            // Add all to selected set
            state.attachments.forEach(att => state.selectedAttachments.add(att.id));

            renderAttachments();
            showSearchSection();
            updateStatus(`Found ${state.attachments.length} attachment(s)`, 'success');

        } catch (error) {
            console.error('Error loading attachments:', error);
            updateStatus('Error loading attachments', 'error');
        }
    }

    function renderAttachments() {
        const html = state.attachments.map(attachment => `
            <li class="attachment-item">
                <div class="attachment-checkbox ${attachment.selected ? 'checked' : ''}" 
                     data-id="${attachment.id}"></div>
                <div class="attachment-info">
                    <div class="attachment-name">${escapeHtml(attachment.name)}</div>
                    <div class="attachment-size">${formatFileSize(attachment.size)}</div>
                </div>
            </li>
        `).join('');

        elements.attachmentList.innerHTML = html;

        // Bind checkbox events
        document.querySelectorAll('.attachment-checkbox').forEach(checkbox => {
            checkbox.addEventListener('click', (e) => {
                toggleAttachment(e.target.dataset.id);
            });
        });
    }

    function toggleAttachment(attachmentId) {
        const attachment = state.attachments.find(a => a.id === attachmentId);
        if (attachment) {
            attachment.selected = !attachment.selected;
            
            if (attachment.selected) {
                state.selectedAttachments.add(attachmentId);
            } else {
                state.selectedAttachments.delete(attachmentId);
            }

            renderAttachments();
            updateSaveButton();
        }
    }

    function showSearchSection() {
        elements.searchSection.classList.remove('hidden');
        elements.searchInput.focus();
    }

    async function performSearch(query) {
        if (!query || query.length < 2) {
            elements.resultsList.innerHTML = '';
            return;
        }

        try {
            updateStatus('Searching matters...', 'pending');
            
            const response = await fetch(`${CONFIG.API_BASE}/search?q=${encodeURIComponent(query)}`, {
                headers: { 'X-API-Key': CONFIG.API_KEY }
            });
            
            if (!response.ok) {
                throw new Error(`Search failed: ${response.status}`);
            }

            const data = await response.json();
            state.searchResults = data.results || [];
            
            renderSearchResults();
            
            if (state.searchResults.length > 0) {
                updateStatus(`Found ${state.searchResults.length} matter(s)`, 'success');
            } else {
                updateStatus('No matters found', 'pending');
            }

        } catch (error) {
            console.error('Search error:', error);
            updateStatus('Search failed - check connection', 'error');
            elements.resultsList.innerHTML = '<div class="result-item" style="color: var(--error-red); text-align: center;">Search service unavailable</div>';
        }
    }

   function renderSearchResults() {
        const html = state.searchResults.map(matter => `
            <div class="result-item" data-matter-number="${escapeHtml(matter.matter_number)}">
                <div class="matter-number">${escapeHtml(matter.matter_number)}</div>
                <div class="matter-name">${escapeHtml(matter.display_name.split(' - ').slice(1).join(' - '))}</div>
            </div>
        `).join('');

        elements.resultsList.innerHTML = html;

        // Bind click events
        document.querySelectorAll('.result-item').forEach(item => {
            item.addEventListener('click', () => {
                const matterNumber = item.dataset.matterNumber;
                const matter = state.searchResults.find(m => m.matter_number === matterNumber);
                if (matter) selectMatter(matter);
            });
        });
    }
    
 
    async function selectMatter(matter) {
        try {
            // Highlight selected matter
            document.querySelectorAll('.result-item').forEach(item => {
                item.classList.remove('selected');
            });
            event.target.closest('.result-item').classList.add('selected');

            // Get detailed matter information
            updateStatus('Loading matter details...', 'pending');
            
            const response = await fetch(`${CONFIG.API_BASE}/matter/${matter.matter_number}`, {
                headers: { 'X-API-Key': CONFIG.API_KEY }
            });
            
            if (!response.ok) {
                throw new Error(`Failed to load matter details: ${response.status}`);
            }

            const matterDetails = await response.json();
            
            // Validate matter has SharePoint integration
            if (!matterDetails.site_id || !matterDetails.drive_id) {
                updateStatus('Matter not configured for SharePoint', 'error');
                return;
            }

            state.selectedMatter = matterDetails;
            showFolderSelection();
            updateStatus(`Selected: ${matter.matter_number}`, 'success');

        } catch (error) {
            console.error('Error selecting matter:', error);
            updateStatus('Failed to load matter details', 'error');
        }
    }

    function showFolderSelection() {
        elements.folderSelection.classList.remove('hidden');
        
        // Clear previous selection
        clearFolderSelection();
    }

    function selectFolder(folderName) {
        clearFolderSelection();
        
        const option = document.querySelector(`[data-folder="${folderName}"]`);
        if (option) {
            option.classList.add('selected');
        }
        
        state.selectedFolder = folderName;
        elements.customFolderInput.value = '';
        
        updateSaveButton();
    }

    function clearFolderSelection() {
        document.querySelectorAll('.folder-option').forEach(option => {
            option.classList.remove('selected');
        });
    }

    function updateSaveButton() {
        const hasSelectedAttachments = state.selectedAttachments.size > 0;
        const hasMatter = state.selectedMatter !== null;
        const hasFolder = state.selectedFolder !== null;
        
        if (hasSelectedAttachments && hasMatter && hasFolder) {
            elements.saveBtn.classList.remove('hidden');
            elements.saveBtn.textContent = `Save ${state.selectedAttachments.size} file(s) to ${state.selectedFolder}`;
        } else {
            elements.saveBtn.classList.add('hidden');
        }
    }

    async function saveAttachments() {
        if (state.isLoading) return;
        
        try {
            state.isLoading = true;
            updateStatus('Preparing upload...', 'pending');
            showProgress(10);
            
            // Get access token
            const token = await getAccessToken();
            if (!token) {
                throw new Error('Failed to authenticate with SharePoint');
            }
            
            showProgress(25);
            updateStatus('Uploading attachments...', 'pending');
            
            const selectedAttachments = state.attachments.filter(att => 
                state.selectedAttachments.has(att.id)
            );
            
            let successCount = 0;
            
            // Upload each selected attachment
            for (let i = 0; i < selectedAttachments.length; i++) {
                const attachment = selectedAttachments[i];
                
                updateStatus(`Uploading ${attachment.name}...`, 'pending');
                
                try {
                    await uploadAttachment(attachment, token);
                    successCount++;
                    
                    const progress = 25 + (70 * (i + 1) / selectedAttachments.length);
                    showProgress(progress);
                    
                } catch (error) {
                    console.error(`Failed to upload ${attachment.name}:`, error);
                    // Continue with other attachments
                }
            }
            
            showProgress(100);
            
            if (successCount === selectedAttachments.length) {
                updateStatus(`Successfully saved ${successCount} file(s)`, 'success');
                setTimeout(() => closeAddIn(), 2000);
            } else {
                updateStatus(`Saved ${successCount} of ${selectedAttachments.length} file(s)`, 'error');
            }
            
        } catch (error) {
            console.error('Save error:', error);
            updateStatus('Upload failed', 'error');
        } finally {
            state.isLoading = false;
            hideProgress();
        }
    }

    async function uploadAttachment(attachment, accessToken) {
        return new Promise((resolve, reject) => {
            // Get attachment content from Outlook
            Office.context.mailbox.item.getAttachmentContentAsync(
                attachment.id,
                { asyncContext: { attachment, accessToken } },
                async (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        reject(new Error(`Failed to get attachment content: ${result.error.message}`));
                        return;
                    }

                    try {
                        const { attachment: att, accessToken: token } = result.asyncContext;
                        const content = result.value.content;
                        
                        // Upload to SharePoint
                        await uploadToSharePoint(att, content, token);
                        resolve();
                        
                    } catch (error) {
                        reject(error);
                    }
                }
            );
        });
    }

    async function uploadToSharePoint(attachment, content, accessToken) {
        const driveId = state.selectedMatter.drive_id;
        const fileName = attachment.name;
        const folderPath = state.selectedFolder;
        
        // Build upload URL
        let uploadUrl = `${CONFIG.GRAPH_BASE}/drives/${driveId}/root:`;
        
        if (folderPath && folderPath !== 'root') {
            uploadUrl += `/${encodeURIComponent(folderPath)}`;
        }
        
        uploadUrl += `/${encodeURIComponent(fileName)}:/content`;
        
        // Convert base64 to blob for upload
        const byteCharacters = atob(content);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        
        const response = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/octet-stream'
            },
            body: byteArray
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`SharePoint upload failed: ${response.status} - ${errorText}`);
        }
        
        return await response.json();
    }

    async function getAccessToken() {
        // This is a simplified version - in production you'd implement proper certificate-based authentication
        // For now, return null to indicate authentication needs to be implemented
        
        updateStatus('Authentication setup required', 'error');
        return null;
        
        // TODO: Implement certificate-based authentication with Microsoft Graph
        // See: https://docs.microsoft.com/en-us/graph/auth-v2-service
    }

    function updateStatus(message, type = 'pending') {
        elements.statusText.textContent = message;
        
        // Update icon and styling
        elements.statusIcon.className = `status-icon status-${type}`;
        
        const icons = {
            pending: '⏳',
            success: '✅',
            error: '❌'
        };
        
        elements.statusIcon.textContent = icons[type] || '⏳';
    }

    function showProgress(percentage) {
        elements.progressBar.classList.remove('hidden');
        elements.progressFill.style.width = `${Math.min(100, Math.max(0, percentage))}%`;
        document.body.classList.add('loading');
    }

    function hideProgress() {
        elements.progressBar.classList.add('hidden');
        elements.progressFill.style.width = '0%';
        document.body.classList.remove('loading');
    }

    function closeAddIn() {
        if (Office.context.ui) {
            Office.context.ui.closeContainer();
        }
    }

    // Utility functions
    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(1024));
        
        return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${sizes[i]}`;
    }

})();
