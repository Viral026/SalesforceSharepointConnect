import { LightningElement, track, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import searchDocuments from '@salesforce/apex/SharePointDocumentController.searchDocuments';
import getRecentDocuments from '@salesforce/apex/SharePointDocumentController.getRecentDocuments';
import getAllDocuments from '@salesforce/apex/SharePointDocumentController.getAllDocuments';
import getFoldersAndFiles from '@salesforce/apex/SharePointDocumentController.getFoldersAndFiles';
import uploadDocument from '@salesforce/apex/SharePointDocumentController.uploadDocument';
import createFolder from '@salesforce/apex/SharePointDocumentController.createFolder';

export default class KnowledgeCenter extends LightningElement {
    @track documents = [];
    @track recentDocuments = [];
    @track searchKeyword = '';
    @track selectedCategory = 'All';
    @track isLoading = false;
    @track errorMessage = '';
    @track hasSearched = false;
    @track isSearchActive = false; // Track if user performed manual search
    @track currentFolderId = 'root'; // Track current folder location
    @track breadcrumbs = []; // Navigation breadcrumb trail
    @track isUploading = false; // Track upload status
    @track selectedFile = null; // Selected file for upload
    @track isCreatingFolder = false; // Track folder creation status
    @track newFolderName = ''; // New folder name input
    @track viewMode = 'grid'; // View mode: 'grid' or 'list'
    @track showUploadModal = false; // Show/hide upload modal
    @track showFolderModal = false; // Show/hide folder modal
    searchTimeout; // Debounce timer for auto-search
    pendingItemsByFolder = {}; // Track recently uploaded/created items per folder (key: folderId, value: array of items)
    syncRetryCount = 0; // Track number of sync retry attempts
    maxSyncRetries = 3; // Maximum number of retry attempts

    // Category options for the filter dropdown
    categoryOptions = [
        { label: 'All Categories', value: 'All' },
        { label: 'Alerts', value: 'Alerts' },
        { label: 'Functional Documentation', value: 'Functional Docs' },
        { label: 'Technical Documentation', value: 'Technical Docs' },
        { label: 'Standard Operating Procedures (SOPs)', value: 'SOPs' },
        { label: 'General', value: 'General' }
    ];

    // Lifecycle hook - load folders and files on component initialization
    connectedCallback() {
        this.initializeBreadcrumbs();
        this.loadFolderContents('root');
        this.loadRecentDocuments();
    }

    /**
     * Initialize breadcrumbs with root
     */
    initializeBreadcrumbs() {
        this.breadcrumbs = [{
            id: 'root',
            name: 'Home',
            isRoot: true
        }];
    }

    /**
     * Load folder contents (folders and files)
     */
    loadFolderContents(folderId) {
        this.isLoading = true;
        this.isSearchActive = false;
        this.errorMessage = '';
        this.currentFolderId = folderId;

        getFoldersAndFiles({ folderId: folderId })
            .then(result => {
                if (result && result.length > 0) {
                    const enrichedResults = this.enrichDocuments(result);

                    // Merge with any pending items for this folder
                    const pendingItems = this.pendingItemsByFolder[folderId] || [];
                    const pendingItemsStillMissing = [];

                    pendingItems.forEach(pendingItem => {
                        const foundInResults = enrichedResults.some(doc =>
                            doc.name === pendingItem.name && doc.isFolder === pendingItem.isFolder
                        );
                        if (!foundInResults) {
                            // If it's a pending folder, update its item count if it has pending items inside
                            if (pendingItem.isFolder && pendingItem.id) {
                                const pendingItemsInFolder = this.pendingItemsByFolder[pendingItem.id] || [];
                                if (pendingItemsInFolder.length > 0) {
                                    pendingItem.size = pendingItemsInFolder.length;
                                    pendingItem.formattedSize = pendingItemsInFolder.length + ' item' + (pendingItemsInFolder.length !== 1 ? 's' : '');
                                }
                            }
                            pendingItemsStillMissing.push(pendingItem);
                        }
                    });

                    // Update folder item counts to include pending items inside them
                    enrichedResults.forEach(doc => {
                        if (doc.isFolder && doc.id) {
                            const pendingItemsInFolder = this.pendingItemsByFolder[doc.id] || [];
                            if (pendingItemsInFolder.length > 0) {
                                // Update the folder's size to include pending items
                                const currentCount = parseInt(doc.size) || 0;
                                const totalCount = currentCount + pendingItemsInFolder.length;
                                doc.size = totalCount;
                                doc.formattedSize = totalCount + ' item' + (totalCount !== 1 ? 's' : '');
                                console.log(`📁 Updated folder "${doc.name}" count: ${totalCount} (${pendingItemsInFolder.length} pending)`);
                            }
                        }
                    });

                    // Update pending items for this folder
                    if (pendingItemsStillMissing.length > 0) {
                        this.pendingItemsByFolder[folderId] = pendingItemsStillMissing;
                    } else {
                        delete this.pendingItemsByFolder[folderId];
                    }

                    // Merge and display
                    this.documents = [...pendingItemsStillMissing, ...enrichedResults];
                    this.hasSearched = true;
                    console.log('Loaded ' + result.length + ' items from SharePoint');
                    if (pendingItemsStillMissing.length > 0) {
                        console.log('📊 Plus ' + pendingItemsStillMissing.length + ' pending items');
                    }
                } else {
                    // Show pending items if any, even if SharePoint returns empty
                    const pendingItems = this.pendingItemsByFolder[folderId] || [];
                    this.documents = [...pendingItems];
                    this.hasSearched = true;
                }
            })
            .catch(error => {
                console.error('Error loading folder contents:', error);
                this.errorMessage = this.getErrorMessage(error);
                this.showToast('Error', 'Failed to load items from SharePoint', 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    /**
     * Load all documents from SharePoint (recursive - for search)
     */
    loadAllDocuments() {
        this.isLoading = true;
        this.isSearchActive = false; // Not a manual search
        getAllDocuments()
            .then(result => {
                if (result && result.length > 0) {
                    this.documents = this.enrichDocuments(result);
                    this.hasSearched = true; // Mark as searched so documents display
                    console.log('Loaded ' + result.length + ' documents from SharePoint');
                } else {
                    this.documents = [];
                    this.hasSearched = true;
                }
            })
            .catch(error => {
                console.error('Error loading all documents:', error);
                this.errorMessage = this.getErrorMessage(error);
                this.showToast('Error', 'Failed to load documents from SharePoint', 'error');
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    /**
     * Load recently modified documents
     */
    loadRecentDocuments() {
        getRecentDocuments({ limitCount: 5 })
            .then(result => {
                if (result && result.length > 0) {
                    this.recentDocuments = this.enrichDocuments(result);
                }
            })
            .catch(error => {
                console.error('Error loading recent documents:', error);
                // Don't show error for recent documents failure
            });
    }

    /**
     * Handle search input change
     */
    handleSearchInput(event) {
        this.searchKeyword = event.target.value;

        // Clear existing timeout
        if (this.searchTimeout) {
            clearTimeout(this.searchTimeout);
        }

        // If search is cleared (empty), automatically return to folder view
        if (!this.searchKeyword || this.searchKeyword.trim() === '') {
            this.selectedCategory = 'All';
            this.errorMessage = '';
            // Return to current folder or root
            this.initializeBreadcrumbs();
            this.loadFolderContents('root');
            return;
        }

        // Auto-search after typing 3 or more characters with debounce
        if (this.searchKeyword.trim().length >= 3) {
            this.searchTimeout = setTimeout(() => {
                this.performSearch();
            }, 500); // 500ms debounce delay
        }
    }

    /**
     * Handle category filter change
     */
    handleCategoryChange(event) {
        this.selectedCategory = event.detail.value;
    }

    /**
     * Handle search button click
     */
    handleSearch() {
        // Validate search input
        if (!this.searchKeyword || this.searchKeyword.trim().length === 0) {
            this.showToast('Warning', 'Please enter a search keyword', 'warning');
            return;
        }

        // Clear any pending auto-search
        if (this.searchTimeout) {
            clearTimeout(this.searchTimeout);
        }

        this.performSearch();
    }

    /**
     * Perform search operation
     */
    performSearch() {
        this.isLoading = true;
        this.errorMessage = '';
        this.hasSearched = true;
        this.isSearchActive = true; // User performed manual search
        this.documents = [];
        // Don't clear pending items during search - they're folder-specific

        // Search all documents with keyword
        this.searchAllDocuments();
    }

    /**
     * Search all documents with keyword
     */
    searchAllDocuments() {
        searchDocuments({ keyword: this.searchKeyword })
            .then(result => {
                this.handleSearchResults(result);
            })
            .catch(error => {
                this.handleSearchError(error);
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    /**
     * Handle search results
     */
    handleSearchResults(result) {
        if (result && result.length > 0) {
            this.documents = this.enrichDocuments(result);
            const fileCount = result.filter(item => !item.isFolder).length;
            const folderCount = result.filter(item => item.isFolder).length;
            let message = `Found ${result.length} item(s)`;
            if (fileCount > 0 && folderCount > 0) {
                message = `Found ${fileCount} file(s) and ${folderCount} folder(s)`;
            } else if (fileCount > 0) {
                message = `Found ${fileCount} file(s)`;
            } else if (folderCount > 0) {
                message = `Found ${folderCount} folder(s)`;
            }
            this.showToast('Success', message, 'success');
        } else {
            this.documents = [];
            this.showToast('Info', 'No files or folders found matching your search criteria', 'info');
        }
    }

    /**
     * Handle search error
     */
    handleSearchError(error) {
        this.errorMessage = this.getErrorMessage(error);
        this.showToast('Error', this.errorMessage, 'error');
        console.error('Search error:', error);
    }

    /**
     * Handle clear button click
     */
    handleClear() {
        this.searchKeyword = '';
        this.selectedCategory = 'All';
        this.errorMessage = '';
        this.showToast('Info', 'Search cleared - returning to folder view', 'info');
        // Return to current folder or root
        this.initializeBreadcrumbs();
        this.loadFolderContents('root');
    }

    /**
     * Handle folder click to navigate into folder
     */
    handleFolderClick(event) {
        const folderId = event.currentTarget.dataset.id;
        const folderName = event.currentTarget.dataset.name;

        if (folderId) {
            // Add to breadcrumbs
            this.breadcrumbs.push({
                id: folderId,
                name: folderName,
                isRoot: false
            });

            // Load folder contents
            this.loadFolderContents(folderId);
        }
    }

    /**
     * Handle breadcrumb navigation
     */
    handleBreadcrumbClick(event) {
        const folderId = event.currentTarget.dataset.id;
        const folderIndex = this.breadcrumbs.findIndex(b => b.id === folderId);

        if (folderIndex !== -1) {
            // Remove all breadcrumbs after clicked one
            this.breadcrumbs = this.breadcrumbs.slice(0, folderIndex + 1);

            // Load that folder's contents
            this.loadFolderContents(folderId);
        }
    }

    /**
     * Handle view document button click
     */
    handleViewDocument(event) {
        const documentUrl = event.target.dataset.url;
        if (documentUrl) {
            window.open(documentUrl, '_blank');
        } else {
            this.showToast('Error', 'Document URL not available', 'error');
        }
    }

    /**
     * Handle download document button click
     */
    handleDownloadDocument(event) {
        const documentUrl = event.target.dataset.url;
        if (documentUrl) {
            // Open document URL which will trigger download based on SharePoint settings
            window.open(documentUrl + '?download=1', '_blank');
            this.showToast('Success', 'Document download initiated', 'success');
        } else {
            this.showToast('Error', 'Document URL not available', 'error');
        }
    }

    /**
     * Handle open upload modal
     */
    handleOpenUploadModal() {
        this.showUploadModal = true;
    }

    /**
     * Handle close upload modal
     */
    handleCloseUploadModal() {
        if (!this.isUploading) {
            this.showUploadModal = false;
            this.selectedFile = null;
            // Reset the file input
            const fileInput = this.template.querySelector('input[type="file"]');
            if (fileInput) {
                fileInput.value = '';
            }
        }
    }

    /**
     * Handle open folder modal
     */
    handleOpenFolderModal() {
        this.showFolderModal = true;
    }

    /**
     * Handle close folder modal
     */
    handleCloseFolderModal() {
        if (!this.isCreatingFolder) {
            this.showFolderModal = false;
            this.newFolderName = '';
        }
    }

    /**
     * Handle choose file button click
     */
    handleChooseFile() {
        const fileInput = this.template.querySelector('input[type="file"]');
        if (fileInput) {
            fileInput.click();
        }
    }

    /**
     * Handle file selection for upload
     */
    handleFileChange(event) {
        const files = event.target.files;
        if (files && files.length > 0) {
            this.selectedFile = files[0];
            console.log('File selected: ' + this.selectedFile.name);
        } else {
            this.selectedFile = null;
        }
    }

    /**
     * Handle remove selected file
     */
    handleRemoveFile() {
        this.selectedFile = null;
        // Reset the file input
        const fileInput = this.template.querySelector('input[type="file"]');
        if (fileInput) {
            fileInput.value = '';
        }
        console.log('File removed');
    }

    /**
     * Handle file upload to SharePoint
     */
    handleUploadFile() {
        if (!this.selectedFile) {
            this.showToast('Warning', 'Please select a file to upload', 'warning');
            return;
        }

        // Check file size (max 4MB - Graph API simple upload limit)
        const maxSize = 4 * 1024 * 1024; // 4MB
        if (this.selectedFile.size > maxSize) {
            this.showToast('Error', 'File size exceeds 4MB limit. Please upload a smaller file.', 'error');
            return;
        }

        this.isUploading = true;
        const fileName = this.selectedFile.name;
        const reader = new FileReader();

        reader.onload = () => {
            const base64Content = reader.result.split(',')[1]; // Remove data:* prefix

            uploadDocument({
                fileName: fileName,
                fileContent: base64Content,
                folderId: this.currentFolderId
            })
            .then(result => {
                // Result contains the uploaded file ID
                console.log('Upload successful, file ID:', result);
                this.showToast('Success', 'File uploaded successfully to SharePoint', 'success');

                // Optimistically add the uploaded file to the UI immediately
                this.addUploadedFileToUI(fileName, result, this.selectedFile.size);

                // Clear file input
                this.selectedFile = null;
                const fileInput = this.template.querySelector('input[type="file"]');
                if (fileInput) {
                    fileInput.value = '';
                }

                // Reset uploading state and close modal
                this.isUploading = false;
                this.showUploadModal = false;

                // Reset retry counter for new upload
                this.syncRetryCount = 0;

                // Refresh in background to sync with actual SharePoint data
                setTimeout(() => {
                    this.refreshAfterUpload();
                }, 2000);
            })
            .catch(error => {
                console.error('Error uploading file:', error);
                this.showToast('Error', 'Failed to upload file: ' + this.getErrorMessage(error), 'error');
                this.isUploading = false;
            });
        };

        reader.onerror = () => {
            this.showToast('Error', 'Error reading file', 'error');
            this.isUploading = false;
        };

        reader.readAsDataURL(this.selectedFile);
    }

    /**
     * Handle folder name input change
     */
    handleFolderNameChange(event) {
        this.newFolderName = event.target.value;
    }

    /**
     * Handle create folder button click
     */
    handleCreateFolder() {
        if (!this.newFolderName || this.newFolderName.trim().length === 0) {
            this.showToast('Warning', 'Please enter a folder name', 'warning');
            return;
        }

        // Validate folder name (no special characters)
        const invalidChars = /[<>:"/\\|?*]/;
        if (invalidChars.test(this.newFolderName)) {
            this.showToast('Error', 'Folder name contains invalid characters', 'error');
            return;
        }

        // Check if folder with same name already exists
        const folderExists = this.documents.some(doc =>
            doc.isFolder && doc.name.toLowerCase() === this.newFolderName.trim().toLowerCase()
        );

        if (folderExists) {
            this.showToast('Error', 'A folder with the name "' + this.newFolderName.trim() + '" already exists in this location', 'error');
            return;
        }

        this.isCreatingFolder = true;

        createFolder({
            folderName: this.newFolderName.trim(),
            parentFolderId: this.currentFolderId
        })
            .then(result => {
                // Result contains the created folder ID
                const createdFolderName = this.newFolderName;
                console.log('Folder created successfully, ID:', result);
                this.showToast('Success', 'Folder "' + createdFolderName + '" created successfully', 'success');

                // Optimistically add the created folder to the UI immediately
                this.addCreatedFolderToUI(createdFolderName, result);

                // Clear folder name input
                this.newFolderName = '';

                // Reset creating state and close modal
                this.isCreatingFolder = false;
                this.showFolderModal = false;

                // Reset retry counter for new folder
                this.syncRetryCount = 0;

                // Refresh in background to sync with actual SharePoint data
                setTimeout(() => {
                    this.refreshAfterFolderCreation();
                }, 2000);
            })
            .catch(error => {
                console.error('Error creating folder:', error);
                this.showToast('Error', 'Failed to create folder: ' + this.getErrorMessage(error), 'error');
                this.isCreatingFolder = false;
            });
    }

    /**
     * Add uploaded file to UI optimistically
     */
    addUploadedFileToUI(fileName, fileId, fileSize) {
        const newFile = {
            id: fileId || 'temp_' + Date.now(),
            name: fileName,
            size: fileSize || 0,
            isFolder: false,
            webUrl: '#', // Placeholder URL (will be updated on sync)
            modifiedDateTime: new Date().toISOString(),
            iconName: this.getFileIcon(fileName),
            formattedSize: this.formatFileSize(fileSize || 0),
            itemType: 'File',
            isPending: true // Mark as pending to preserve during sync
        };

        // Add to pending items for this folder
        if (!this.pendingItemsByFolder[this.currentFolderId]) {
            this.pendingItemsByFolder[this.currentFolderId] = [];
        }
        this.pendingItemsByFolder[this.currentFolderId].push(newFile);

        // Create new array to ensure reactivity
        this.documents = [newFile, ...this.documents];

        // Ensure hasSearched is true so results display
        this.hasSearched = true;

        console.log('✅ File added to UI optimistically:', fileName);
        console.log('📊 Total documents now:', this.documents.length);
    }

    /**
     * Add created folder to UI optimistically
     */
    addCreatedFolderToUI(folderName, folderId) {
        const newFolder = {
            id: folderId || 'temp_' + Date.now(),
            name: folderName,
            size: 0,
            isFolder: true,
            webUrl: '#', // Placeholder URL (will be updated on sync)
            modifiedDateTime: new Date().toISOString(),
            iconName: 'doctype:folder',
            formattedSize: '0 items',
            itemType: 'Folder',
            isPending: true // Mark as pending to preserve during sync
        };

        // Add to pending items for this folder
        if (!this.pendingItemsByFolder[this.currentFolderId]) {
            this.pendingItemsByFolder[this.currentFolderId] = [];
        }
        this.pendingItemsByFolder[this.currentFolderId].push(newFolder);

        // Create new array to ensure reactivity
        this.documents = [newFolder, ...this.documents];

        // Ensure hasSearched is true so results display
        this.hasSearched = true;

        console.log('✅ Folder added to UI optimistically:', folderName);
        console.log('📊 Total documents now:', this.documents.length);
    }

    /**
     * Refresh folder contents after file upload (background sync)
     */
    refreshAfterUpload() {
        const targetFolderId = this.currentFolderId;
        // Refresh current folder contents in background
        getFoldersAndFiles({ folderId: targetFolderId })
            .then(result => {
                if (result && result.length > 0) {
                    const enrichedResults = this.enrichDocuments(result);

                    // Check if pending items are now in SharePoint results
                    const pendingItems = this.pendingItemsByFolder[targetFolderId] || [];
                    const pendingItemsStillMissing = [];
                    pendingItems.forEach(pendingItem => {
                        const foundInResults = enrichedResults.some(doc =>
                            doc.name === pendingItem.name && doc.isFolder === pendingItem.isFolder
                        );
                        if (!foundInResults) {
                            // SharePoint hasn't indexed this yet, keep it
                            // If it's a pending folder, update its item count if it has pending items inside
                            if (pendingItem.isFolder && pendingItem.id) {
                                const pendingItemsInFolder = this.pendingItemsByFolder[pendingItem.id] || [];
                                if (pendingItemsInFolder.length > 0) {
                                    pendingItem.size = pendingItemsInFolder.length;
                                    pendingItem.formattedSize = pendingItemsInFolder.length + ' item' + (pendingItemsInFolder.length !== 1 ? 's' : '');
                                }
                            }
                            pendingItemsStillMissing.push(pendingItem);
                        }
                    });

                    // Update pending items for this folder
                    if (pendingItemsStillMissing.length > 0) {
                        this.pendingItemsByFolder[targetFolderId] = pendingItemsStillMissing;
                    } else {
                        delete this.pendingItemsByFolder[targetFolderId];
                    }

                    // Update folder item counts to include pending items inside them
                    enrichedResults.forEach(doc => {
                        if (doc.isFolder && doc.id) {
                            const pendingItemsInFolder = this.pendingItemsByFolder[doc.id] || [];
                            if (pendingItemsInFolder.length > 0) {
                                const currentCount = parseInt(doc.size) || 0;
                                const totalCount = currentCount + pendingItemsInFolder.length;
                                doc.size = totalCount;
                                doc.formattedSize = totalCount + ' item' + (totalCount !== 1 ? 's' : '');
                            }
                        }
                    });

                    // Only update UI if we're still viewing this folder
                    if (this.currentFolderId === targetFolderId) {
                        // Merge: SharePoint results + any pending items not yet in SharePoint
                        this.documents = [...pendingItemsStillMissing, ...enrichedResults];
                    }

                    console.log('✅ Folder contents synced with SharePoint after upload');
                    console.log('📊 Pending items still awaiting indexing:', pendingItemsStillMissing.length);

                    // If there are still pending items and we haven't exceeded retry limit, retry sync
                    if (pendingItemsStillMissing.length > 0 && this.syncRetryCount < this.maxSyncRetries) {
                        this.syncRetryCount++;
                        setTimeout(() => {
                            console.log(`🔄 Retrying sync for pending items (attempt ${this.syncRetryCount}/${this.maxSyncRetries})...`);
                            this.refreshAfterUpload();
                        }, 3000); // Retry after 3 more seconds
                    } else if (pendingItemsStillMissing.length > 0) {
                        console.log('⚠️ Max retry attempts reached. Keeping optimistic items.');
                        this.syncRetryCount = 0; // Reset for future uploads
                    } else {
                        this.syncRetryCount = 0; // Reset for future uploads
                    }
                } else {
                    // Keep pending items even if SharePoint returns nothing
                    const pendingItems = this.pendingItemsByFolder[targetFolderId] || [];
                    if (this.currentFolderId === targetFolderId) {
                        this.documents = [...pendingItems];
                    }
                }
                // Also refresh recent documents to show the newly uploaded file
                this.loadRecentDocuments();
            })
            .catch(error => {
                console.error('Error syncing after upload:', error);
                // Silent failure - optimistic update already shown
            });
    }

    /**
     * Refresh folder contents after folder creation (background sync)
     */
    refreshAfterFolderCreation() {
        const targetFolderId = this.currentFolderId;
        // Refresh current folder contents in background
        getFoldersAndFiles({ folderId: targetFolderId })
            .then(result => {
                if (result && result.length > 0) {
                    const enrichedResults = this.enrichDocuments(result);

                    // Check if pending items are now in SharePoint results
                    const pendingItems = this.pendingItemsByFolder[targetFolderId] || [];
                    const pendingItemsStillMissing = [];
                    pendingItems.forEach(pendingItem => {
                        const foundInResults = enrichedResults.some(doc =>
                            doc.name === pendingItem.name && doc.isFolder === pendingItem.isFolder
                        );
                        if (!foundInResults) {
                            // SharePoint hasn't indexed this yet, keep it
                            // If it's a pending folder, update its item count if it has pending items inside
                            if (pendingItem.isFolder && pendingItem.id) {
                                const pendingItemsInFolder = this.pendingItemsByFolder[pendingItem.id] || [];
                                if (pendingItemsInFolder.length > 0) {
                                    pendingItem.size = pendingItemsInFolder.length;
                                    pendingItem.formattedSize = pendingItemsInFolder.length + ' item' + (pendingItemsInFolder.length !== 1 ? 's' : '');
                                }
                            }
                            pendingItemsStillMissing.push(pendingItem);
                        }
                    });

                    // Update pending items for this folder
                    if (pendingItemsStillMissing.length > 0) {
                        this.pendingItemsByFolder[targetFolderId] = pendingItemsStillMissing;
                    } else {
                        delete this.pendingItemsByFolder[targetFolderId];
                    }

                    // Update folder item counts to include pending items inside them
                    enrichedResults.forEach(doc => {
                        if (doc.isFolder && doc.id) {
                            const pendingItemsInFolder = this.pendingItemsByFolder[doc.id] || [];
                            if (pendingItemsInFolder.length > 0) {
                                const currentCount = parseInt(doc.size) || 0;
                                const totalCount = currentCount + pendingItemsInFolder.length;
                                doc.size = totalCount;
                                doc.formattedSize = totalCount + ' item' + (totalCount !== 1 ? 's' : '');
                            }
                        }
                    });

                    // Only update UI if we're still viewing this folder
                    if (this.currentFolderId === targetFolderId) {
                        // Merge: SharePoint results + any pending items not yet in SharePoint
                        this.documents = [...pendingItemsStillMissing, ...enrichedResults];
                    }

                    console.log('✅ Folder contents synced with SharePoint after folder creation');
                    console.log('📊 Pending items still awaiting indexing:', pendingItemsStillMissing.length);

                    // If there are still pending items and we haven't exceeded retry limit, retry sync
                    if (pendingItemsStillMissing.length > 0 && this.syncRetryCount < this.maxSyncRetries) {
                        this.syncRetryCount++;
                        setTimeout(() => {
                            console.log(`🔄 Retrying sync for pending items (attempt ${this.syncRetryCount}/${this.maxSyncRetries})...`);
                            this.refreshAfterFolderCreation();
                        }, 3000); // Retry after 3 more seconds
                    } else if (pendingItemsStillMissing.length > 0) {
                        console.log('⚠️ Max retry attempts reached. Keeping optimistic items.');
                        this.syncRetryCount = 0; // Reset for future operations
                    } else {
                        this.syncRetryCount = 0; // Reset for future operations
                    }
                } else {
                    // Keep pending items even if SharePoint returns nothing
                    const pendingItems = this.pendingItemsByFolder[targetFolderId] || [];
                    if (this.currentFolderId === targetFolderId) {
                        this.documents = [...pendingItems];
                    }
                }
            })
            .catch(error => {
                console.error('Error syncing after folder creation:', error);
                // Silent failure - optimistic update already shown
            });
    }

    /**
     * Enrich documents with additional properties for UI display
     */
    enrichDocuments(docs) {
        return docs.map(doc => {
            return {
                ...doc,
                iconName: doc.isFolder ? 'doctype:folder' : this.getFileIcon(doc.name),
                formattedSize: doc.isFolder ? (doc.size + ' items') : this.formatFileSize(doc.size),
                itemType: doc.isFolder ? 'Folder' : 'File'
            };
        });
    }

    /**
     * Get file icon based on file extension
     */
    getFileIcon(fileName) {
        if (!fileName) return 'doctype:attachment';

        const extension = fileName.split('.').pop().toLowerCase();

        const iconMap = {
            'pdf': 'doctype:pdf',
            'doc': 'doctype:word',
            'docx': 'doctype:word',
            'xls': 'doctype:excel',
            'xlsx': 'doctype:excel',
            'ppt': 'doctype:ppt',
            'pptx': 'doctype:ppt',
            'txt': 'doctype:txt',
            'zip': 'doctype:zip',
            'png': 'doctype:image',
            'jpg': 'doctype:image',
            'jpeg': 'doctype:image',
            'gif': 'doctype:image',
            'csv': 'doctype:csv',
            'xml': 'doctype:xml',
            'json': 'doctype:xml'
        };

        return iconMap[extension] || 'doctype:attachment';
    }

    /**
     * Format file size to human-readable format
     */
    formatFileSize(bytes) {
        if (!bytes || bytes === 0) return 'N/A';

        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(1024));
        const size = (bytes / Math.pow(1024, i)).toFixed(2);

        return size + ' ' + sizes[i];
    }

    /**
     * Extract error message from error object
     */
    getErrorMessage(error) {
        if (error.body && error.body.message) {
            return error.body.message;
        } else if (error.message) {
            return error.message;
        } else if (typeof error === 'string') {
            return error;
        }
        return 'An unexpected error occurred. Please try again.';
    }

    /**
     * Show toast notification
     */
    showToast(title, message, variant) {
        const event = new ShowToastEvent({
            title: title,
            message: message,
            variant: variant
        });
        this.dispatchEvent(event);
    }

    /**
     * Computed property - check if there are search results
     */
    get hasResults() {
        return !this.isLoading && this.documents && this.documents.length > 0;
    }

    /**
     * Computed property - check if should show no results message
     */
    get showNoResults() {
        return !this.isLoading && this.hasSearched &&
               (!this.documents || this.documents.length === 0) &&
               !this.errorMessage;
    }

    /**
     * Computed property - check if should show recent documents
     */
    get showRecentDocuments() {
        return !this.hasSearched && !this.isLoading &&
               this.recentDocuments && this.recentDocuments.length > 0;
    }

    /**
     * Computed property - get header text based on search state
     */
    get resultsHeaderText() {
        if (this.isSearchActive) {
            const fileCount = this.documents.filter(item => !item.isFolder).length;
            const folderCount = this.documents.filter(item => item.isFolder).length;
            return `Search Results (${fileCount} file(s), ${folderCount} folder(s))`;
        }
        const currentFolder = this.breadcrumbs[this.breadcrumbs.length - 1];
        const folderName = currentFolder ? currentFolder.name : 'Home';
        return folderName + ' (' + this.documents.length + ' items)';
    }

    /**
     * Computed property - show breadcrumbs if not in search mode
     */
    get showBreadcrumbs() {
        return !this.isSearchActive && this.breadcrumbs && this.breadcrumbs.length > 0;
    }

    /**
     * Computed property - breadcrumbs with current folder highlighted
     */
    get breadcrumbsWithCurrent() {
        if (!this.breadcrumbs || this.breadcrumbs.length === 0) {
            return [];
        }
        return this.breadcrumbs.map((crumb, index) => ({
            ...crumb,
            isCurrent: index === this.breadcrumbs.length - 1
        }));
    }

    /**
     * Computed property - upload button label
     */
    get uploadButtonLabel() {
        return this.selectedFile ? 'Upload Now' : 'Choose File';
    }

    /**
     * Computed property - selected file name
     */
    get selectedFileName() {
        return this.selectedFile ? this.selectedFile.name : null;
    }

    /**
     * Computed property - selected file size formatted
     */
    get selectedFileSize() {
        return this.selectedFile ? this.formatFileSize(this.selectedFile.size) : null;
    }

    /**
     * Computed property - upload button disabled state
     */
    get uploadDisabled() {
        return !this.selectedFile || this.isUploading;
    }

    /**
     * Computed property - create folder button disabled state
     */
    get createFolderDisabled() {
        return !this.newFolderName || this.newFolderName.trim().length === 0 || this.isCreatingFolder;
    }

    /**
     * Computed property - upload modal container class
     */
    get uploadModalContainerClass() {
        return this.isUploading ? 'slds-modal__container modal-transparent' : 'slds-modal__container';
    }

    /**
     * Computed property - folder modal container class
     */
    get folderModalContainerClass() {
        return this.isCreatingFolder ? 'slds-modal__container modal-transparent' : 'slds-modal__container';
    }

    /**
     * Computed property - is grid view active
     */
    get isGridView() {
        return this.viewMode === 'grid';
    }

    /**
     * Computed property - is list view active
     */
    get isListView() {
        return this.viewMode === 'list';
    }

    /**
     * Computed property - grid button variant
     */
    get gridButtonVariant() {
        return this.viewMode === 'grid' ? 'brand' : 'border-filled';
    }

    /**
     * Computed property - list button variant
     */
    get listButtonVariant() {
        return this.viewMode === 'list' ? 'brand' : 'border-filled';
    }

    /**
     * Handle view mode toggle
     */
    handleViewToggle(event) {
        const selectedView = event.currentTarget.dataset.view;
        if (selectedView) {
            this.viewMode = selectedView;
        }
    }

    /**
     * Computed property - check if can go back (not at root)
     */
    get canGoBack() {
        return this.breadcrumbs && this.breadcrumbs.length > 1;
    }
}