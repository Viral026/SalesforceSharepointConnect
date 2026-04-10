/**
 * @description Test class for SharePointDocumentController
 */
@isTest
private class SharePointDocumentControllerTest {

    /**
     * @description Test searchDocuments method with successful response
     */
    @isTest
    static void testSearchDocumentsSuccess() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockHttpResponseSuccess());

        Test.startTest();
        List<SharePointDocumentController.DocumentInfo> results =
            SharePointDocumentController.searchDocuments('test');
        Test.stopTest();

        System.assertNotEquals(null, results, 'Results should not be null');
        System.assertEquals(true, results.size() > 0, 'Should return documents');
    }

    /**
     * @description Test searchDocuments method with blank keyword
     */
    @isTest
    static void testSearchDocumentsBlankKeyword() {
        Test.startTest();
        List<SharePointDocumentController.DocumentInfo> results =
            SharePointDocumentController.searchDocuments('');
        Test.stopTest();

        System.assertNotEquals(null, results, 'Results should not be null');
        System.assertEquals(0, results.size(), 'Should return empty list for blank keyword');
    }

    /**
     * @description Test searchDocuments method with error response
     */
    @isTest
    static void testSearchDocumentsError() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockHttpResponseError());

        Test.startTest();
        try {
            SharePointDocumentController.searchDocuments('test');
            System.assert(false, 'Should have thrown an exception');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('Error'), 'Should contain error message');
        }
        Test.stopTest();
    }

    /**
     * @description Test getRecentDocuments method with limit
     */
    @isTest
    static void testGetRecentDocuments() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockHttpResponseSuccess());

        Test.startTest();
        List<SharePointDocumentController.DocumentInfo> results =
            SharePointDocumentController.getRecentDocuments(5);
        Test.stopTest();

        System.assertNotEquals(null, results, 'Results should not be null');
    }

    /**
     * @description Test getRecentDocuments method with null limit
     */
    @isTest
    static void testGetRecentDocumentsNullLimit() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockHttpResponseSuccess());

        Test.startTest();
        List<SharePointDocumentController.DocumentInfo> results =
            SharePointDocumentController.getRecentDocuments(null);
        Test.stopTest();

        System.assertNotEquals(null, results, 'Results should not be null');
    }

    /**
     * @description Test getFoldersAndFiles with root folder
     */
    @isTest
    static void testGetFoldersAndFilesRoot() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockFoldersAndFilesSuccess());

        Test.startTest();
        List<SharePointDocumentController.DocumentInfo> results =
            SharePointDocumentController.getFoldersAndFiles('root');
        Test.stopTest();

        System.assertNotEquals(null, results, 'Results should not be null');
        System.assertEquals(true, results.size() > 0, 'Should return items');

        // Verify folder is marked correctly
        Boolean hasFolder = false;
        for(SharePointDocumentController.DocumentInfo item : results) {
            if(item.isFolder) {
                hasFolder = true;
                break;
            }
        }
        System.assertEquals(true, hasFolder, 'Should contain at least one folder');
    }

    /**
     * @description Test getFoldersAndFiles with specific folder ID
     */
    @isTest
    static void testGetFoldersAndFilesWithId() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockFoldersAndFilesSuccess());

        Test.startTest();
        List<SharePointDocumentController.DocumentInfo> results =
            SharePointDocumentController.getFoldersAndFiles('01FOLDER123');
        Test.stopTest();

        System.assertNotEquals(null, results, 'Results should not be null');
    }

    /**
     * @description Test getAllDocuments method
     */
    @isTest
    static void testGetAllDocuments() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockFoldersAndFilesSuccess());

        Test.startTest();
        List<SharePointDocumentController.DocumentInfo> results =
            SharePointDocumentController.getAllDocuments();
        Test.stopTest();

        System.assertNotEquals(null, results, 'Results should not be null');
    }

    /**
     * @description Test uploadDocument success to root
     */
    @isTest
    static void testUploadDocumentSuccess() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockUploadSuccess());

        String base64Content = EncodingUtil.base64Encode(Blob.valueOf('Test file content'));

        Test.startTest();
        String fileId = SharePointDocumentController.uploadDocument(
            'TestFile.txt',
            base64Content,
            'root'
        );
        Test.stopTest();

        System.assertNotEquals(null, fileId, 'File ID should not be null');
        System.assertEquals('UPLOADED_FILE_ID', fileId, 'Should return correct file ID');
    }

    /**
     * @description Test uploadDocument success to specific folder
     */
    @isTest
    static void testUploadDocumentToFolder() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockUploadSuccess());

        String base64Content = EncodingUtil.base64Encode(Blob.valueOf('Test file content'));

        Test.startTest();
        String fileId = SharePointDocumentController.uploadDocument(
            'TestFile.txt',
            base64Content,
            '01FOLDER123'
        );
        Test.stopTest();

        System.assertNotEquals(null, fileId, 'File ID should not be null');
    }

    /**
     * @description Test uploadDocument error
     */
    @isTest
    static void testUploadDocumentError() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockHttpResponseError());

        String base64Content = EncodingUtil.base64Encode(Blob.valueOf('Test file content'));

        Test.startTest();
        try {
            SharePointDocumentController.uploadDocument('TestFile.txt', base64Content, 'root');
            System.assert(false, 'Should have thrown an exception');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('Error'), 'Should contain error message');
        }
        Test.stopTest();
    }

    /**
     * @description Test createFolder success in root
     */
    @isTest
    static void testCreateFolderSuccess() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockFolderCreationSuccess());

        Test.startTest();
        String folderId = SharePointDocumentController.createFolder('NewFolder', 'root');
        Test.stopTest();

        System.assertNotEquals(null, folderId, 'Folder ID should not be null');
        System.assertEquals('CREATED_FOLDER_ID', folderId, 'Should return correct folder ID');
    }

    /**
     * @description Test createFolder in specific parent folder
     */
    @isTest
    static void testCreateFolderInParent() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockFolderCreationSuccess());

        Test.startTest();
        String folderId = SharePointDocumentController.createFolder('SubFolder', '01PARENT123');
        Test.stopTest();

        System.assertNotEquals(null, folderId, 'Folder ID should not be null');
    }

    /**
     * @description Test createFolder error
     */
    @isTest
    static void testCreateFolderError() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockHttpResponseError());

        Test.startTest();
        try {
            SharePointDocumentController.createFolder('NewFolder', 'root');
            System.assert(false, 'Should have thrown an exception');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('Error'), 'Should contain error message');
        }
        Test.stopTest();
    }

    /**
     * @description Test getItemById success for file
     */
    @isTest
    static void testGetItemByIdFile() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockGetItemSuccess());

        Test.startTest();
        SharePointDocumentController.DocumentInfo item =
            SharePointDocumentController.getItemById('01FILE123');
        Test.stopTest();

        System.assertNotEquals(null, item, 'Item should not be null');
        System.assertEquals('TestDocument.pdf', item.name, 'Should return correct name');
        System.assertEquals(false, item.isFolder, 'Should be a file, not a folder');
    }

    /**
     * @description Test getItemById success for folder
     */
    @isTest
    static void testGetItemByIdFolder() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockGetItemFolderSuccess());

        Test.startTest();
        SharePointDocumentController.DocumentInfo item =
            SharePointDocumentController.getItemById('01FOLDER123');
        Test.stopTest();

        System.assertNotEquals(null, item, 'Item should not be null');
        System.assertEquals(true, item.isFolder, 'Should be a folder');
    }

    /**
     * @description Test getItemById error
     */
    @isTest
    static void testGetItemByIdError() {
        Test.setMock(HttpCalloutMock.class, new SharePointMockHttpResponseError());

        Test.startTest();
        try {
            SharePointDocumentController.getItemById('01FILE123');
            System.assert(false, 'Should have thrown an exception');
        } catch (AuraHandledException e) {
            System.assert(e.getMessage().contains('Error'), 'Should contain error message');
        }
        Test.stopTest();
    }

    /**
     * Mock HTTP Response - Success for search and recent documents
     */
    private class SharePointMockHttpResponseSuccess implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');

            // Check if it's a token request
            if(req.getEndpoint().contains('oauth2/v2.0/token')) {
                res.setStatusCode(200);
                res.setBody('{"access_token": "mock_access_token_12345", "token_type": "Bearer", "expires_in": 3599}');
                return res;
            }

            // Otherwise return documents
            res.setStatusCode(200);
            String responseBody = '{"value": [' +
                '{' +
                    '"id": "01ABCDEF123456789",' +
                    '"name": "TestDocument.pdf",' +
                    '"webUrl": "https://sharepoint.com/test.pdf",' +
                    '"lastModifiedDateTime": "2024-01-15T10:30:00Z",' +
                    '"size": 1024000,' +
                    '"file": {' +
                        '"mimeType": "application/pdf"' +
                    '},' +
                    '"lastModifiedBy": {' +
                        '"user": {' +
                            '"displayName": "Test User"' +
                        '}' +
                    '},' +
                    '"parentReference": {' +
                        '"name": "Alerts",' +
                        '"path": "/drives/drive-id/root:/Alerts"' +
                    '}' +
                '},' +
                '{' +
                    '"id": "02GHIJKL987654321",' +
                    '"name": "TechnicalDoc.docx",' +
                    '"webUrl": "https://sharepoint.com/technical.docx",' +
                    '"lastModifiedDateTime": "2024-01-14T09:15:00Z",' +
                    '"size": 2048000,' +
                    '"file": {' +
                        '"mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"' +
                    '},' +
                    '"lastModifiedBy": {' +
                        '"user": {' +
                            '"displayName": "Admin User"' +
                        '}' +
                    '},' +
                    '"parentReference": {' +
                        '"name": "Technical",' +
                        '"path": "/drives/drive-id/root:/Technical"' +
                    '}' +
                '}' +
            ']}';

            res.setBody(responseBody);
            return res;
        }
    }

    /**
     * Mock HTTP Response - Success for folders and files
     */
    private class SharePointMockFoldersAndFilesSuccess implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');

            // Check if it's a token request
            if(req.getEndpoint().contains('oauth2/v2.0/token')) {
                res.setStatusCode(200);
                res.setBody('{"access_token": "mock_access_token_12345", "token_type": "Bearer", "expires_in": 3599}');
                return res;
            }

            // Return folders and files
            res.setStatusCode(200);
            String responseBody = '{"value": [' +
                '{' +
                    '"id": "01FOLDER123",' +
                    '"name": "Documents",' +
                    '"webUrl": "https://sharepoint.com/documents",' +
                    '"lastModifiedDateTime": "2024-01-15T10:30:00Z",' +
                    '"folder": {' +
                        '"childCount": 5' +
                    '},' +
                    '"parentReference": {' +
                        '"name": "root"' +
                    '}' +
                '},' +
                '{' +
                    '"id": "01FILE123",' +
                    '"name": "Report.pdf",' +
                    '"webUrl": "https://sharepoint.com/report.pdf",' +
                    '"lastModifiedDateTime": "2024-01-14T09:15:00Z",' +
                    '"size": 1024000,' +
                    '"file": {' +
                        '"mimeType": "application/pdf"' +
                    '},' +
                    '"lastModifiedBy": {' +
                        '"user": {' +
                            '"displayName": "Test User"' +
                        '}' +
                    '},' +
                    '"parentReference": {' +
                        '"name": "Reports"' +
                    '}' +
                '}' +
            ']}';

            res.setBody(responseBody);
            return res;
        }
    }

    /**
     * Mock HTTP Response - Upload Success
     */
    private class SharePointMockUploadSuccess implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');

            // Check if it's a token request
            if(req.getEndpoint().contains('oauth2/v2.0/token')) {
                res.setStatusCode(200);
                res.setBody('{"access_token": "mock_access_token_12345", "token_type": "Bearer", "expires_in": 3599}');
                return res;
            }

            // Return upload success
            res.setStatusCode(201);
            String responseBody = '{' +
                '"id": "UPLOADED_FILE_ID",' +
                '"name": "TestFile.txt",' +
                '"size": 100,' +
                '"webUrl": "https://sharepoint.com/testfile.txt"' +
            '}';

            res.setBody(responseBody);
            return res;
        }
    }

    /**
     * Mock HTTP Response - Folder Creation Success
     */
    private class SharePointMockFolderCreationSuccess implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');

            // Check if it's a token request
            if(req.getEndpoint().contains('oauth2/v2.0/token')) {
                res.setStatusCode(200);
                res.setBody('{"access_token": "mock_access_token_12345", "token_type": "Bearer", "expires_in": 3599}');
                return res;
            }

            // Return folder creation success
            res.setStatusCode(201);
            String responseBody = '{' +
                '"id": "CREATED_FOLDER_ID",' +
                '"name": "NewFolder",' +
                '"folder": {' +
                    '"childCount": 0' +
                '},' +
                '"webUrl": "https://sharepoint.com/newfolder"' +
            '}';

            res.setBody(responseBody);
            return res;
        }
    }

    /**
     * Mock HTTP Response - Get Item Success (File)
     */
    private class SharePointMockGetItemSuccess implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');

            // Check if it's a token request
            if(req.getEndpoint().contains('oauth2/v2.0/token')) {
                res.setStatusCode(200);
                res.setBody('{"access_token": "mock_access_token_12345", "token_type": "Bearer", "expires_in": 3599}');
                return res;
            }

            // Return item details (file)
            res.setStatusCode(200);
            String responseBody = '{' +
                '"id": "01FILE123",' +
                '"name": "TestDocument.pdf",' +
                '"webUrl": "https://sharepoint.com/test.pdf",' +
                '"lastModifiedDateTime": "2024-01-15T10:30:00Z",' +
                '"size": 1024000,' +
                '"file": {' +
                    '"mimeType": "application/pdf"' +
                '}' +
            '}';

            res.setBody(responseBody);
            return res;
        }
    }

    /**
     * Mock HTTP Response - Get Item Success (Folder)
     */
    private class SharePointMockGetItemFolderSuccess implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');

            // Check if it's a token request
            if(req.getEndpoint().contains('oauth2/v2.0/token')) {
                res.setStatusCode(200);
                res.setBody('{"access_token": "mock_access_token_12345", "token_type": "Bearer", "expires_in": 3599}');
                return res;
            }

            // Return item details (folder)
            res.setStatusCode(200);
            String responseBody = '{' +
                '"id": "01FOLDER123",' +
                '"name": "Documents",' +
                '"webUrl": "https://sharepoint.com/documents",' +
                '"lastModifiedDateTime": "2024-01-15T10:30:00Z",' +
                '"folder": {' +
                    '"childCount": 10' +
                '}' +
            '}';

            res.setBody(responseBody);
            return res;
        }
    }

    /**
     * Mock HTTP Response - Error
     */
    private class SharePointMockHttpResponseError implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');

            // Check if it's a token request - return error for token too
            if(req.getEndpoint().contains('oauth2/v2.0/token')) {
                res.setStatusCode(401);
                res.setStatus('Unauthorized');
                res.setBody('{"error": {"code": "Unauthorized", "message": "Authentication failed"}}');
                return res;
            }

            res.setStatusCode(401);
            res.setStatus('Unauthorized');
            res.setBody('{"error": {"code": "Unauthorized", "message": "Authentication failed"}}');
            return res;
        }
    }
}
