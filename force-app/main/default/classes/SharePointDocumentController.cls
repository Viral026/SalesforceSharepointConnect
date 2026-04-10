public with sharing class SharePointDocumentController {

    private static SharePointConnectionKeys__mdt cachedConfig;

    private static SharePointConnectionKeys__mdt getConfig() {
        if (cachedConfig == null) {
            Map<String, SharePointConnectionKeys__mdt> allConfigs = SharePointConnectionKeys__mdt.getAll();
            if (allConfigs.isEmpty()) {
                throw new AuraHandledException('SharePoint connection configuration not found. Please create a SharePointConnectionKeys custom metadata record.');
            }
            cachedConfig = allConfigs.values()[0];
        }
        return cachedConfig;
    }

    public class DocumentInfo {
        @AuraEnabled public String id;
        @AuraEnabled public String name;
        @AuraEnabled public String webUrl;
        @AuraEnabled public String fileType;
        @AuraEnabled public String lastModified;
        @AuraEnabled public String lastModifiedBy;
        @AuraEnabled public Long size;
        @AuraEnabled public String category;
        @AuraEnabled public Boolean isFolder;
    }

    public class FolderInfo {
        @AuraEnabled public String id;
        @AuraEnabled public String name;
        @AuraEnabled public String webUrl;
        @AuraEnabled public Integer childCount;
        @AuraEnabled public String lastModified;
        @AuraEnabled public Boolean isFolder;
    }

    // -----------------------------
    // GET ACCESS TOKEN
    // -----------------------------
    private static String getAccessToken(){

        HttpRequest req = new HttpRequest();

        req.setEndpoint(
        'https://login.microsoftonline.com/' +
        getConfig().Directory_tenant_ID__c +
        '/oauth2/v2.0/token'
        );

        req.setMethod('POST');

        req.setHeader(
        'Content-Type',
        'application/x-www-form-urlencoded'
        );

        String body =
        'client_id=' + EncodingUtil.urlEncode(getConfig().CLIENT_ID__c,'UTF-8') +
        '&client_secret=' + EncodingUtil.urlEncode(getConfig().CLIENT_SECRET__c,'UTF-8') +
        '&scope=' + EncodingUtil.urlEncode(
        'https://graph.microsoft.com/.default','UTF-8') +
        '&grant_type=client_credentials';

        req.setBody(body);

        Http http = new Http();
        HttpResponse res = http.send(req);

        if(res.getStatusCode()!=200){
            throw new AuraHandledException(
            'Token Error: '+res.getBody());
        }

        Map<String,Object> result =
        (Map<String,Object>)JSON.deserializeUntyped(res.getBody());

        return (String)result.get('access_token');
    }

    // -----------------------------
    // CALL GRAPH API
    // -----------------------------
    private static List<DocumentInfo> callGraphAPI(String endpoint){

        String token = getAccessToken();

        HttpRequest req = new HttpRequest();

        req.setEndpoint(getConfig().GRAPH_BASE__c + endpoint);

        req.setMethod('GET');

        req.setHeader(
        'Authorization',
        'Bearer ' + token
        );

        req.setHeader(
        'Content-Type',
        'application/json'
        );

        Http http = new Http();
        HttpResponse res = http.send(req);

        if(res.getStatusCode()!=200){

            throw new AuraHandledException(
            'SharePoint API Error: ' +
            res.getStatusCode() +
            ' - ' +
            res.getBody()
            );
        }

        return parseDocumentResponse(res.getBody());
    }

    // -----------------------------
    // GET FOLDERS AND FILES FROM ROOT
    // -----------------------------
    @AuraEnabled(cacheable=true)
    public static List<DocumentInfo> getFoldersAndFiles(String folderId){

        List<DocumentInfo> items = new List<DocumentInfo>();

        String endpoint;
        if(String.isBlank(folderId) || folderId == 'root'){
            endpoint = '/drives/' + getConfig().DRIVE_ID__c + '/root/children?$top=999';
        } else {
            endpoint = '/drives/' + getConfig().DRIVE_ID__c + '/items/' + folderId + '/children?$top=999';
        }

        List<Map<String,Object>> rawItems = getRawItemsFromEndpoint(endpoint);

        for(Map<String,Object> item : rawItems){
            DocumentInfo docInfo = new DocumentInfo();

            docInfo.id = (String)item.get('id');
            String rawName = (String)item.get('name');
            // Decode URL encoded characters (+ becomes space, %20 becomes space, etc.)
            docInfo.name = rawName != null ? EncodingUtil.urlDecode(rawName, 'UTF-8') : rawName;
            docInfo.webUrl = (String)item.get('webUrl');
            docInfo.lastModified = (String)item.get('lastModifiedDateTime');

            if(item.containsKey('folder')){
                // It's a folder
                docInfo.isFolder = true;
                Map<String,Object> folder = (Map<String,Object>)item.get('folder');
                if(folder.containsKey('childCount')){
                    Object childCountObj = folder.get('childCount');
                    if(childCountObj instanceof Integer){
                        Integer count = (Integer)childCountObj;
                        docInfo.size = count;
                    } else if(childCountObj instanceof Long){
                        docInfo.size = (Long)childCountObj;
                    }
                }
                docInfo.category = 'Folder';
            } else {
                // It's a file
                docInfo.isFolder = false;

                if(item.containsKey('size')){
                    docInfo.size = (Long)item.get('size');
                }

                if(item.containsKey('file')){
                    Map<String,Object> file = (Map<String,Object>)item.get('file');
                    docInfo.fileType = (String)file.get('mimeType');
                }

                // Parse last modified by
                if(item.containsKey('lastModifiedBy')){
                    Map<String,Object> modifiedBy = (Map<String,Object>)item.get('lastModifiedBy');
                    if(modifiedBy.containsKey('user')){
                        Map<String,Object> user = (Map<String,Object>)modifiedBy.get('user');
                        docInfo.lastModifiedBy = (String)user.get('displayName');
                    }
                }

                // Extract category from parent path
                if(item.containsKey('parentReference')){
                    Map<String,Object> parentRef = (Map<String,Object>)item.get('parentReference');
                    String parentName = (String)parentRef.get('name');
                    docInfo.category = String.isNotBlank(parentName) ? parentName : 'General';
                }
            }

            items.add(docInfo);
        }

        return items;
    }

    // -----------------------------
    // GET ALL DOCUMENTS (RECURSIVE)
    // -----------------------------
    @AuraEnabled(cacheable=true)
    public static List<DocumentInfo> getAllDocuments(){

        List<DocumentInfo> allDocs = new List<DocumentInfo>();

        String endpoint =
        '/drives/' + getConfig().DRIVE_ID__c + '/root/children?$top=999';

        // Get all items (files and folders) from root
        List<Map<String,Object>> items = getRawItemsFromEndpoint(endpoint);

        // Process each item
        for(Map<String,Object> item : items){
            if(item.containsKey('folder')){
                // It's a folder - get documents recursively
                String itemId = (String)item.get('id');
                allDocs.addAll(getDocumentsFromFolder(itemId));
            } else {
                // It's a file - add to results
                DocumentInfo doc = parseDocumentItem(item);
                allDocs.add(doc);
            }
        }

        return allDocs;
    }

    // -----------------------------
    // GET DOCUMENTS FROM FOLDER
    // -----------------------------
    private static List<DocumentInfo> getDocumentsFromFolder(String folderId){

        List<DocumentInfo> docs = new List<DocumentInfo>();

        String endpoint =
        '/drives/' + getConfig().DRIVE_ID__c + '/items/' + folderId + '/children?$top=999';

        List<Map<String,Object>> items = getRawItemsFromEndpoint(endpoint);

        for(Map<String,Object> item : items){
            if(item.containsKey('folder')){
                // Recurse into subfolder
                String subFolderId = (String)item.get('id');
                docs.addAll(getDocumentsFromFolder(subFolderId));
            } else {
                // Add file
                DocumentInfo doc = parseDocumentItem(item);
                docs.add(doc);
            }
        }

        return docs;
    }

    // -----------------------------
    // GET RAW ITEMS FROM ENDPOINT
    // -----------------------------
    private static List<Map<String,Object>> getRawItemsFromEndpoint(String endpoint){

        String token = getAccessToken();

        HttpRequest req = new HttpRequest();
        req.setEndpoint(getConfig().GRAPH_BASE__c + endpoint);
        req.setMethod('GET');
        req.setHeader('Authorization', 'Bearer ' + token);
        req.setHeader('Content-Type', 'application/json');

        Http http = new Http();
        HttpResponse res = http.send(req);

        List<Map<String,Object>> items = new List<Map<String,Object>>();

        if(res.getStatusCode() == 200){
            Map<String,Object> result =
            (Map<String,Object>)JSON.deserializeUntyped(res.getBody());

            if(result.containsKey('value')){
                List<Object> valueList = (List<Object>)result.get('value');
                for(Object obj : valueList){
                    items.add((Map<String,Object>)obj);
                }
            }
        }

        return items;
    }

    // -----------------------------
    // PARSE SINGLE DOCUMENT ITEM
    // -----------------------------
    private static DocumentInfo parseDocumentItem(Map<String,Object> item){

        DocumentInfo doc = new DocumentInfo();

        doc.id = (String)item.get('id');
        String rawName = (String)item.get('name');
        // Decode URL encoded characters (+ becomes space, %20 becomes space, etc.)
        doc.name = rawName != null ? EncodingUtil.urlDecode(rawName, 'UTF-8') : rawName;
        doc.webUrl = (String)item.get('webUrl');
        doc.lastModified = (String)item.get('lastModifiedDateTime');

        if(item.containsKey('size')){
            doc.size = (Long)item.get('size');
        }

        if(item.containsKey('file')){
            Map<String,Object> file = (Map<String,Object>)item.get('file');
            doc.fileType = (String)file.get('mimeType');
        }

        // Parse last modified by
        if(item.containsKey('lastModifiedBy')){
            Map<String,Object> modifiedBy = (Map<String,Object>)item.get('lastModifiedBy');
            if(modifiedBy.containsKey('user')){
                Map<String,Object> user = (Map<String,Object>)modifiedBy.get('user');
                doc.lastModifiedBy = (String)user.get('displayName');
            }
        }

        // Extract category from parent path
        if(item.containsKey('parentReference')){
            Map<String,Object> parentRef = (Map<String,Object>)item.get('parentReference');
            String parentName = (String)parentRef.get('name');
            doc.category = String.isNotBlank(parentName) ? parentName : 'General';
        }

        return doc;
    }

    // -----------------------------
    // GET RECENT DOCUMENTS
    // -----------------------------
    @AuraEnabled(cacheable=true)
    public static List<DocumentInfo> getRecentDocuments(Integer limitCount){

        if(limitCount==null){
            limitCount=10;
        }

        String endpoint =
        '/drives/' + getConfig().DRIVE_ID__c +
        '/root/children?$orderby=lastModifiedDateTime desc&$top=' +
        limitCount;

        return callGraphAPI(endpoint);
    }

    // -----------------------------
    // SEARCH DOCUMENTS
    // -----------------------------
    @AuraEnabled(cacheable=true)
    public static List<DocumentInfo> searchDocuments(String keyword){

        if(String.isBlank(keyword)){
            return new List<DocumentInfo>();
        }

        String endpoint =
        '/drives/' + getConfig().DRIVE_ID__c +
        '/root/search(q=\'' +
        EncodingUtil.urlEncode(keyword,'UTF-8') +
        '\')';

        return callGraphAPI(endpoint);
    }

    // -----------------------------
    // UPLOAD DOCUMENT TO SHAREPOINT
    // -----------------------------
    @AuraEnabled
    public static String uploadDocument(String fileName, String fileContent, String folderId){

        try {
            String token = getAccessToken();

            // Build endpoint based on folder
            String endpoint;
            if(String.isBlank(folderId) || folderId == 'root'){
                // Upload to root
                endpoint = getConfig().GRAPH_BASE__c + '/drives/' + getConfig().DRIVE_ID__c + '/root:/' +
                           EncodingUtil.urlEncode(fileName, 'UTF-8') + ':/content';
            } else {
                // Upload to specific folder
                endpoint = getConfig().GRAPH_BASE__c + '/drives/' + getConfig().DRIVE_ID__c + '/items/' +
                           folderId + ':/' + EncodingUtil.urlEncode(fileName, 'UTF-8') + ':/content';
            }

            // Decode base64 content
            Blob fileBlob = EncodingUtil.base64Decode(fileContent);

            // Create HTTP request
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('PUT');
            req.setHeader('Authorization', 'Bearer ' + token);
            req.setHeader('Content-Type', 'application/octet-stream');
            req.setBodyAsBlob(fileBlob);
            req.setTimeout(120000);

            Http http = new Http();
            HttpResponse res = http.send(req);

            if(res.getStatusCode() == 200 || res.getStatusCode() == 201){
                Map<String, Object> result =
                    (Map<String, Object>)JSON.deserializeUntyped(res.getBody());
                return (String)result.get('id'); // Return the uploaded file ID
            } else {
                throw new AuraHandledException(
                    'Upload failed: ' + res.getStatusCode() + ' - ' + res.getBody()
                );
            }

        } catch(Exception e){
            throw new AuraHandledException('Error uploading document: ' + e.getMessage());
        }
    }

    // -----------------------------
    // CREATE FOLDER IN SHAREPOINT
    // -----------------------------
    @AuraEnabled
    public static String createFolder(String folderName, String parentFolderId){

        try {
            String token = getAccessToken();

            // Build endpoint based on parent folder
            String endpoint;
            if(String.isBlank(parentFolderId) || parentFolderId == 'root'){
                // Create in root
                endpoint = getConfig().GRAPH_BASE__c + '/drives/' + getConfig().DRIVE_ID__c + '/root/children';
            } else {
                // Create in specific folder
                endpoint = getConfig().GRAPH_BASE__c + '/drives/' + getConfig().DRIVE_ID__c + '/items/' + parentFolderId + '/children';
            }

            // Build request body for folder creation
            Map<String, Object> requestBody = new Map<String, Object>{
                'name' => folderName,
                'folder' => new Map<String, Object>(),
                '@microsoft.graph.conflictBehavior' => 'rename'
            };

            // Create HTTP request
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('POST');
            req.setHeader('Authorization', 'Bearer ' + token);
            req.setHeader('Content-Type', 'application/json');
            req.setBody(JSON.serialize(requestBody));
            req.setTimeout(120000);

            Http http = new Http();
            HttpResponse res = http.send(req);

            if(res.getStatusCode() == 200 || res.getStatusCode() == 201){
                Map<String, Object> result =
                    (Map<String, Object>)JSON.deserializeUntyped(res.getBody());
                return (String)result.get('id'); // Return the created folder ID
            } else {
                throw new AuraHandledException(
                    'Folder creation failed: ' + res.getStatusCode() + ' - ' + res.getBody()
                );
            }

        } catch(Exception e){
            throw new AuraHandledException('Error creating folder: ' + e.getMessage());
        }
    }

    // -----------------------------
    // GET SINGLE ITEM BY ID
    // -----------------------------
    @AuraEnabled
    public static DocumentInfo getItemById(String itemId){

        try {
            String token = getAccessToken();

            // Build endpoint for specific item
            String endpoint = getConfig().GRAPH_BASE__c + '/drives/' + getConfig().DRIVE_ID__c + '/items/' + itemId;

            // Create HTTP request
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', 'Bearer ' + token);
            req.setHeader('Content-Type', 'application/json');
            req.setTimeout(120000);

            Http http = new Http();
            HttpResponse res = http.send(req);

            if(res.getStatusCode() == 200){
                Map<String, Object> itemData =
                    (Map<String, Object>)JSON.deserializeUntyped(res.getBody());

                DocumentInfo docInfo = new DocumentInfo();
                docInfo.id = (String)itemData.get('id');
                String rawName = (String)itemData.get('name');
                // Decode URL encoded characters (+ becomes space, %20 becomes space, etc.)
                docInfo.name = rawName != null ? EncodingUtil.urlDecode(rawName, 'UTF-8') : rawName;
                docInfo.webUrl = (String)itemData.get('webUrl');
                docInfo.lastModified = (String)itemData.get('lastModifiedDateTime');

                // Check if it's a folder
                docInfo.isFolder = itemData.containsKey('folder');

                if(docInfo.isFolder){
                    Map<String, Object> folderInfo =
                        (Map<String, Object>)itemData.get('folder');
                    Object childCountObj = folderInfo.get('childCount');
                    if(childCountObj != null){
                        Integer count = (Integer)childCountObj;
                        docInfo.size = count;
                    } else {
                        docInfo.size = 0;
                    }
                } else {
                    // It's a file
                    Object sizeObj = itemData.get('size');
                    if(sizeObj != null){
                        docInfo.size = (Integer)sizeObj;
                    }
                }

                return docInfo;
            } else {
                throw new AuraHandledException(
                    'Failed to get item: ' + res.getStatusCode() + ' - ' + res.getBody()
                );
            }

        } catch(Exception e){
            throw new AuraHandledException('Error getting item: ' + e.getMessage());
        }
    }

    // -----------------------------
    // PARSE RESPONSE
    // -----------------------------
    private static List<DocumentInfo>
    parseDocumentResponse(String responseBody){

        List<DocumentInfo> documents =
        new List<DocumentInfo>();

        Map<String,Object> result =
        (Map<String,Object>)
        JSON.deserializeUntyped(responseBody);

        if(!result.containsKey('value')){
            return documents;
        }

        List<Object> items =
        (List<Object>)result.get('value');

        for(Object obj:items){

            Map<String,Object> item =
            (Map<String,Object>)obj;

            DocumentInfo doc =
            new DocumentInfo();

            doc.id =
            (String)item.get('id');

            String rawName =
            (String)item.get('name');

            // Decode URL encoded characters (+ becomes space, %20 becomes space, etc.)
            doc.name =
            rawName != null ? EncodingUtil.urlDecode(rawName, 'UTF-8') : rawName;

            doc.lastModified =
            (String)item.get('lastModifiedDateTime');

            // Check if it's a folder
            if(item.containsKey('folder')){
                doc.isFolder = true;
                Map<String,Object> folderInfo =
                (Map<String,Object>)item.get('folder');

                // Get child count for folders
                Object childCountObj = folderInfo.get('childCount');
                if(childCountObj != null){
                    Integer count = (Integer)childCountObj;
                    doc.size = count;
                } else {
                    doc.size = 0;
                }
            } else {
                // It's a file
                doc.isFolder = false;

                if(item.containsKey('size')){
                    doc.size =
                    (Long)item.get('size');
                }

                if(item.containsKey('file')){
                    Map<String,Object> file =
                    (Map<String,Object>)item.get('file');

                    doc.fileType =
                    (String)file.get('mimeType');
                }
            }

            documents.add(doc);
        }

        return documents;
    }
}