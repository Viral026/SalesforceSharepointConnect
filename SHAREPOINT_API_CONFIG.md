# SharePoint API Configuration - Quick Start Guide

This guide walks you through configuring SharePoint/Microsoft Graph API integration with Salesforce from scratch.

---

## 🎯 Overview

You'll set up:
1. **Azure AD App** (for authentication)
2. **SharePoint API Endpoints** (Microsoft Graph)
3. **Salesforce Auth Provider** (OAuth 2.0)
4. **Salesforce Named Credential** (API callouts)

---

## 📋 Prerequisites Checklist

- [ ] Microsoft 365 admin account
- [ ] Azure AD access
- [ ] Salesforce org with Setup permissions
- [ ] SharePoint site with documents

---

## PART 1: Azure AD App Registration (15 minutes)

### Step 1: Create App Registration

1. Go to **https://portal.azure.com**
2. Navigate: **Azure Active Directory → App registrations**
3. Click **+ New registration**
4. Fill in:
   ```
   Name: Salesforce SharePoint Integration
   Supported account types: Single tenant
   Redirect URI: (leave blank for now)
   ```
5. Click **Register**

### Step 2: Capture Credentials

After registration, copy these values (you'll need them later):

| Setting | Location | Example |
|---------|----------|---------|
| **Application (client) ID** | Overview page | `12345678-1234-1234-1234-123456789abc` |
| **Directory (tenant) ID** | Overview page | `87654321-4321-4321-4321-cba987654321` |

### Step 3: Create Client Secret

1. In your app, go to **Certificates & secrets**
2. Click **+ New client secret**
3. Description: `Salesforce Integration`
4. Expires: **24 months** (recommended)
5. Click **Add**
6. **⚠️ IMPORTANT**: Copy the **Value** immediately (not the Secret ID)
   ```
   Example: AbC123~XyZ789-MnOpQrStUvWxYz
   ```
7. Store securely - you cannot retrieve this later!

### Step 4: Grant API Permissions

1. Go to **API permissions** in your app
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions** (not Delegated)
5. Search and add these permissions:
   - ✅ `Sites.Read.All`
   - ✅ `Files.Read.All`
6. Click **Add permissions**
7. 🔑 Click **Grant admin consent for [Your Organization]**
8. Confirm - status should show ✅ green checkmarks

---

## PART 2: Get SharePoint API Endpoints (10 minutes)

### Step 1: Find Your SharePoint Site URL

1. Open your SharePoint site in browser
2. Copy the URL, for example:
   ```
   https://contoso.sharepoint.com/sites/CompanyDocs
   ```
3. Break it down:
   - Hostname: `contoso.sharepoint.com`
   - Site path: `sites/CompanyDocs`

### Step 2: Get Site ID via Microsoft Graph

Use **Graph Explorer**: https://developer.microsoft.com/graph/graph-explorer

1. Sign in with your M365 account
2. Run this query (replace with your values):
   ```
   GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/CompanyDocs
   ```
3. Copy the **id** from response:
   ```json
   {
     "id": "contoso.sharepoint.com,abc12345-1234-abcd-5678-def987654321,xyz98765-4321-zyxw-8765-fed123456789"
   }
   ```
4. **Save this Site ID** - you'll need it for Named Credential

### Step 3: Get Drive ID (Document Library)

1. In Graph Explorer, run (use your Site ID):
   ```
   GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
   ```
2. Find your document library in the response:
   ```json
   {
     "value": [
       {
         "id": "b!abc123...xyz789",
         "name": "Documents",
         "driveType": "documentLibrary"
       }
     ]
   }
   ```
3. **Save the Drive ID** (the `id` field) - you'll update the Apex code with this

---

## PART 3: Configure Salesforce (20 minutes)

### Step 1: Create Authentication Provider

1. In Salesforce: **Setup → Auth. Providers**
2. Click **New**
3. Select **Provider Type: Microsoft**
4. Fill in the form:

```
Provider Type: Microsoft
Name: SharePoint_Auth
URL Suffix: SharePoint_Auth

Consumer Key: [Your Client ID from Azure]
Consumer Secret: [Your Client Secret from Azure]

Authorize Endpoint URL:
https://login.microsoftonline.com/{YOUR-TENANT-ID}/oauth2/v2.0/authorize

Token Endpoint URL:
https://login.microsoftonline.com/{YOUR-TENANT-ID}/oauth2/v2.0/token

Default Scopes:
https://graph.microsoft.com/.default offline_access

Include Consumer Secret in API Responses: ✅ (checked)
```

5. Click **Save**
6. **Copy the Callback URL** shown on the saved page (looks like):
   ```
   https://yourinstance.salesforce.com/services/authcallback/SharePoint_Auth
   ```

### Step 2: Update Azure AD Redirect URI

1. Go back to **Azure Portal → Your App → Authentication**
2. Click **+ Add a platform**
3. Select **Web**
4. Paste the **Callback URL** from Salesforce
5. Click **Configure**

### Step 3: Create Named Credential

1. In Salesforce: **Setup → Named Credentials → Named Credentials** (the new version)
2. Click **New**
3. Select **Named Credential** (not Legacy)
4. Configure:

#### Named Credential Details
```
Label: SharePoint Integration
Name: SharePoint_Named_Credential
URL: https://graph.microsoft.com/v1.0
```

#### External Credentials
- Click **New** to create new external credential
  ```
  Name: SharePoint_External_Credential
  Authentication Protocol: OAuth 2.0
  Authentication Provider: SharePoint_Auth
  Scope: https://graph.microsoft.com/.default offline_access
  ```

#### Permission Set Mapping
- Create a permission set mapping for your users/profile
- Identity Type: **Named Principal** or **Per User**

5. **Save**

> **Note**: If using Legacy Named Credentials, configure as:
> ```
> Label: SharePoint Integration
> Name: SharePoint_Named_Credential
> URL: https://graph.microsoft.com/v1.0
> Identity Type: Named Principal
> Authentication Protocol: OAuth 2.0
> Authentication Provider: SharePoint_Auth
> Scope: https://graph.microsoft.com/.default
> Generate Authorization Header: ✅
> Allow Merge Fields in HTTP Header: ✅
> Allow Merge Fields in HTTP Body: ✅
> ```

### Step 4: Update Apex Controller

Open [SharePointDocumentController.cls](force-app/main/default/classes/SharePointDocumentController.cls) and replace `{drive-id}` with your actual Drive ID:

**Find all occurrences:**
```apex
'/drives/{drive-id}/root:/'
'/drives/{drive-id}/root/children'
```

**Replace with:**
```apex
'/drives/b!abc123...xyz789/root:/'
'/drives/b!abc123...xyz789/root/children'
```

---

## PART 4: Test the Integration (5 minutes)

### Test 1: Authenticate Named Credential

1. **Setup → Named Credentials**
2. Find your **SharePoint_Named_Credential**
3. If using new Named Credentials:
   - Go to **External Credentials → SharePoint_External_Credential**
   - Click **Authenticate** next to your permission set mapping
   - Sign in with Microsoft account
   - Consent to permissions
4. Should show **Authenticated** status

### Test 2: Test API Call (Developer Console)

1. Open **Developer Console** (Setup → Developer Console)
2. Click **Debug → Open Execute Anonymous Window**
3. Run this test code:

```apex
HttpRequest req = new HttpRequest();
req.setEndpoint('callout:SharePoint_Named_Credential/sites/{YOUR-SITE-ID}/drives');
req.setMethod('GET');
req.setHeader('Content-Type', 'application/json');

Http http = new Http();
HttpResponse res = http.send(req);

System.debug('Status: ' + res.getStatusCode());
System.debug('Response: ' + res.getBody());
```

4. Check logs - should see **Status: 200** and JSON response with drives

### Test 3: Deploy and Test LWC

1. Deploy all components to Salesforce
2. Add **knowledgeCenter** component to a Lightning page
3. Open the page
4. Try searching for documents
5. Verify documents load from SharePoint

---

## 🔧 Quick Reference: Key Configuration Values

| Item | Where to Find | Where to Use |
|------|---------------|--------------|
| **Client ID** | Azure → App Overview | Salesforce Auth Provider (Consumer Key) |
| **Client Secret** | Azure → Certificates & secrets | Salesforce Auth Provider (Consumer Secret) |
| **Tenant ID** | Azure → App Overview | Auth endpoints URLs |
| **Site ID** | Graph Explorer API call | Named Credential URL (optional) |
| **Drive ID** | Graph Explorer API call | Apex Controller code |
| **Callback URL** | Saved Auth Provider page | Azure → App → Authentication |

---

## 🎯 Complete Endpoint Examples

### Named Credential Setup
```
URL: https://graph.microsoft.com/v1.0
```

### Apex Callouts
```apex
// Search documents
'callout:SharePoint_Named_Credential/search/query'

// Get documents from folder
'callout:SharePoint_Named_Credential/drives/{drive-id}/root:/{folder-path}:/children'

// Get recent documents
'callout:SharePoint_Named_Credential/drives/{drive-id}/root/children?$orderby=lastModifiedDateTime desc&$top=10'

// Get specific site info
'callout:SharePoint_Named_Credential/sites/{site-id}'
```

---

## ❌ Troubleshooting Common Issues

### Error: "INVALID_SESSION_ID" or "Unauthorized"
**Solution:**
- Authenticate the Named Credential (Step 4 → Test 1)
- Verify OAuth scopes include `offline_access`
- Check Client Secret hasn't expired

### Error: "Remote endpoint not allowed"
**Solution:**
- Ensure Named Credential URL is correct
- Don't need Remote Site Settings when using Named Credentials

### Error: "Access Denied" or 403
**Solution:**
- Grant admin consent in Azure AD (Part 1, Step 4)
- Verify permissions: Sites.Read.All and Files.Read.All
- Check they're Application permissions (not Delegated)

### No Documents Returned
**Solution:**
- Verify Drive ID is correct (use Graph Explorer)
- Check folder paths match your SharePoint structure
- Ensure documents exist in SharePoint
- Test API endpoint directly in Graph Explorer

### Auth Provider Error: "Invalid redirect_uri"
**Solution:**
- Copy exact Callback URL from Salesforce Auth Provider
- Paste in Azure → App → Authentication → Web platform
- Must match exactly (including https://)

---

## 📚 API Documentation Links

- **Microsoft Graph API**: https://learn.microsoft.com/graph/api/overview
- **Graph Explorer**: https://developer.microsoft.com/graph/graph-explorer
- **Salesforce Named Credentials**: https://help.salesforce.com/s/articleView?id=sf.named_credentials_about.htm
- **Azure AD App Setup**: https://learn.microsoft.com/azure/active-directory/develop/quickstart-register-app

---

## 🎓 Understanding the Flow

1. **User accesses LWC** → Component loads
2. **JavaScript calls Apex** → `searchDocuments()`
3. **Apex makes callout** → Via Named Credential
4. **Named Credential adds auth** → OAuth token from Auth Provider
5. **Microsoft Graph API** → Returns SharePoint data
6. **Apex parses response** → Returns to LWC
7. **LWC displays documents** → User sees results

---

## ✅ Post-Setup Checklist

After configuration, verify:

- [ ] Azure App has Client ID, Secret, and API permissions
- [ ] Admin consent granted in Azure AD
- [ ] Salesforce Auth Provider created with correct endpoints
- [ ] Azure App has Salesforce Callback URL in Redirect URIs
- [ ] Salesforce Named Credential created and authenticated
- [ ] Apex code updated with actual Drive ID
- [ ] Test API call returns HTTP 200
- [ ] LWC component deployed and accessible
- [ ] Documents load in the Knowledge Center

---

## 🔐 Security Best Practices

1. ✅ Use Named Principals for service accounts
2. ✅ Rotate Client Secrets every 12-24 months
3. ✅ Set calendar reminders before secret expiration
4. ✅ Grant only minimum required API permissions
5. ✅ Store credentials securely (use password manager)
6. ✅ Enable audit logs in both Azure and Salesforce
7. ✅ Restrict Auth Provider to specific profiles if needed

---

## 📞 Need Help?

1. Check the detailed [SETUP_GUIDE.md](SETUP_GUIDE.md) for more information
2. Review Salesforce Debug Logs (Setup → Debug Logs)
3. Check Azure AD Sign-in Logs (Azure → Azure AD → Sign-in logs)
4. Test endpoints in Graph Explorer first
5. Verify all IDs and secrets are correct

---

**Last Updated**: February 2026
**Version**: 1.0
