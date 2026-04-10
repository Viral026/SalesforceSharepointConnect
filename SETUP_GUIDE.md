# SharePoint Knowledge Center - Setup Guide

## Overview
This guide will help you configure the SharePoint Knowledge Center integration with Salesforce. The integration allows users to search and access SharePoint documents directly from Salesforce.

## Prerequisites
- Salesforce Org with access to Setup
- Microsoft Azure AD account with admin privileges
- SharePoint Online site with document libraries
- Access to create Named Credentials in Salesforce

---

## Part 1: Azure AD App Registration

### Step 1: Register Application in Azure AD

1. Navigate to **Azure Portal** (https://portal.azure.com)
2. Go to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure the application:
   - **Name**: `Salesforce SharePoint Integration`
   - **Supported account types**: Select appropriate option (usually "Accounts in this organizational directory only")
   - **Redirect URI**: Leave blank for now (we'll add it later)
5. Click **Register**

### Step 2: Note Important IDs

After registration, note down the following:
- **Application (client) ID**: Found on the Overview page
- **Directory (tenant) ID**: Found on the Overview page

### Step 3: Create Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description: `Salesforce Integration Secret`
4. Select expiration period (recommended: 24 months)
5. Click **Add**
6. **IMPORTANT**: Copy the secret **Value** immediately (you won't be able to see it again)

### Step 4: Configure API Permissions

1. Go to **API permissions** in your app registration
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions**
5. Add the following permissions:
   - `Sites.Read.All` - Read items in all site collections
   - `Files.Read.All` - Read files in all site collections
6. Click **Add permissions**
7. Click **Grant admin consent for [Your Organization]**
8. Confirm the consent

---

## Part 2: SharePoint Configuration

### Step 1: Identify Your SharePoint Site

1. Navigate to your SharePoint site
2. Note the site URL (e.g., `https://yourtenant.sharepoint.com/sites/YourSite`)
3. Identify the Document Library name where your documents are stored

### Step 2: Organize Documents by Category

For best results, organize your SharePoint documents in folders:
```
Document Library/
├── Alerts/
│   └── (Alert documents)
├── Functional Docs/
│   └── (Functional documentation)
├── Technical Docs/
│   └── (Technical documentation)
├── SOPs/
│   └── (Standard Operating Procedures)
└── General/
    └── (Other documents)
```

### Step 3: Get Site ID and Drive ID

Run this in Microsoft Graph Explorer (https://developer.microsoft.com/en-us/graph/graph-explorer):

**Get Site ID:**
```
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-name}
```

Example:
```
GET https://graph.microsoft.com/v1.0/sites/yourtenant.sharepoint.com:/sites/YourSite
```

**Get Drive ID:**
```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
```

Note down the `id` values from the responses.

---

## Part 3: Salesforce Configuration

### Step 1: Create Named Credential

1. In Salesforce, go to **Setup**
2. Search for **Named Credentials** in Quick Find
3. Click **New Named Credential**
4. Configure as follows:

**Named Credential Settings:**
- **Label**: `SharePoint Integration`
- **Name**: `SharePoint_Named_Credential`
- **URL**: `https://graph.microsoft.com/v1.0/sites/{your-site-id}`
- **Identity Type**: Named Principal
- **Authentication Protocol**: OAuth 2.0
- **Authentication Provider**: (Create new - see below)
- **Scope**: `https://graph.microsoft.com/.default`
- **Generate Authorization Header**: Checked
- **Allow Merge Fields in HTTP Header**: Checked
- **Allow Merge Fields in HTTP Body**: Checked

### Step 2: Create Authentication Provider

Before saving the Named Credential, you need to create an Auth Provider:

1. Go to **Setup** > **Auth. Providers**
2. Click **New**
3. Select **Microsoft Azure AD** or **OAuth 2.0** provider
4. Configure:

**Auth Provider Settings:**
- **Provider Type**: Microsoft
- **Name**: `SharePoint_Auth`
- **URL Suffix**: `SharePoint_Auth`
- **Consumer Key**: {Your Azure Application (client) ID}
- **Consumer Secret**: {Your Azure Client Secret Value}
- **Authorize Endpoint URL**: `https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/authorize`
- **Token Endpoint URL**: `https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token`
- **Default Scopes**: `https://graph.microsoft.com/.default offline_access`

5. Click **Save**
6. Copy the **Callback URL** shown after saving

### Step 3: Update Azure AD Redirect URI

1. Go back to Azure Portal
2. Navigate to your App Registration
3. Go to **Authentication**
4. Click **Add a platform** > **Web**
5. Add the **Callback URL** from Salesforce Auth Provider
6. Click **Configure**

### Step 4: Update Apex Controller with Site/Drive IDs

1. Open **SharePointDocumentController.cls**
2. Find the placeholder `{drive-id}` in the code
3. Replace it with your actual Drive ID from Step 3 of SharePoint Configuration

Example:
```apex
// Before
String folderPath = '/drives/{drive-id}/root:/' + category + ':/children';

// After
String folderPath = '/drives/b!abc123xyz456/root:/' + category + ':/children';
```

### Step 5: Configure Remote Site Settings (If Needed)

If not using Named Credentials for all callouts:

1. Go to **Setup** > **Remote Site Settings**
2. Click **New Remote Site**
3. Configure:
   - **Remote Site Name**: `Microsoft_Graph`
   - **Remote Site URL**: `https://graph.microsoft.com`
   - **Active**: Checked
4. Click **Save**

---

## Part 4: Deploy and Configure LWC

### Step 1: Deploy All Components

Deploy the following files to your Salesforce org:

**Apex Classes:**
- `SharePointDocumentController.cls`
- `SharePointDocumentController.cls-meta.xml`
- `SharePointDocumentControllerTest.cls`
- `SharePointDocumentControllerTest.cls-meta.xml`

**LWC Components:**
- `knowledgeCenter/knowledgeCenter.html`
- `knowledgeCenter/knowledgeCenter.js`
- `knowledgeCenter/knowledgeCenter.css`
- `knowledgeCenter/knowledgeCenter.js-meta.xml`

### Step 2: Create Lightning App or Tab

**Option A: Create Custom Lightning Tab (Recommended)**

1. Go to **Setup** > **Tabs**
2. Click **New** in Lightning Component Tabs section
3. Select your `knowledgeCenter` component
4. Configure:
   - **Tab Label**: `Knowledge Hub`
   - **Tab Name**: `Knowledge_Hub`
   - **Tab Style**: Choose an icon (e.g., `Document`)
5. Click **Next** > **Save**
6. Add the tab to relevant Lightning Apps

**Option B: Add to Lightning App Builder**

1. Go to **Setup** > **Lightning App Builder**
2. Create or edit an app page
3. Drag the `knowledgeCenter` component onto the page
4. Configure height if needed
5. Save and activate

**Option C: Add to Record Page**

1. Go to any record (e.g., Account, Case)
2. Click gear icon > **Edit Page**
3. Drag `knowledgeCenter` component to desired location
4. Save and activate

---

## Part 5: Testing the Integration

### Step 1: Test Authentication

1. Navigate to the Knowledge Hub tab or page
2. The component should load without errors
3. Check browser console for any authentication issues

### Step 2: Test Document Search

1. Enter a search keyword in the search box
2. Click **Search**
3. Verify that documents are displayed
4. Test the **View** and **Download** buttons

### Step 3: Test Category Filters

1. Select different categories from the dropdown
2. Verify that results are filtered appropriately
3. Test the **Clear** button

### Step 4: Verify Recent Documents

1. When no search is performed, verify that recent documents are shown
2. Check that the last modified dates are correct

---

## Troubleshooting

### Common Issues and Solutions

#### Issue 1: "Authentication failed" error
**Solution:**
- Verify Client ID and Client Secret in Auth Provider
- Check that Admin Consent was granted in Azure AD
- Verify Token Endpoint URL includes correct Tenant ID

#### Issue 2: "Remote Site Settings" error
**Solution:**
- Ensure Named Credential is created with correct URL
- Verify authentication protocol is OAuth 2.0

#### Issue 3: No documents returned
**Solution:**
- Verify Site ID and Drive ID are correct
- Check that the Azure AD app has proper permissions
- Ensure documents exist in SharePoint
- Check SharePoint folder structure matches expected categories

#### Issue 4: "CORS" or "Mixed Content" errors
**Solution:**
- This shouldn't occur with proper Named Credential setup
- Ensure all URLs use HTTPS

#### Issue 5: Token expiration
**Solution:**
- Named Credentials handle token refresh automatically
- If issues persist, check Auth Provider configuration
- Verify "offline_access" scope is included

---

## Security Best Practices

1. **Never hardcode credentials** - Always use Named Credentials
2. **Rotate secrets regularly** - Update Azure Client Secret periodically
3. **Use principle of least privilege** - Only grant necessary permissions
4. **Monitor access logs** - Review who's accessing documents
5. **Implement field-level security** - Restrict component visibility as needed
6. **Enable audit trail** - Track document access in Salesforce

---

## Maintenance

### Regular Tasks

1. **Monthly**: Review Azure AD app permissions and access logs
2. **Quarterly**: Rotate Azure Client Secret (before expiration)
3. **As needed**: Update Named Credential if endpoints change
4. **As needed**: Update category mappings in Apex code

---

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/api/overview)
- [Salesforce Named Credentials Guide](https://help.salesforce.com/articleView?id=named_credentials_about.htm)
- [Azure AD App Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Lightning Web Components Developer Guide](https://developer.salesforce.com/docs/component-library/documentation/en/lwc)

---

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Review Salesforce and Azure AD logs
3. Contact your Salesforce administrator
4. Consult Microsoft Graph API documentation

---

## Version History

- **v1.0** (Current): Initial implementation with search, category filtering, and recent documents
- Future enhancements: AI-powered recommendations, bookmarking, advanced filters

---

## API Endpoints Reference

The integration uses these Microsoft Graph API endpoints:

1. **Search Documents**: `POST /search/query`
2. **Get Documents by Folder**: `GET /drives/{drive-id}/root:/{path}:/children`
3. **Get Recent Documents**: `GET /drives/{drive-id}/root/children?$orderby=lastModifiedDateTime desc`

All endpoints require Bearer token authentication handled by Named Credential.
