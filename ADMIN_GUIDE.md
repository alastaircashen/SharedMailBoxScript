# Shared Mailbox Security Group Sync Tool - Admin Guide

This script creates and maintains security groups that mirror the membership of shared mailboxes. It can be run repeatedly to keep group membership in sync.

## Prerequisites

### 1. PowerShell Modules
The script will automatically install these if missing:
- `ExchangeOnlineManagement`
- `PnP.PowerShell`

### 2. Azure AD App Registration

You need an Azure AD app registration for PnP PowerShell to access Microsoft Graph and SharePoint.

#### Option A: Auto-Register (Recommended)
Run this command once in PowerShell:
```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP SharedMailbox Script" -Tenant yourtenant.onmicrosoft.com -Interactive
```

#### Option B: Manual Registration
1. Go to **Azure Portal** > **App Registrations** > **New Registration**
2. Name: `PnP SharedMailbox Script`
3. Supported account types: Single tenant
4. Redirect URI: `http://localhost` (Web)

### 3. API Permissions

Your app registration needs these **Delegated** permissions:

| API | Permission | Type | Purpose |
|-----|------------|------|---------|
| Microsoft Graph | `Group.ReadWrite.All` | Delegated | Create and manage security groups |
| Microsoft Graph | `User.Read.All` | Delegated | Look up users by email |
| Microsoft Graph | `GroupMember.ReadWrite.All` | Delegated | Add members to groups |
| SharePoint | `Sites.FullControl.All` | Delegated | Update SharePoint list |

**Important:** After adding permissions, click **Grant admin consent** for your tenant.

### 4. SharePoint List

Create a SharePoint list named `Shared Mailboxes Mapping` with these columns:
- `SharedMailbox` (Single line of text)
- `SecurityGroup` (Single line of text)

### 5. Admin Roles Required

The user running the script needs:
- **Exchange Administrator** or **Global Administrator** (for Exchange Online access)
- **SharePoint site access** to the target site

---

## Usage

### Step 1: Export Shared Mailbox Data
```powershell
.\SharedMailboxManager.ps1 -Step Export -CsvPath "C:\temp\SharedMailboxes.csv"
```
This exports all shared mailboxes and their current permissions to a CSV file.

### Step 2: Import and Sync Groups
```powershell
.\SharedMailboxManager.ps1 -Step Import -CsvPath "C:\temp\SharedMailboxes.csv" -ClientId "your-app-client-id"
```
This creates security groups and syncs membership to match mailbox permissions.

### Run Both Steps Together
```powershell
.\SharedMailboxManager.ps1 -Step Both -CsvPath "C:\temp\SharedMailboxes.csv" -ClientId "your-app-client-id"
```

### Test Mode (First 5 Mailboxes Only)
```powershell
.\SharedMailboxManager.ps1 -Step Both -CsvPath "C:\temp\SharedMailboxes.csv" -ClientId "your-app-client-id" -TestMode
```

---

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-Step` | Yes | - | `Export`, `Import`, or `Both` |
| `-CsvPath` | Yes | - | Path to CSV file |
| `-ClientId` | For Import/Both | - | Azure AD App Client ID |
| `-SharePointSiteUrl` | No | (configured) | SharePoint site URL |
| `-SharePointListName` | No | `Shared Mailboxes Mapping` | SharePoint list name |
| `-TestMode` | No | Off | Process only first 5 mailboxes |

---

## What the Script Does

1. **Export**: Reads all shared mailboxes and their SendAs/SendOnBehalf permissions
2. **Import/Sync**:
   - Creates a security group named `smb-mailbox-{mailboxprefix}` for each mailbox
   - Adds users with mailbox permissions as group members
   - Compares existing group membership and adds missing users
   - Records the mailbox-to-group mapping in SharePoint

---

## Running Regularly

The script is idempotent - safe to run multiple times:
- Existing groups are reused (not duplicated)
- Only missing members are added to groups
- SharePoint mappings are only created if they don't exist

You can schedule this script to run periodically to keep groups in sync with mailbox permissions.

---

## Troubleshooting

### "Insufficient privileges" Error
- Verify API permissions are added to your app registration
- Ensure admin consent has been granted
- Check the user has Exchange Administrator role

### "Group not found" Warnings
- Newly created groups may take a few moments to sync across Azure AD

### SharePoint Errors
- Verify the SharePoint list exists with correct column names
- Check user has access to the SharePoint site
