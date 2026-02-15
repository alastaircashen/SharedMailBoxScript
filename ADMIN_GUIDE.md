# Shared Mailbox Security Group Sync Tool - Admin Guide

This script creates and maintains security groups that mirror the membership of shared mailboxes. It can be run repeatedly to keep group membership in sync.

## Prerequisites

### 1. PowerShell Modules
The script will automatically install these if missing:
- `ExchangeOnlineManagement`
- `PnP.PowerShell`

### 2. Azure AD App Registration (Interactive Mode)

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

### 3. API Permissions (Interactive Mode)

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

### 5. Admin Roles Required (Interactive Mode)

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

### Exclusion List
Skip specific shared mailboxes by providing a CSV file with addresses to exclude:
```powershell
.\SharedMailboxManager.ps1 -Step Both -CsvPath "C:\temp\SharedMailboxes.csv" -ClientId "your-app-client-id" -ExclusionListPath "C:\temp\ExcludeMailboxes.csv"
```

The exclusion CSV should have a single column with a header row, e.g.:
```csv
MailboxAddress
shared-noreply@contoso.com
shared-archive@contoso.com
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
| `-ExclusionListPath` | No | - | CSV file of mailbox addresses to skip |
| `-Unattended` | No | Off | Enable certificate-based auth (no user interaction) |
| `-TenantDomain` | For Unattended | - | Tenant domain (e.g. `yourtenant.onmicrosoft.com`) |
| `-CertificateThumbprint` | For Unattended | - | Certificate thumbprint for app-only auth |

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

## Unattended Mode (Scheduled / Automated Runs)

Unattended mode uses certificate-based authentication so the script can run without user interaction, e.g. as a Windows Scheduled Task or in Azure Automation.

### Setup Overview

1. Create a self-signed certificate (or use an existing one)
2. Register an Azure AD app with **Application** permissions
3. Upload the certificate to the app registration
4. Assign the Exchange Administrator role to the app
5. Run the script with the `-Unattended` flag

### Step 1: Create a Self-Signed Certificate

Run this in PowerShell **as Administrator** on the machine that will run the script:

```powershell
# Create a self-signed certificate (valid for 2 years)
$cert = New-SelfSignedCertificate `
    -Subject "CN=SharedMailboxScript" `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -NotAfter (Get-Date).AddYears(2)

# Note the thumbprint - you will need this
Write-Host "Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green

# Export the public key (.cer) to upload to Azure AD
$cerPath = "C:\temp\SharedMailboxScript.cer"
Export-Certificate -Cert $cert -FilePath $cerPath
Write-Host "Public certificate exported to: $cerPath" -ForegroundColor Green
```

> **Note:** Save the thumbprint value. You will pass it to the script via `-CertificateThumbprint`.

### Step 2: Register the Azure AD App

1. Go to **Azure Portal** > **Microsoft Entra ID** > **App Registrations** > **New Registration**
2. Name: `SharedMailbox Script (Unattended)`
3. Supported account types: **Single tenant**
4. Click **Register**
5. Copy the **Application (client) ID** - this is your `-ClientId`

### Step 3: Upload the Certificate

1. In your app registration, go to **Certificates & secrets** > **Certificates**
2. Click **Upload certificate**
3. Select the `.cer` file exported in Step 1
4. Click **Add**

### Step 4: Add Application Permissions

In the app registration, go to **API Permissions** > **Add a permission**:

| API | Permission | Type | Purpose |
|-----|------------|------|---------|
| Microsoft Graph | `Group.ReadWrite.All` | **Application** | Create and manage security groups |
| Microsoft Graph | `User.Read.All` | **Application** | Look up users by email |
| Microsoft Graph | `GroupMember.ReadWrite.All` | **Application** | Add members to groups |
| Microsoft Graph | `Sites.FullControl.All` | **Application** | Update SharePoint list |

After adding all permissions, click **Grant admin consent for [your tenant]**.

> **Important:** Unattended mode requires **Application** permissions, not Delegated. This is different from interactive mode.

### Step 5: Assign Exchange Administrator Role

The app's service principal needs the Exchange Administrator role to read mailbox data:

1. Go to **Azure Portal** > **Microsoft Entra ID** > **Roles and administrators**
2. Search for and click **Exchange Administrator**
3. Click **Add assignments**
4. Click **No member selected** and search for your app name (`SharedMailbox Script (Unattended)`)
5. Select it and click **Add**

Alternatively, via PowerShell:
```powershell
# Get the service principal for the app
$sp = Get-MgServicePrincipal -Filter "displayName eq 'SharedMailbox Script (Unattended)'"

# Get the Exchange Administrator role
$role = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'"

# Assign the role
New-MgDirectoryRoleMember -DirectoryRoleId $role.Id -DirectoryObjectId $sp.Id
```

### Step 6: Run the Script in Unattended Mode

```powershell
.\SharedMailboxManager.ps1 -Step Both `
    -CsvPath "C:\temp\SharedMailboxes.csv" `
    -Unattended `
    -ClientId "your-app-client-id" `
    -TenantDomain "yourtenant.onmicrosoft.com" `
    -CertificateThumbprint "AB12CD34EF5678901234567890ABCDEF12345678"
```

You can combine unattended mode with other flags:
```powershell
# Unattended with exclusion list and test mode
.\SharedMailboxManager.ps1 -Step Both `
    -CsvPath "C:\temp\SharedMailboxes.csv" `
    -Unattended `
    -ClientId "your-app-client-id" `
    -TenantDomain "yourtenant.onmicrosoft.com" `
    -CertificateThumbprint "AB12CD34EF5678901234567890ABCDEF12345678" `
    -ExclusionListPath "C:\temp\ExcludeMailboxes.csv" `
    -TestMode
```

### Step 7: Schedule with Windows Task Scheduler (Optional)

1. Open **Task Scheduler** and click **Create Task**
2. **General** tab:
   - Name: `SharedMailbox Group Sync`
   - Select **Run whether user is logged on or not**
   - Check **Run with highest privileges**
   - Configure for your Windows version
3. **Triggers** tab:
   - Click **New**, set your desired schedule (e.g. daily at 2:00 AM)
4. **Actions** tab:
   - Click **New**
   - Program: `powershell.exe`
   - Arguments:
     ```
     -ExecutionPolicy Bypass -File "C:\Scripts\SharedMailboxManager.ps1" -Step Both -CsvPath "C:\temp\SharedMailboxes.csv" -Unattended -ClientId "your-client-id" -TenantDomain "yourtenant.onmicrosoft.com" -CertificateThumbprint "your-thumbprint"
     ```
5. **Settings** tab:
   - Check **Allow task to be run on demand**
   - Set **Stop the task if it runs longer than** to a reasonable value (e.g. 4 hours)
6. Click **OK** and enter the service account credentials

> **Note:** The service account running the scheduled task must have access to the certificate in `Cert:\LocalMachine\My`. Grant the account read access to the certificate's private key via the Certificates MMC snap-in.

---

## Troubleshooting

### "Insufficient privileges" Error
- Verify API permissions are added to your app registration
- Ensure admin consent has been granted
- Check the user has Exchange Administrator role
- For unattended mode: verify **Application** permissions (not Delegated) are used

### "Group not found" Warnings
- Newly created groups may take a few moments to sync across Azure AD

### SharePoint Errors
- Verify the SharePoint list exists with correct column names
- Check user has access to the SharePoint site
- For unattended mode: ensure `Sites.FullControl.All` Application permission is granted

### Unattended Mode: Certificate Errors
- Verify the certificate thumbprint matches exactly (no spaces)
- Ensure the certificate is installed in `Cert:\LocalMachine\My`
- Check the certificate has not expired
- Verify the `.cer` public key is uploaded to the Azure AD app registration
- If running as a scheduled task, ensure the service account has private key access

### Unattended Mode: Exchange Connection Fails
- Verify the Exchange Administrator role is assigned to the app's service principal
- Ensure the `-TenantDomain` value is correct (use `yourtenant.onmicrosoft.com` format)
- Check the `-ClientId` matches the app registration
