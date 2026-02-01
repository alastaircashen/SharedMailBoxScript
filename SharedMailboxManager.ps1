<#
.SYNOPSIS
    Manages M365 Shared Mailboxes - Exports mailbox data and creates security groups for permissions.

.DESCRIPTION
    This script provides two main functions:
    Step 1: Export all shared mailboxes and their send permissions to a CSV file
    Step 2: Import the CSV, create security groups, assign permissions, and update SharePoint

    Uses only ExchangeOnlineManagement and PnP.PowerShell modules.

    PREREQUISITE: You must register an Azure AD application for PnP PowerShell.
    Run this command once to register: Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell" -Tenant yourtenant.onmicrosoft.com -Interactive
    See: https://pnp.github.io/powershell/articles/registerapplication.html

.PARAMETER Step
    Specify which step to run: 'Export', 'Import', or 'Both'

.PARAMETER CsvPath
    Path to the CSV file (for export or import)

.PARAMETER ClientId
    The Azure AD Application (Client) ID for PnP PowerShell authentication.
    Required for Import and Both steps.

.PARAMETER SharePointSiteUrl
    SharePoint site URL for the list update

.EXAMPLE
    .\SharedMailboxManager.ps1 -Step Export -CsvPath "C:\temp\SharedMailboxes.csv"
    .\SharedMailboxManager.ps1 -Step Import -CsvPath "C:\temp\SharedMailboxes.csv" -ClientId "your-app-client-id"
    .\SharedMailboxManager.ps1 -Step Both -CsvPath "C:\temp\SharedMailboxes.csv" -ClientId "your-app-client-id"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Export', 'Import', 'Both')]
    [string]$Step,

    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "6042b520-90d0-4286-9192-fdbbe025740e",

    [Parameter(Mandatory = $false)]
    [string]$SharePointSiteUrl = "https://zn8r8.sharepoint.com/sites/DMData",

    [Parameter(Mandatory = $false)]
    [string]$SharePointListName = "SharedMailboxesMapping"
)

# Set execution policy to allow running unsigned scripts (current process only)
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Load System.Web assembly for URL encoding
Add-Type -AssemblyName System.Web

#region Module Installation and Connection Functions

function Install-RequiredModules {
    <#
    .SYNOPSIS
        Installs required PowerShell modules if not already installed.
    #>
    $modules = @(
        'ExchangeOnlineManagement',
        'PnP.PowerShell'
    )

    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing module: $module" -ForegroundColor Yellow
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
        else {
            Write-Host "Module already installed: $module" -ForegroundColor Green
        }
    }
}

function Connect-ExchangeOnlineService {
    <#
    .SYNOPSIS
        Connects to Exchange Online.
    #>
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    $connectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($connectionInfo) {
        Write-Host "Already connected to Exchange Online" -ForegroundColor Green
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false
        $connectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if (-not $connectionInfo) {
            throw "Failed to connect to Exchange Online."
        }
        Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
    }
}

function Connect-PnPService {
    <#
    .SYNOPSIS
        Connects to PnP PowerShell for Microsoft 365 operations.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$ClientId
    )

    Write-Host "Connecting to PnP PowerShell..." -ForegroundColor Cyan

    try {
        # Check if already connected
        $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
        if ($currentConnection -and $currentConnection.Url -eq $SiteUrl) {
            Write-Host "Already connected to PnP PowerShell" -ForegroundColor Green
            return
        }
    }
    catch {
        # Not connected, proceed with connection
    }

    # Connect with interactive login using registered app
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
    Write-Host "Successfully connected to PnP PowerShell" -ForegroundColor Green
}

function Disconnect-Services {
    <#
    .SYNOPSIS
        Disconnects from all services.
    #>
    Write-Host "Disconnecting from services..." -ForegroundColor Cyan
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
}

#endregion

#region Step 1: Export Functions

function Get-SharedMailboxSendPermissions {
    <#
    .SYNOPSIS
        Gets all users with SendAs or SendOnBehalf permissions for a shared mailbox.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxIdentity
    )

    $permissionUsers = @()

    # Get SendAs permissions
    try {
        $sendAsPermissions = Get-RecipientPermission -Identity $MailboxIdentity -ErrorAction Stop |
            Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.AccessRights -contains "SendAs" }

        foreach ($permission in $sendAsPermissions) {
            if ($permission.Trustee -notlike "S-1-*") {
                $permissionUsers += $permission.Trustee
            }
        }
    }
    catch {
        Write-Warning "Could not get SendAs permissions for $MailboxIdentity : $_"
    }

    # Get SendOnBehalf permissions
    try {
        $mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop
        if ($mailbox.GrantSendOnBehalfTo) {
            foreach ($delegate in $mailbox.GrantSendOnBehalfTo) {
                try {
                    $user = Get-Recipient -Identity $delegate -ErrorAction Stop
                    $permissionUsers += $user.PrimarySmtpAddress
                }
                catch {
                    $permissionUsers += $delegate
                }
            }
        }
    }
    catch {
        Write-Warning "Could not get SendOnBehalf permissions for $MailboxIdentity : $_"
    }

    # Return unique users
    return ($permissionUsers | Select-Object -Unique)
}

function Export-SharedMailboxData {
    <#
    .SYNOPSIS
        Exports all shared mailboxes and their permissions to a CSV file.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    Write-Host "Retrieving all shared mailboxes..." -ForegroundColor Cyan
    $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

    if (-not $sharedMailboxes) {
        Write-Warning "No shared mailboxes found in the tenant."
        return
    }

    Write-Host "Found $($sharedMailboxes.Count) shared mailboxes. Processing..." -ForegroundColor Green

    $exportData = @()
    $counter = 0

    foreach ($mailbox in $sharedMailboxes) {
        $counter++
        $percentComplete = [math]::Round(($counter / $sharedMailboxes.Count) * 100, 2)
        Write-Progress -Activity "Processing Shared Mailboxes" -Status "$counter of $($sharedMailboxes.Count) - $($mailbox.PrimarySmtpAddress)" -PercentComplete $percentComplete

        # Get users with send permissions
        $permissionUsers = Get-SharedMailboxSendPermissions -MailboxIdentity $mailbox.Identity
        $membersString = ($permissionUsers | Sort-Object) -join ";"

        # Calculate group name (replace invalid characters)
        $mailboxPrefix = ($mailbox.PrimarySmtpAddress -split "@")[0]
        $groupName = "smb-mailbox-$mailboxPrefix"
        # Ensure group name is valid (max 256 chars, remove invalid chars)
        $groupName = $groupName -replace '[<>:"/\\|?*]', ''
        if ($groupName.Length -gt 256) {
            $groupName = $groupName.Substring(0, 256)
        }

        $exportData += [PSCustomObject]@{
            'Mailbox address'  = $mailbox.PrimarySmtpAddress
            'Members'          = $membersString
            'MailboxID'        = $mailbox.ExchangeGuid.ToString()
            'M365 Group Name'  = $groupName
            'M365 Group ID'    = ""
            'Status'           = "Ready to create"
        }
    }

    Write-Progress -Activity "Processing Shared Mailboxes" -Completed

    # Export to CSV
    $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported $($exportData.Count) shared mailboxes to: $OutputPath" -ForegroundColor Green

    return $exportData
}

#endregion

#region Step 2: Import and Create Functions

function Get-OrCreateSecurityGroup {
    <#
    .SYNOPSIS
        Gets an existing security group or creates a new one using PnP PowerShell Graph API.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    try {
        # Check if group already exists using Graph API
        $encodedFilter = [System.Web.HttpUtility]::UrlEncode("displayName eq '$DisplayName'")
        $existingGroups = Invoke-PnPGraphMethod -Url "groups?`$filter=$encodedFilter" -Method Get -ErrorAction SilentlyContinue

        if ($existingGroups -and $existingGroups.value -and $existingGroups.value.Count -gt 0) {
            $existingGroup = $existingGroups.value[0]
            Write-Warning "Group '$DisplayName' already exists with ID: $($existingGroup.id)"
            return [PSCustomObject]@{
                Id          = $existingGroup.id
                DisplayName = $existingGroup.displayName
            }
        }

        # Create mail nickname (alphanumeric only)
        $mailNickname = $DisplayName -replace '[^a-zA-Z0-9]', ''
        if ($mailNickname.Length -gt 64) {
            $mailNickname = $mailNickname.Substring(0, 64)
        }

        # Create new security group (mail-disabled) using Graph API
        $groupBody = @{
            displayName     = $DisplayName
            description     = $Description
            mailEnabled     = $false
            mailNickname    = $mailNickname
            securityEnabled = $true
            groupTypes      = @()
        }

        $newGroup = Invoke-PnPGraphMethod -Url "groups" -Method Post -Content $groupBody

        if (-not $newGroup -or -not $newGroup.id) {
            Write-Error "Failed to create group '$DisplayName': No group ID returned"
            return $null
        }

        Write-Host "Created security group: $DisplayName (ID: $($newGroup.id))" -ForegroundColor Green
        return [PSCustomObject]@{
            Id          = $newGroup.id
            DisplayName = $newGroup.displayName
        }
    }
    catch {
        Write-Error "Failed to create group '$DisplayName': $_"
        Write-Host ""
        Write-Host "If you see 'Insufficient privileges', your Azure AD app needs these API permissions:" -ForegroundColor Yellow
        Write-Host "  - Group.ReadWrite.All (Application)" -ForegroundColor Cyan
        Write-Host "  - User.Read.All (Application)" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Add permissions in Azure Portal > App Registrations > Your App > API Permissions" -ForegroundColor Yellow
        Write-Host "Don't forget to click 'Grant admin consent' after adding permissions!" -ForegroundColor Yellow
        return $null
    }
}

function Add-UsersToSecurityGroup {
    <#
    .SYNOPSIS
        Adds users to a security group using PnP PowerShell Graph API.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string[]]$UserEmails
    )

    $addedCount = 0
    $failedCount = 0

    # Get existing members using Graph API
    $existingMemberIds = @()
    try {
        $members = Invoke-PnPGraphMethod -Url "groups/$GroupId/members" -Method Get -ErrorAction SilentlyContinue
        if ($members -and $members.value) {
            $existingMemberIds = $members.value | ForEach-Object { $_.id }
        }
    }
    catch {
        # No existing members or error getting them
    }

    foreach ($email in $UserEmails) {
        if ([string]::IsNullOrWhiteSpace($email)) {
            continue
        }

        try {
            # Get user by email using Graph API
            $encodedFilter = [System.Web.HttpUtility]::UrlEncode("mail eq '$email' or userPrincipalName eq '$email'")
            $userResult = Invoke-PnPGraphMethod -Url "users?`$filter=$encodedFilter" -Method Get -ErrorAction SilentlyContinue

            if (-not $userResult -or -not $userResult.value -or $userResult.value.Count -eq 0) {
                Write-Warning "User not found: $email"
                $failedCount++
                continue
            }

            $user = $userResult.value[0]

            # Check if user is already a member
            if ($existingMemberIds -contains $user.id) {
                Write-Host "User $email is already a member of the group" -ForegroundColor Yellow
                continue
            }

            # Add user to group using Graph API
            $memberBody = @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.id)"
            }
            Invoke-PnPGraphMethod -Url "groups/$GroupId/members/`$ref" -Method Post -Content $memberBody

            Write-Host "Added $email to group" -ForegroundColor Green
            $addedCount++
        }
        catch {
            Write-Warning "Failed to add $email to group: $_"
            $failedCount++
        }
    }

    return @{
        Added  = $addedCount
        Failed = $failedCount
    }
}

function Set-GroupMailboxPermission {
    <#
    .SYNOPSIS
        Grants SendAs permission to a security group for a shared mailbox.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$MailboxIdentity,

        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )

    try {
        # Add SendAs permission using the group's display name
        Add-RecipientPermission -Identity $MailboxIdentity -Trustee $GroupName -AccessRights SendAs -Confirm:$false
        Write-Host "Granted SendAs permission to group for mailbox: $MailboxIdentity" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "Failed to grant SendAs permission: $_"
        return $false
    }
}

function Update-SharePointList {
    <#
    .SYNOPSIS
        Adds or updates an item in the SharePoint list using PnP PowerShell.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListName,

        [Parameter(Mandatory = $true)]
        [string]$SharedMailboxId,

        [Parameter(Mandatory = $true)]
        [string]$SecurityGroupId
    )

    try {
        # Check if item already exists
        $existingItems = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='SharedMailbox'/><Value Type='Text'>$SharedMailboxId</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue

        if ($existingItems -and $existingItems.Count -gt 0) {
            # Update existing item
            $existingItem = $existingItems | Select-Object -First 1
            Set-PnPListItem -List $ListName -Identity $existingItem.Id -Values @{
                "SecurityGroup" = $SecurityGroupId
            } | Out-Null
            Write-Host "Updated SharePoint list item for mailbox: $SharedMailboxId" -ForegroundColor Green
        }
        else {
            # Create new item
            Add-PnPListItem -List $ListName -Values @{
                "SharedMailbox"  = $SharedMailboxId
                "SecurityGroup"  = $SecurityGroupId
            } | Out-Null
            Write-Host "Created SharePoint list item for mailbox: $SharedMailboxId" -ForegroundColor Green
        }

        return $true
    }
    catch {
        Write-Warning "Failed to update SharePoint list: $_"
        return $false
    }
}

function Import-AndProcessMailboxData {
    <#
    .SYNOPSIS
        Imports CSV data and creates security groups for shared mailboxes.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string]$SharePointListName
    )

    if (-not (Test-Path $CsvPath)) {
        Write-Error "CSV file not found: $CsvPath"
        return
    }

    Write-Host "Importing CSV data from: $CsvPath" -ForegroundColor Cyan
    $mailboxData = Import-Csv -Path $CsvPath

    # Filter for rows ready to create
    $rowsToProcess = $mailboxData | Where-Object { $_.'Status' -eq "Ready to create" }

    if (-not $rowsToProcess) {
        Write-Warning "No rows with 'Ready to create' status found."
        return
    }

    Write-Host "Found $($rowsToProcess.Count) mailboxes to process." -ForegroundColor Green

    $counter = 0
    foreach ($row in $rowsToProcess) {
        $counter++
        $percentComplete = [math]::Round(($counter / $rowsToProcess.Count) * 100, 2)
        Write-Progress -Activity "Processing Mailboxes" -Status "$counter of $($rowsToProcess.Count) - $($row.'Mailbox address')" -PercentComplete $percentComplete

        Write-Host "`n--- Processing: $($row.'Mailbox address') ---" -ForegroundColor Cyan

        # Step 1: Create security group
        $description = "Security group for shared mailbox: $($row.'Mailbox address')"
        $group = Get-OrCreateSecurityGroup -DisplayName $row.'M365 Group Name' -Description $description

        if (-not $group) {
            Write-Warning "Skipping mailbox $($row.'Mailbox address') due to group creation failure"
            continue
        }

        # Step 2: Add users to the group
        if ($row.Members) {
            $members = $row.Members -split ";"
            $result = Add-UsersToSecurityGroup -GroupId $group.Id -UserEmails $members
            Write-Host "Added $($result.Added) users, $($result.Failed) failed" -ForegroundColor Yellow
        }
        else {
            Write-Host "No members to add for this mailbox" -ForegroundColor Yellow
        }

        # Step 3: Grant group permission to mailbox
        $permissionGranted = Set-GroupMailboxPermission -MailboxIdentity $row.'Mailbox address' -GroupName $row.'M365 Group Name'

        # Step 4: Update SharePoint list
        $spUpdated = Update-SharePointList -ListName $SharePointListName -SharedMailboxId $row.'MailboxID' -SecurityGroupId $group.Id

        # Step 5: Update the CSV row
        $row.'M365 Group ID' = $group.Id
        if ($permissionGranted -and $spUpdated) {
            $row.'Status' = "Created"
        }
        else {
            $row.'Status' = "Partial - Check logs"
        }
    }

    Write-Progress -Activity "Processing Mailboxes" -Completed

    # Save updated CSV
    $mailboxData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nUpdated CSV saved to: $CsvPath" -ForegroundColor Green
}

#endregion

#region Main Execution

# Main script execution
try {
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "   M365 Shared Mailbox Security Group Manager     " -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""

    # Install required modules
    Write-Host "Checking required modules..." -ForegroundColor Cyan
    Install-RequiredModules

    # Import modules
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Import-Module PnP.PowerShell -ErrorAction Stop

    # Validate ClientId is provided for Import/Both steps
    if ($Step -in @('Import', 'Both') -and [string]::IsNullOrWhiteSpace($ClientId)) {
        Write-Host ""
        Write-Host "ERROR: ClientId is required for Import/Both steps." -ForegroundColor Red
        Write-Host ""
        Write-Host "You need to register an Azure AD application for PnP PowerShell:" -ForegroundColor Yellow
        Write-Host "1. Run this command once:" -ForegroundColor White
        Write-Host "   Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'PnP SharedMailbox Script' -Tenant yourtenant.onmicrosoft.com -Interactive" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "2. After registration, copy the Application (Client) ID" -ForegroundColor White
        Write-Host ""
        Write-Host "3. Run this script with the -ClientId parameter:" -ForegroundColor White
        Write-Host "   .\SharedMailboxManager.ps1 -Step Import -CsvPath 'C:\temp\SharedMailboxes.csv' -ClientId 'your-client-id'" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "For more info: https://pnp.github.io/powershell/articles/registerapplication.html" -ForegroundColor Yellow
        throw "ClientId parameter is required for Import/Both steps."
    }

    switch ($Step) {
        'Export' {
            # Connect to Exchange Online only
            Connect-ExchangeOnlineService

            # Export shared mailbox data
            Export-SharedMailboxData -OutputPath $CsvPath
        }

        'Import' {
            # Connect to Exchange Online
            Connect-ExchangeOnlineService

            # Connect to PnP for Graph and SharePoint operations
            Connect-PnPService -SiteUrl $SharePointSiteUrl -ClientId $ClientId

            # Import and process
            Import-AndProcessMailboxData -CsvPath $CsvPath -SharePointListName $SharePointListName
        }

        'Both' {
            # Connect to Exchange Online
            Connect-ExchangeOnlineService

            # Connect to PnP for Graph and SharePoint operations
            Connect-PnPService -SiteUrl $SharePointSiteUrl -ClientId $ClientId

            # Export first
            Export-SharedMailboxData -OutputPath $CsvPath

            # Then import and process
            Import-AndProcessMailboxData -CsvPath $CsvPath -SharePointListName $SharePointListName
        }
    }

    Write-Host "`nScript completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Script failed: $_"
    Write-Error $_.ScriptStackTrace
}
finally {
    # Disconnect from services
    Disconnect-Services
}

#endregion
