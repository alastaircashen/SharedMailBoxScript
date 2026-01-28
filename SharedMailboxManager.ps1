<#
.SYNOPSIS
    Manages M365 Shared Mailboxes - Exports mailbox data and creates security groups for permissions.

.DESCRIPTION
    This script provides two main functions:
    Step 1: Export all shared mailboxes and their send permissions to a CSV file
    Step 2: Import the CSV, create security groups, assign permissions, and update SharePoint

.PARAMETER Step
    Specify which step to run: 'Export', 'Import', or 'Both'

.PARAMETER CsvPath
    Path to the CSV file (for export or import)

.PARAMETER SharePointSiteUrl
    SharePoint site URL for the list update

.EXAMPLE
    .\SharedMailboxManager.ps1 -Step Export -CsvPath "C:\temp\SharedMailboxes.csv"
    .\SharedMailboxManager.ps1 -Step Import -CsvPath "C:\temp\SharedMailboxes.csv"
    .\SharedMailboxManager.ps1 -Step Both -CsvPath "C:\temp\SharedMailboxes.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Export', 'Import', 'Both')]
    [string]$Step,

    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$SharePointSiteUrl = "https://zn8r8.sharepoint.com/sites/DMData",

    [Parameter(Mandatory = $false)]
    [string]$SharePointListName = "SharedMailboxesMapping"
)

# Set execution policy to allow running unsigned scripts (current process only)
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

#region Global Variables
$script:GraphToken = $null
$script:SharePointToken = $null
#endregion

#region Module and Authentication Functions

function Install-RequiredModules {
    <#
    .SYNOPSIS
        Installs required PowerShell modules if not already installed.
    #>
    $modules = @(
        'ExchangeOnlineManagement',
        'MSAL.PS'
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

function Get-GraphAccessToken {
    <#
    .SYNOPSIS
        Gets an access token for Microsoft Graph using device code flow.
    #>
    Write-Host "Authenticating to Microsoft Graph..." -ForegroundColor Cyan
    Write-Host "A browser window will open for authentication." -ForegroundColor Yellow

    # Use Azure AD PowerShell app ID (well-known) for device code flow
    $clientId = "1950a258-227b-4e31-a9cf-717495945fc2"  # Azure PowerShell
    $tenantId = "organizations"
    $scope = "https://graph.microsoft.com/.default"

    try {
        $tokenResponse = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Scopes $scope -DeviceCode
        $script:GraphToken = $tokenResponse.AccessToken
        Write-Host "Successfully authenticated to Microsoft Graph" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to get Graph token: $_"
        return $false
    }
}

function Get-SharePointAccessToken {
    <#
    .SYNOPSIS
        Gets an access token for SharePoint using device code flow.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SharePointUrl
    )

    Write-Host "Authenticating to SharePoint..." -ForegroundColor Cyan

    # Extract tenant name from SharePoint URL
    $uri = [System.Uri]$SharePointUrl
    $sharePointResource = "https://$($uri.Host)"

    $clientId = "1950a258-227b-4e31-a9cf-717495945fc2"  # Azure PowerShell
    $tenantId = "organizations"
    $scope = "$sharePointResource/.default"

    try {
        $tokenResponse = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Scopes $scope -DeviceCode
        $script:SharePointToken = $tokenResponse.AccessToken
        Write-Host "Successfully authenticated to SharePoint" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to get SharePoint token: $_"
        return $false
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

function Disconnect-Services {
    <#
    .SYNOPSIS
        Disconnects from all services.
    #>
    Write-Host "Disconnecting from services..." -ForegroundColor Cyan
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    $script:GraphToken = $null
    $script:SharePointToken = $null
}

#endregion

#region Microsoft Graph REST API Functions

function Invoke-GraphRequest {
    <#
    .SYNOPSIS
        Makes a REST API call to Microsoft Graph.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Endpoint,

        [Parameter(Mandatory = $false)]
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',

        [Parameter(Mandatory = $false)]
        [object]$Body = $null
    )

    $headers = @{
        'Authorization' = "Bearer $($script:GraphToken)"
        'Content-Type'  = 'application/json'
    }

    $uri = "https://graph.microsoft.com/v1.0$Endpoint"

    $params = @{
        Uri     = $uri
        Headers = $headers
        Method  = $Method
    }

    if ($Body) {
        $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
    }

    return Invoke-RestMethod @params
}

function Get-GraphGroup {
    <#
    .SYNOPSIS
        Gets a group by display name from Microsoft Graph.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    $encodedName = [System.Web.HttpUtility]::UrlEncode("displayName eq '$DisplayName'")
    $result = Invoke-GraphRequest -Endpoint "/groups?`$filter=$encodedName"
    return $result.value | Select-Object -First 1
}

function New-GraphSecurityGroup {
    <#
    .SYNOPSIS
        Creates a new security group via Microsoft Graph.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    # Check if group already exists
    $existingGroup = Get-GraphGroup -DisplayName $DisplayName
    if ($existingGroup) {
        Write-Warning "Group '$DisplayName' already exists with ID: $($existingGroup.id)"
        return $existingGroup
    }

    $mailNickname = $DisplayName -replace '[^a-zA-Z0-9]', ''
    if ($mailNickname.Length -gt 64) {
        $mailNickname = $mailNickname.Substring(0, 64)
    }

    $body = @{
        displayName     = $DisplayName
        description     = $Description
        mailEnabled     = $false
        mailNickname    = $mailNickname
        securityEnabled = $true
        groupTypes      = @()
    }

    $newGroup = Invoke-GraphRequest -Endpoint "/groups" -Method POST -Body $body
    Write-Host "Created security group: $DisplayName (ID: $($newGroup.id))" -ForegroundColor Green
    return $newGroup
}

function Get-GraphUser {
    <#
    .SYNOPSIS
        Gets a user by email from Microsoft Graph.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    try {
        $encodedFilter = [System.Web.HttpUtility]::UrlEncode("mail eq '$Email' or userPrincipalName eq '$Email'")
        $result = Invoke-GraphRequest -Endpoint "/users?`$filter=$encodedFilter"
        return $result.value | Select-Object -First 1
    }
    catch {
        return $null
    }
}

function Add-GraphGroupMember {
    <#
    .SYNOPSIS
        Adds a user to a group via Microsoft Graph.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    $body = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
    }

    Invoke-GraphRequest -Endpoint "/groups/$GroupId/members/`$ref" -Method POST -Body $body
}

function Get-GraphGroupMembers {
    <#
    .SYNOPSIS
        Gets members of a group via Microsoft Graph.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    $result = Invoke-GraphRequest -Endpoint "/groups/$GroupId/members"
    return $result.value
}

#endregion

#region SharePoint REST API Functions

function Invoke-SharePointRequest {
    <#
    .SYNOPSIS
        Makes a REST API call to SharePoint.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$Endpoint,

        [Parameter(Mandatory = $false)]
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',

        [Parameter(Mandatory = $false)]
        [object]$Body = $null
    )

    $headers = @{
        'Authorization' = "Bearer $($script:SharePointToken)"
        'Accept'        = 'application/json;odata=verbose'
        'Content-Type'  = 'application/json;odata=verbose'
    }

    $uri = "$SiteUrl/_api$Endpoint"

    $params = @{
        Uri     = $uri
        Headers = $headers
        Method  = $Method
    }

    if ($Body) {
        $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
    }

    return Invoke-RestMethod @params
}

function Get-SharePointListItemType {
    <#
    .SYNOPSIS
        Gets the list item type for a SharePoint list.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    $result = Invoke-SharePointRequest -SiteUrl $SiteUrl -Endpoint "/web/lists/getbytitle('$ListName')?`$select=ListItemEntityTypeFullName"
    return $result.d.ListItemEntityTypeFullName
}

function Add-SharePointListItem {
    <#
    .SYNOPSIS
        Adds an item to a SharePoint list.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$ListName,

        [Parameter(Mandatory = $true)]
        [string]$SharedMailboxId,

        [Parameter(Mandatory = $true)]
        [string]$SecurityGroupId
    )

    try {
        $listItemType = Get-SharePointListItemType -SiteUrl $SiteUrl -ListName $ListName

        $body = @{
            '__metadata'    = @{ 'type' = $listItemType }
            'SharedMailbox' = $SharedMailboxId
            'SecurityGroup' = $SecurityGroupId
        }

        $result = Invoke-SharePointRequest -SiteUrl $SiteUrl -Endpoint "/web/lists/getbytitle('$ListName')/items" -Method POST -Body $body
        Write-Host "Created SharePoint list item for mailbox: $SharedMailboxId" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "Failed to add SharePoint list item: $_"
        return $false
    }
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

function Add-UsersToSecurityGroup {
    <#
    .SYNOPSIS
        Adds users to a security group using Graph API.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string[]]$UserEmails
    )

    $addedCount = 0
    $failedCount = 0

    # Get existing members
    $existingMembers = Get-GraphGroupMembers -GroupId $GroupId
    $existingMemberIds = $existingMembers | ForEach-Object { $_.id }

    foreach ($email in $UserEmails) {
        if ([string]::IsNullOrWhiteSpace($email)) {
            continue
        }

        try {
            # Get user by email
            $user = Get-GraphUser -Email $email

            if (-not $user) {
                Write-Warning "User not found: $email"
                $failedCount++
                continue
            }

            # Check if user is already a member
            if ($existingMemberIds -contains $user.id) {
                Write-Host "User $email is already a member of the group" -ForegroundColor Yellow
                continue
            }

            # Add user to group
            Add-GraphGroupMember -GroupId $GroupId -UserId $user.id
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

function Import-AndProcessMailboxData {
    <#
    .SYNOPSIS
        Imports CSV data and creates security groups for shared mailboxes.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string]$SharePointSiteUrl,

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
        $group = New-GraphSecurityGroup -DisplayName $row.'M365 Group Name' -Description $description

        if (-not $group) {
            Write-Warning "Skipping mailbox $($row.'Mailbox address') due to group creation failure"
            continue
        }

        # Step 2: Add users to the group
        if ($row.Members) {
            $members = $row.Members -split ";"
            $result = Add-UsersToSecurityGroup -GroupId $group.id -UserEmails $members
            Write-Host "Added $($result.Added) users, $($result.Failed) failed" -ForegroundColor Yellow
        }
        else {
            Write-Host "No members to add for this mailbox" -ForegroundColor Yellow
        }

        # Step 3: Grant group permission to mailbox
        $permissionGranted = Set-GroupMailboxPermission -MailboxIdentity $row.'Mailbox address' -GroupName $row.'M365 Group Name'

        # Step 4: Update SharePoint list
        $spUpdated = Add-SharePointListItem -SiteUrl $SharePointSiteUrl -ListName $SharePointListName -SharedMailboxId $row.'MailboxID' -SecurityGroupId $group.id

        # Step 5: Update the CSV row
        $row.'M365 Group ID' = $group.id
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

# Add System.Web assembly for URL encoding
Add-Type -AssemblyName System.Web

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
    Import-Module MSAL.PS -ErrorAction Stop

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

            # Get Graph token
            if (-not (Get-GraphAccessToken)) {
                throw "Failed to authenticate to Microsoft Graph"
            }

            # Get SharePoint token
            if (-not (Get-SharePointAccessToken -SharePointUrl $SharePointSiteUrl)) {
                throw "Failed to authenticate to SharePoint"
            }

            # Import and process
            Import-AndProcessMailboxData -CsvPath $CsvPath -SharePointSiteUrl $SharePointSiteUrl -SharePointListName $SharePointListName
        }

        'Both' {
            # Connect to Exchange Online
            Connect-ExchangeOnlineService

            # Get Graph token
            if (-not (Get-GraphAccessToken)) {
                throw "Failed to authenticate to Microsoft Graph"
            }

            # Get SharePoint token
            if (-not (Get-SharePointAccessToken -SharePointUrl $SharePointSiteUrl)) {
                throw "Failed to authenticate to SharePoint"
            }

            # Export first
            Export-SharedMailboxData -OutputPath $CsvPath

            # Then import and process
            Import-AndProcessMailboxData -CsvPath $CsvPath -SharePointSiteUrl $SharePointSiteUrl -SharePointListName $SharePointListName
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
