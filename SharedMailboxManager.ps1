<#
.SYNOPSIS
    Manages M365 Shared Mailboxes - Exports mailbox data and syncs security groups with matching membership.

.DESCRIPTION
    This script provides two main functions:
    Step 1: Export all shared mailboxes and their send permissions to a CSV file
    Step 2: Import the CSV, create security groups with matching membership, and update SharePoint

    The script ensures each shared mailbox has a corresponding security group with identical
    membership. It can be run at any time to sync group membership with mailbox permissions.

    Uses only ExchangeOnlineManagement and PnP.PowerShell modules.

    PREREQUISITES:

    1. SHAREPOINT LIST
       The SharePoint list "Shared Mailboxes Mapping" should already exist in your environment.
       Update the -SharePointSiteUrl parameter to point to your SharePoint site.
       List requires two text columns: SharedMailbox, SecurityGroup

    2. PNP APP REGISTRATION
       First, check if you have an existing PnP PowerShell app registration:
       - Go to Azure Portal > App Registrations > Search for "PnP"
       - Verify it has these DELEGATED permissions (not Application):
         * Microsoft Graph: Group.ReadWrite.All, User.Read.All
         * SharePoint: AllSites.FullControl
       - Ensure admin consent has been granted

       If no existing registration, create one:
       Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell" -Tenant yourtenant.onmicrosoft.com -Interactive

       See: https://pnp.github.io/powershell/articles/registerapplication.html

.PARAMETER Step
    Specify which step to run: 'Export', 'Import', or 'Both'

.PARAMETER CsvPath
    Path to the CSV file (for export or import)

.PARAMETER ClientId
    The Azure AD Application (Client) ID for PnP PowerShell authentication.
    Required for Import and Both steps.

.PARAMETER SharePointSiteUrl
    SharePoint site URL containing the Shared Mailboxes Mapping list.
    Update this to your client's SharePoint site URL.

.PARAMETER SharePointListName
    Name of the SharePoint list for storing mailbox-to-group mappings.
    Default: "Shared Mailboxes Mapping"

.PARAMETER TestMode
    When specified, only processes the first 5 mailboxes. Useful for testing.

.PARAMETER ExclusionListPath
    Path to a CSV file containing shared mailbox addresses to exclude from processing.
    The CSV should have a single column with mailbox email addresses. A header row is expected
    (the column can have any name). Excluded mailboxes will be skipped during both export and import.

.EXAMPLE
    # Export shared mailboxes to CSV
    .\SharedMailboxManager.ps1 -Step Export -CsvPath "C:\temp\SharedMailboxes.csv"

.EXAMPLE
    # Import and process mailboxes (test mode - first 5 only)
    .\SharedMailboxManager.ps1 -Step Import -CsvPath "C:\temp\SharedMailboxes.csv" -TestMode

.EXAMPLE
    # Import and process all mailboxes
    .\SharedMailboxManager.ps1 -Step Import -CsvPath "C:\temp\SharedMailboxes.csv"

.EXAMPLE
    # Export and import in one run
    .\SharedMailboxManager.ps1 -Step Both -CsvPath "C:\temp\SharedMailboxes.csv"

.EXAMPLE
    # Export with exclusion list (skip specific mailboxes)
    .\SharedMailboxManager.ps1 -Step Export -CsvPath "C:\temp\SharedMailboxes.csv" -ExclusionListPath "C:\temp\ExcludeMailboxes.csv"

.EXAMPLE
    # Import with exclusion list
    .\SharedMailboxManager.ps1 -Step Import -CsvPath "C:\temp\SharedMailboxes.csv" -ExclusionListPath "C:\temp\ExcludeMailboxes.csv"
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
    [string]$SharePointListName = "Shared Mailboxes Mapping",

    [Parameter(Mandatory = $false)]
    [switch]$TestMode,

    [Parameter(Mandatory = $false)]
    [string]$ExclusionListPath
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

function Get-ExclusionList {
    <#
    .SYNOPSIS
        Loads a list of shared mailbox addresses to exclude from a CSV file.
        The CSV should have a single column with mailbox addresses (with or without a header).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        Write-Error "Exclusion list file not found: $Path"
        throw "Exclusion list file not found: $Path"
    }

    $exclusions = @()

    # Read the CSV content
    $csvContent = Import-Csv -Path $Path

    if ($csvContent -and $csvContent.Count -gt 0) {
        # Get the first (and expected only) column name
        $columnNames = $csvContent[0].PSObject.Properties.Name
        if ($columnNames.Count -gt 0) {
            $columnName = $columnNames[0]
            $exclusions = $csvContent | ForEach-Object {
                $_.$columnName.Trim().ToLower()
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
    }

    Write-Host "Loaded $($exclusions.Count) mailbox(es) from exclusion list" -ForegroundColor Yellow
    foreach ($addr in $exclusions) {
        Write-Host "  Excluding: $addr" -ForegroundColor Yellow
    }

    return $exclusions
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
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [string[]]$ExclusionList = @()
    )

    Write-Host "Retrieving all shared mailboxes..." -ForegroundColor Cyan
    $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

    if (-not $sharedMailboxes) {
        Write-Warning "No shared mailboxes found in the tenant."
        return
    }

    Write-Host "Found $($sharedMailboxes.Count) shared mailboxes." -ForegroundColor Green

    # Apply exclusion list
    if ($ExclusionList -and $ExclusionList.Count -gt 0) {
        $beforeCount = $sharedMailboxes.Count
        $sharedMailboxes = $sharedMailboxes | Where-Object {
            $_.PrimarySmtpAddress.ToLower() -notin $ExclusionList
        }
        $excludedCount = $beforeCount - @($sharedMailboxes).Count
        Write-Host "Excluded $excludedCount mailbox(es) from exclusion list. Processing $(@($sharedMailboxes).Count) remaining..." -ForegroundColor Yellow
    }
    else {
        Write-Host "Processing..." -ForegroundColor Green
    }

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
            'Mailbox address'       = $mailbox.PrimarySmtpAddress
            'Members'               = $membersString
            'MailboxID'             = $mailbox.ExchangeGuid.ToString()
            'SharedMailboxObjectID' = $mailbox.ExternalDirectoryObjectId
            'M365 Group Name'       = $groupName
            'M365 Group ID'         = ""
            'ExistingMapping'       = ""
            'Status'                = "Ready to create"
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
        If group exists, compares members to expected list.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $false)]
        [string[]]$ExpectedMembers = @()
    )

    try {
        # Check if group already exists using Graph API
        $encodedFilter = [System.Web.HttpUtility]::UrlEncode("displayName eq '$DisplayName'")
        $existingGroups = Invoke-PnPGraphMethod -Url "groups?`$filter=$encodedFilter" -Method Get -ErrorAction SilentlyContinue

        if ($existingGroups -and $existingGroups.value -and $existingGroups.value.Count -gt 0) {
            $existingGroup = $existingGroups.value[0]
            Write-Host "Group '$DisplayName' already exists with ID: $($existingGroup.id)" -ForegroundColor Yellow

            # Get current members of the existing group
            $currentMembers = @()
            try {
                $membersResult = Invoke-PnPGraphMethod -Url "groups/$($existingGroup.id)/members" -Method Get -ErrorAction SilentlyContinue
                if ($membersResult -and $membersResult.value) {
                    $currentMembers = $membersResult.value | ForEach-Object {
                        if ($_.mail) { $_.mail.ToLower() }
                        elseif ($_.userPrincipalName) { $_.userPrincipalName.ToLower() }
                    } | Where-Object { $_ }
                }
            }
            catch {
                Write-Warning "Could not retrieve current group members: $_"
            }

            # Compare members if expected members provided
            if ($ExpectedMembers -and $ExpectedMembers.Count -gt 0) {
                $expectedLower = $ExpectedMembers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToLower() }

                $missingMembers = $expectedLower | Where-Object { $_ -notin $currentMembers }
                $extraMembers = $currentMembers | Where-Object { $_ -notin $expectedLower }

                if ($missingMembers -or $extraMembers) {
                    Write-Warning "Group members differ from expected mailbox permissions!"
                    if ($missingMembers) {
                        Write-Host "  Missing from group: $($missingMembers -join ', ')" -ForegroundColor Red
                    }
                    if ($extraMembers) {
                        Write-Host "  Extra in group: $($extraMembers -join ', ')" -ForegroundColor Red
                    }
                    return [PSCustomObject]@{
                        Id            = $existingGroup.id
                        DisplayName   = $existingGroup.displayName
                        AlreadyExists = $true
                        MembersMismatch = $true
                        MissingMembers = $missingMembers
                        ExtraMembers   = $extraMembers
                    }
                }
            }

            return [PSCustomObject]@{
                Id            = $existingGroup.id
                DisplayName   = $existingGroup.displayName
                AlreadyExists = $true
                MembersMismatch = $false
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
            Id            = $newGroup.id
            DisplayName   = $newGroup.displayName
            AlreadyExists = $false
            MembersMismatch = $false
        }
    }
    catch {
        Write-Error "Failed to create group '$DisplayName': $_"
        Write-Host ""
        Write-Host "If you see 'Insufficient privileges', your Azure AD app needs these API permissions:" -ForegroundColor Yellow
        Write-Host "  - Group.ReadWrite.All (Delegated)" -ForegroundColor Cyan
        Write-Host "  - User.Read.All (Delegated)" -ForegroundColor Cyan
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

function Get-ExistingSharePointMappings {
    <#
    .SYNOPSIS
        Gets all existing mappings from the SharePoint list.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListName
    )

    try {
        $allItems = Get-PnPListItem -List $ListName -PageSize 500 -ErrorAction SilentlyContinue
        $mappings = @{}

        if ($allItems) {
            foreach ($item in $allItems) {
                $sharedMailboxId = $item.FieldValues["SharedMailbox"]
                if ($sharedMailboxId) {
                    if (-not $mappings.ContainsKey($sharedMailboxId)) {
                        $mappings[$sharedMailboxId] = @()
                    }
                    $mappings[$sharedMailboxId] += [PSCustomObject]@{
                        ItemId         = $item.Id
                        SharedMailbox  = $sharedMailboxId
                        SecurityGroup  = $item.FieldValues["SecurityGroup"]
                    }
                }
            }
        }

        return $mappings
    }
    catch {
        Write-Warning "Failed to get existing SharePoint mappings: $_"
        return @{}
    }
}

function Update-SharePointList {
    <#
    .SYNOPSIS
        Adds an item to the SharePoint list if the mailbox-group mapping doesn't already exist.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListName,

        [Parameter(Mandatory = $true)]
        [string]$SharedMailboxObjectId,

        [Parameter(Mandatory = $true)]
        [string]$SecurityGroupId,

        [Parameter(Mandatory = $false)]
        [hashtable]$ExistingMappings = @{}
    )

    try {
        # Check if this exact mailbox-group mapping already exists
        if ($ExistingMappings.ContainsKey($SharedMailboxObjectId)) {
            $existingItems = $ExistingMappings[$SharedMailboxObjectId]
            $matchingItem = $existingItems | Where-Object { $_.SecurityGroup -eq $SecurityGroupId }
            if ($matchingItem) {
                Write-Host "SharePoint mapping already exists for this mailbox-group combination - skipping" -ForegroundColor Yellow
                return $true
            }
        }

        # Create new item since mapping doesn't exist
        Add-PnPListItem -List $ListName -Values @{
            "SharedMailbox"  = $SharedMailboxObjectId
            "SecurityGroup"  = $SecurityGroupId
        } | Out-Null
        Write-Host "Created SharePoint list item for mailbox: $SharedMailboxObjectId" -ForegroundColor Green

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
        Imports CSV data and creates/syncs security groups for shared mailboxes.
        Ensures group membership matches mailbox permissions. Can be run repeatedly.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string]$SharePointListName,

        [Parameter(Mandatory = $false)]
        [bool]$TestMode = $false,

        [Parameter(Mandatory = $false)]
        [string[]]$ExclusionList = @()
    )

    if (-not (Test-Path $CsvPath)) {
        Write-Error "CSV file not found: $CsvPath"
        return
    }

    if ($TestMode) {
        Write-Host ""
        Write-Host "*** TEST MODE ENABLED - Only processing first 5 mailboxes ***" -ForegroundColor Magenta
        Write-Host ""
    }

    Write-Host "Importing CSV data from: $CsvPath" -ForegroundColor Cyan
    $mailboxData = Import-Csv -Path $CsvPath

    # Apply exclusion list
    if ($ExclusionList -and $ExclusionList.Count -gt 0) {
        $beforeCount = @($mailboxData).Count
        $mailboxData = $mailboxData | Where-Object {
            $_.'Mailbox address'.ToLower() -notin $ExclusionList
        }
        $excludedCount = $beforeCount - @($mailboxData).Count
        Write-Host "Excluded $excludedCount mailbox(es) from exclusion list. $(@($mailboxData).Count) remaining." -ForegroundColor Yellow
    }

    # Get existing SharePoint mappings to check for duplicates
    Write-Host "Checking for existing SharePoint mappings..." -ForegroundColor Cyan
    $existingMappings = Get-ExistingSharePointMappings -ListName $SharePointListName
    Write-Host "Found $($existingMappings.Count) unique mailbox mappings in SharePoint" -ForegroundColor Yellow

    # Filter for rows to process (skip only error rows)
    $rowsToProcess = $mailboxData | Where-Object { $_.'Status' -notlike "Error*" }

    if (-not $rowsToProcess) {
        Write-Warning "No rows to process found."
        return
    }

    # Limit to first 5 in test mode
    $totalCount = @($rowsToProcess).Count
    if ($TestMode -and $totalCount -gt 5) {
        $rowsToProcess = $rowsToProcess | Select-Object -First 5
        Write-Host "TEST MODE: Syncing first 5 of $totalCount mailboxes" -ForegroundColor Magenta
    }

    Write-Host "Found $(@($rowsToProcess).Count) mailboxes to sync." -ForegroundColor Green

    $counter = 0
    foreach ($row in $rowsToProcess) {
        $counter++
        $percentComplete = [math]::Round(($counter / $rowsToProcess.Count) * 100, 2)
        Write-Progress -Activity "Syncing Mailbox Groups" -Status "$counter of $($rowsToProcess.Count) - $($row.'Mailbox address')" -PercentComplete $percentComplete

        Write-Host "`n--- Syncing: $($row.'Mailbox address') ---" -ForegroundColor Cyan

        # Check for existing SharePoint mapping
        $mailboxObjectId = $row.'SharedMailboxObjectID'
        if ($existingMappings.ContainsKey($mailboxObjectId)) {
            $existingCount = $existingMappings[$mailboxObjectId].Count
            $row.'ExistingMapping' = "Yes - $existingCount existing mapping(s)"
            Write-Host "WARNING: Found $existingCount existing mapping(s) in SharePoint for this mailbox" -ForegroundColor Yellow
        }
        else {
            $row.'ExistingMapping' = "No"
        }

        # Parse members list
        $members = @()
        if ($row.Members) {
            $members = $row.Members -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }

        # Step 1: Get or create security group (with member comparison)
        $description = "Security group for shared mailbox: $($row.'Mailbox address')"
        $group = Get-OrCreateSecurityGroup -DisplayName $row.'M365 Group Name' -Description $description -ExpectedMembers $members

        if (-not $group) {
            Write-Warning "Skipping mailbox $($row.'Mailbox address') due to group creation failure"
            $row.'Status' = "Error - Group creation failed"
            continue
        }

        # Step 2: Sync group members to match mailbox permissions
        if ($group.AlreadyExists) {
            Write-Host "Group already exists" -ForegroundColor Yellow
            if ($group.MembersMismatch -and $group.MissingMembers) {
                # Filter out any GUIDs from missing members (these are service principals or groups, not user emails)
                $membersToAdd = $group.MissingMembers | Where-Object {
                    $_ -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
                }

                if ($membersToAdd -and $membersToAdd.Count -gt 0) {
                    Write-Host "Syncing missing members to group: $($membersToAdd -join ', ')" -ForegroundColor Cyan
                    $result = Add-UsersToSecurityGroup -GroupId $group.Id -UserEmails $membersToAdd
                    Write-Host "Added $($result.Added) users, $($result.Failed) failed" -ForegroundColor Yellow
                }
                else {
                    Write-Host "No user members to sync (only non-user entities found in difference)" -ForegroundColor Yellow
                }

                if ($group.ExtraMembers) {
                    Write-Host "Note: Extra members in group (not in mailbox permissions): $($group.ExtraMembers -join ', ')" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "Group members are in sync" -ForegroundColor Green
            }
        }
        else {
            # Add all members for newly created groups
            if ($members.Count -gt 0) {
                $result = Add-UsersToSecurityGroup -GroupId $group.Id -UserEmails $members
                Write-Host "Added $($result.Added) users, $($result.Failed) failed" -ForegroundColor Yellow
            }
            else {
                Write-Host "No members to add for this mailbox" -ForegroundColor Yellow
            }
        }

        # Step 3: Update SharePoint list with SharedMailboxObjectID (only if mapping doesn't exist)
        $spUpdated = Update-SharePointList -ListName $SharePointListName -SharedMailboxObjectId $mailboxObjectId -SecurityGroupId $group.Id -ExistingMappings $existingMappings

        # Step 4: Update the CSV row
        $row.'M365 Group ID' = $group.Id
        if ($spUpdated) {
            $row.'Status' = "Synced"
        }
        else {
            $row.'Status' = "Synced - SharePoint update failed"
        }
    }

    Write-Progress -Activity "Syncing Mailbox Groups" -Completed

    # Save updated CSV
    $mailboxData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nUpdated CSV saved to: $CsvPath" -ForegroundColor Green
}

#endregion

#region Main Execution

# Main script execution
try {
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "  M365 Shared Mailbox Security Group Sync Tool    " -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""

    # Install required modules
    Write-Host "Checking required modules..." -ForegroundColor Cyan
    Install-RequiredModules

    # Import modules
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Import-Module PnP.PowerShell -ErrorAction Stop

    # Load exclusion list if provided
    $exclusionList = @()
    if (-not [string]::IsNullOrWhiteSpace($ExclusionListPath)) {
        $exclusionList = Get-ExclusionList -Path $ExclusionListPath
    }

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
            Export-SharedMailboxData -OutputPath $CsvPath -ExclusionList $exclusionList
        }

        'Import' {
            # Connect to Exchange Online
            Connect-ExchangeOnlineService

            # Connect to PnP for Graph and SharePoint operations
            Connect-PnPService -SiteUrl $SharePointSiteUrl -ClientId $ClientId

            # Import and process
            Import-AndProcessMailboxData -CsvPath $CsvPath -SharePointListName $SharePointListName -TestMode $TestMode.IsPresent -ExclusionList $exclusionList
        }

        'Both' {
            # Connect to Exchange Online
            Connect-ExchangeOnlineService

            # Connect to PnP for Graph and SharePoint operations
            Connect-PnPService -SiteUrl $SharePointSiteUrl -ClientId $ClientId

            # Export first
            Export-SharedMailboxData -OutputPath $CsvPath -ExclusionList $exclusionList

            # Then import and process
            Import-AndProcessMailboxData -CsvPath $CsvPath -SharePointListName $SharePointListName -TestMode $TestMode.IsPresent -ExclusionList $exclusionList
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
