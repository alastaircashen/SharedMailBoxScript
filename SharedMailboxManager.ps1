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

#region Module Installation and Connection Functions

function Install-RequiredModules {
    <#
    .SYNOPSIS
        Installs required PowerShell modules if not already installed.
    #>
    $modules = @(
        'ExchangeOnlineManagement',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Users',
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

function Connect-RequiredServices {
    <#
    .SYNOPSIS
        Connects to required M365 services.
    #>
    param(
        [bool]$ConnectExchange = $true,
        [bool]$ConnectGraph = $true,
        [bool]$ConnectSharePoint = $true,
        [string]$SharePointUrl = ""
    )

    if ($ConnectExchange) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        try {
            Get-ConnectionInformation -ErrorAction Stop | Out-Null
            Write-Host "Already connected to Exchange Online" -ForegroundColor Green
        }
        catch {
            Connect-ExchangeOnline -ShowBanner:$false
        }
    }

    if ($ConnectGraph) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        $graphScopes = @(
            "Group.ReadWrite.All",
            "GroupMember.ReadWrite.All",
            "User.Read.All"
        )
        Connect-MgGraph -Scopes $graphScopes -NoWelcome
    }

    if ($ConnectSharePoint -and $SharePointUrl) {
        Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
        Connect-PnPOnline -Url $SharePointUrl -Interactive
    }
}

function Disconnect-RequiredServices {
    <#
    .SYNOPSIS
        Disconnects from all M365 services.
    #>
    Write-Host "Disconnecting from services..." -ForegroundColor Cyan

    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
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

function New-MailboxSecurityGroup {
    <#
    .SYNOPSIS
        Creates a new security group (mail-disabled) in Azure AD/Entra ID.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    try {
        # Check if group already exists
        $existingGroup = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction SilentlyContinue

        if ($existingGroup) {
            Write-Warning "Group '$GroupName' already exists with ID: $($existingGroup.Id)"
            return $existingGroup
        }

        # Create security group (no mail)
        $groupParams = @{
            DisplayName        = $GroupName
            Description        = $Description
            MailEnabled        = $false
            MailNickname       = ($GroupName -replace '[^a-zA-Z0-9]', '')
            SecurityEnabled    = $true
            GroupTypes         = @()
        }

        $newGroup = New-MgGroup -BodyParameter $groupParams
        Write-Host "Created security group: $GroupName (ID: $($newGroup.Id))" -ForegroundColor Green

        return $newGroup
    }
    catch {
        Write-Error "Failed to create group '$GroupName': $_"
        return $null
    }
}

function Add-UsersToGroup {
    <#
    .SYNOPSIS
        Adds users to a security group.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [string[]]$UserEmails
    )

    $addedCount = 0
    $failedCount = 0

    foreach ($email in $UserEmails) {
        if ([string]::IsNullOrWhiteSpace($email)) {
            continue
        }

        try {
            # Get user by email
            $user = Get-MgUser -Filter "mail eq '$email' or userPrincipalName eq '$email'" -ErrorAction Stop

            if (-not $user) {
                Write-Warning "User not found: $email"
                $failedCount++
                continue
            }

            # Check if user is already a member
            $existingMember = Get-MgGroupMember -GroupId $GroupId -Filter "id eq '$($user.Id)'" -ErrorAction SilentlyContinue

            if ($existingMember) {
                Write-Host "User $email is already a member of the group" -ForegroundColor Yellow
                continue
            }

            # Add user to group
            New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $user.Id
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
        [string]$GroupId
    )

    try {
        # Get the group details
        $group = Get-MgGroup -GroupId $GroupId

        # Add SendAs permission using the group's display name
        Add-RecipientPermission -Identity $MailboxIdentity -Trustee $group.DisplayName -AccessRights SendAs -Confirm:$false
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
        Adds or updates an item in the SharePoint list.
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
        $existingItem = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='SharedMailbox'/><Value Type='Text'>$SharedMailboxId</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue

        if ($existingItem) {
            # Update existing item
            Set-PnPListItem -List $ListName -Identity $existingItem.Id -Values @{
                "SecurityGroup" = $SecurityGroupId
            }
            Write-Host "Updated SharePoint list item for mailbox: $SharedMailboxId" -ForegroundColor Green
        }
        else {
            # Create new item
            Add-PnPListItem -List $ListName -Values @{
                "SharedMailbox"  = $SharedMailboxId
                "SecurityGroup"  = $SecurityGroupId
            }
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
        $group = New-MailboxSecurityGroup -GroupName $row.'M365 Group Name' -Description $description

        if (-not $group) {
            Write-Warning "Skipping mailbox $($row.'Mailbox address') due to group creation failure"
            continue
        }

        # Step 2: Add users to the group
        if ($row.Members) {
            $members = $row.Members -split ";"
            $result = Add-UsersToGroup -GroupId $group.Id -UserEmails $members
            Write-Host "Added $($result.Added) users, $($result.Failed) failed" -ForegroundColor Yellow
        }
        else {
            Write-Host "No members to add for this mailbox" -ForegroundColor Yellow
        }

        # Step 3: Grant group permission to mailbox
        $permissionGranted = Set-GroupMailboxPermission -MailboxIdentity $row.'Mailbox address' -GroupId $group.Id

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
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module PnP.PowerShell -ErrorAction Stop

    switch ($Step) {
        'Export' {
            # Connect to Exchange Online only
            Connect-RequiredServices -ConnectExchange $true -ConnectGraph $false -ConnectSharePoint $false

            # Export shared mailbox data
            Export-SharedMailboxData -OutputPath $CsvPath
        }

        'Import' {
            # Connect to all services
            Connect-RequiredServices -ConnectExchange $true -ConnectGraph $true -ConnectSharePoint $true -SharePointUrl $SharePointSiteUrl

            # Import and process
            Import-AndProcessMailboxData -CsvPath $CsvPath -SharePointListName $SharePointListName
        }

        'Both' {
            # Connect to all services
            Connect-RequiredServices -ConnectExchange $true -ConnectGraph $true -ConnectSharePoint $true -SharePointUrl $SharePointSiteUrl

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
    Disconnect-RequiredServices
}

#endregion
