# Advanced Azure User Offboarding Script
# Version: 3.0
# Last Updated: 2024-12-09
# This script provides a comprehensive solution for offboarding users from Azure AD/Microsoft 365
# with enhanced error handling, progress tracking, and detailed reporting

#Region Helper Functions

# Function to create consistent styled console output
# Function to show progress with timestamp
function Write-ProgressStatus {
    param (
        [string]$Message,
        [ValidateSet('Progress', 'Success', 'Failed', 'Warning', 'Info')]
        [string]$Status = 'Info',
        [string]$Color = 'White'
    )
    
    # Using simpler ASCII characters that work well in PowerShell
    $statusSymbol = switch ($Status) {
        'Progress' { '>' }
        'Success'  { '+' }
        'Failed'   { '!' }
        'Warning'  { '*' }
        'Info'     { '-' }
        default    { ' ' }
    }
    
    # Build the status line with proper spacing
    $statusLine = "$statusSymbol $Message"
    
    # Use consistent colors for different status types
    $displayColor = switch ($Status) {
        'Progress' { 'Yellow' }
        'Success'  { 'Green' }
        'Failed'   { 'Red' }
        'Warning'  { 'Yellow' }
        'Info'     { $Color }
        default    { $Color }
    }
    
    Write-Host $statusLine -ForegroundColor $displayColor
}

# Retry mechanism for handling transient failures
function Invoke-WithRetry {
    param (
        [Parameter(Mandatory=$true)]
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 2
    )
    
    $attempts = 0
    do {
        $attempts++
        try {
            return $ScriptBlock.Invoke()
        }
        catch {
            if ($attempts -eq $MaxAttempts) {
                throw "Operation failed after $MaxAttempts attempts: $_"
            }
            Write-ProgressStatus -Message "Attempt $attempts failed. Retrying in $DelaySeconds seconds..." -Status Warning -Color Yellow
            Start-Sleep -Seconds $DelaySeconds
        }
    } while ($attempts -lt $MaxAttempts)
}

# Function for menu navigation
function Show-NavigationOptions {
    param (
        [switch]$AllowBack,
        [switch]$AllowMainMenu,
        [switch]$AllowContinue
    )
    
    Write-Host "`nNavigation Options:" -ForegroundColor Cyan
    if ($AllowBack) {
        Write-Host "B: Go Back" -ForegroundColor Yellow
    }
    if ($AllowMainMenu) {
        Write-Host "M: Return to Main Menu" -ForegroundColor Yellow
    }
    if ($AllowContinue) {
        Write-Host "C: Continue" -ForegroundColor Yellow
    }
}

function Get-NavigationChoice {
    param (
        [switch]$AllowBack,
        [switch]$AllowMainMenu,
        [switch]$AllowContinue
    )
    
    $validChoices = @()
    if ($AllowBack) { $validChoices += 'B' }
    if ($AllowMainMenu) { $validChoices += 'M' }
    if ($AllowContinue) { $validChoices += 'C' }
    
    do {
        $choice = Read-Host "`nEnter your choice"
        $choice = $choice.ToUpper()
    } until ($choice -in $validChoices)
    
    return $choice
}

# Function for resource clean up
function Invoke-Cleanup {
    [CmdletBinding()]
    param (
        [switch]$Silent
    )
    
    if (-not $Silent) {
        Write-Host "Cleaning up connections..." -ForegroundColor Yellow
    }
    
    # Suppress verbose output from disconnection commands
    $VerbosePreference = 'SilentlyContinue'
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    $VerbosePreference = 'Continue'
}

#EndRegion

#Region Authentication

# Enhanced tenant selection with validation
function Select-TenantEnvironment {
    $tenants = @{
        "1" = @{
            Name = "Example Organization 1"
            Email = "admin@example1.com"
        }
        "2" = @{
            Name = "Example Organization 2"
            Email = "admin@example2.com"
        }
    }

    Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         Select Tenant Client           ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host

    foreach ($key in $tenants.Keys | Sort-Object) {
        Write-Host "$key. $($tenants[$key].Name)" -ForegroundColor Yellow
    }
    Write-Host "3. Enter Custom Tenant" -ForegroundColor Yellow
    Write-Host "B. Go Back" -ForegroundColor Yellow

    do {
        $selection = Read-Host "Please enter selection (1-3)"
        if ($selection -eq 'B') { return 'B' }
    } until ($selection -in @('1','2','3'))

    if ($selection -eq '3') {
        do {
            $customEmail = Read-Host "Enter tenant Admin email"
            if ($customEmail -eq 'B') { return 'B' }
            if ($customEmail -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$') {
                return $customEmail
            }
            Write-Host "Invalid email format. Please try again." -ForegroundColor Red
        } while ($true)
    }
    
    return $tenants[$selection].Email
}

# Handle Microsoft service connections
function Connect-MicrosoftServices {
    param (
        [Parameter(Mandatory=$false)]
        [string]$AdminEmail,
        
        [Parameter(Mandatory=$false)]
        [string]$OrganizationName,
        
        [Parameter(Mandatory=$false)]
        [string]$Environment = "prod"
    )
    
    try {
        # First check if we have authentication config available
        $useAuthConfig = $false
        if ($OrganizationName) {
            try {
                $authConfig = Get-SecureAuthConfig -OrganizationName $OrganizationName -Environment $Environment
                if ($authConfig) {
                    $useAuthConfig = $true
                }
            }
            catch {
                Write-ProgressStatus -Message "Auth config not found, falling back to interactive login" -Status Warning -Color Yellow
            }
        }

        # Connect to Microsoft Graph
        Write-ProgressStatus -Message "Connecting to Microsoft Graph..." -Status Progress -Color Yellow
        
        # Disconnect existing sessions first
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        
        if ($useAuthConfig) {
            # Use app authentication
            $secureSecret = ConvertTo-SecureString -String $authConfig.ClientSecret -AsPlainText -Force
            $credentials = New-Object System.Management.Automation.PSCredential ($authConfig.ClientId, $secureSecret)
            
            Connect-MgGraph -ClientSecretCredential $credentials -TenantId $authConfig.TenantId -ErrorAction Stop | Out-Null
        }
        else {
            # Use interactive authentication
            $requiredScopes = @(
                "User.ReadWrite.All",
                "GroupMember.ReadWrite.All",
                "Group.ReadWrite.All",
                "Directory.ReadWrite.All",
                "User.EnableDisableAccount.All"
            )
            Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop | Out-Null
        }
        
        Write-ProgressStatus -Message "Microsoft Graph connected successfully" -Status Success -Color Green

        # Connect to Exchange Online
        Write-ProgressStatus -Message "Connecting to Exchange Online..." -Status Progress -Color Yellow
        
        if ($useAuthConfig) {
            # Use certificate-based auth for Exchange Online
            $exchangeParams = @{
                AppId = $authConfig.ClientId
                CertificateThumbprint = $authConfig.CertThumbprint
                Organization = $authConfig.TenantDomain
                ShowBanner = $false
            }
        }
        else {
            # Use interactive auth for Exchange Online
            $exchangeParams = @{
                UserPrincipalName = $AdminEmail
                ShowProgress = $false
                ShowBanner = $false
            }
        }

        Connect-ExchangeOnline @exchangeParams -ErrorAction Stop | Out-Null
        Write-ProgressStatus -Message "Exchange Online connected successfully" -Status Success -Color Green
        Write-Host "`n"

        return $true
    }
    catch {
        Write-ProgressStatus -Message "Connection failed: $($_.Exception.Message)" -Status Failed -Color Red
        return $false
    }
}
function Select-AuthenticationMethod {

    # Intialize config path
    $script:ConfigRootPath = Join-Path $PSScriptRoot "Config"
    if (-not (Test-Path $ConfigRootPath)) {
    New-Item -ItemType Directory -Path $ConfigRootPath -Force | Out-Null
    }

    Clear-Host
    Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║      Select Authentication Method      ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host
    
    Write-Host "1. Use Interactive Login (Username/Password)" -ForegroundColor Yellow
    Write-Host "2. Use Saved Authentication Configuration" -ForegroundColor Yellow
    Write-Host "B. Go Back" -ForegroundColor Yellow
    
    do {
        $choice = Read-Host "Please enter selection (1-2)"
        if ($choice -eq 'B') { return 'B' }
    } while ($choice -notmatch '^[12]$')
    
    if ($choice -eq '1') {
        return @{
            Method = 'Interactive'
            Config = $null
        }
    }
    else {
        # Get available configurations
        $configs = @(Get-AvailableAuthConfigs)
        if ($configs.Count -eq 0) {
            Write-Host "No authentication configurations found. Please create one first." -ForegroundColor Red
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return 'B'
        }
        
        Write-Host "`nAvailable Configurations:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $configs.Count; $i++) {
            Write-Host "$($i + 1). $($configs[$i].Organization) ($($configs[$i].Environment))"
        }
        
        do {
            $configChoice = Read-Host "`nSelect configuration (1-$($configs.Count))"
            if ($configChoice -eq 'B') { return 'B' }
        } while ([int]$configChoice -lt 1 -or [int]$configChoice -gt $configs.Count)
        
        $selectedConfig = $configs[[int]$configChoice - 1]
        return @{
            Method = 'Config'
            Config = @{
                OrganizationName = $selectedConfig.Organization
                Environment = $selectedConfig.Environment
            }
        }
    }
}

#EndRegion

#Region Email/Calendar functions
# Function to handle email delegation and calendar permissions
function EmailDelegation {
    do {
        Clear-Host
        Write-Host "`nEmail Delegation Management" -ForegroundColor Cyan
        
        # Get mailbox to manage (only on first run)
        if (-not $mailboxEmail) {
            $mailboxEmail = Read-Host "Enter email address of mailbox to manage"
        }
        
        try {
            $mailbox = Get-EXOMailbox -Identity $mailboxEmail -ErrorAction Stop
            
            # Show current delegates with numbered list for easy removal
            Write-Host "`nCurrent Delegates:" -ForegroundColor Yellow
            $currentDelegates = Get-MailboxPermission -Identity $mailboxEmail | 
                Where-Object {$_.User -notlike "NT AUTHORITY\*" -and $_.IsInherited -eq $false}
            
            # Also get SendAs permissions
            $sendAsPermissions = Get-RecipientPermission -Identity $mailboxEmail | 
                Where-Object {$_.Trustee -notlike "NT AUTHORITY\*"}
            
            # And SendOnBehalf permissions
            $mailboxDetails = Get-Mailbox -Identity $mailboxEmail
            $sendOnBehalfPermissions = $mailboxDetails.GrantSendOnBehalfTo
            
            if ($currentDelegates -or $sendAsPermissions -or $sendOnBehalfPermissions) {
                $delegateIndex = 1
                $delegateMap = @{}
                
                # Display FullAccess permissions
                $currentDelegates | ForEach-Object {
                    Write-Host "$delegateIndex. $($_.User): FullAccess" -ForegroundColor White
                    $delegateMap[$delegateIndex] = @{
                        Email = $_.User
                        Type = "FullAccess"
                    }
                    $delegateIndex++
                }
                
                # Display SendAs permissions
                $sendAsPermissions | ForEach-Object {
                    Write-Host "$delegateIndex. $($_.Trustee): SendAs" -ForegroundColor White
                    $delegateMap[$delegateIndex] = @{
                        Email = $_.Trustee
                        Type = "SendAs"
                    }
                    $delegateIndex++
                }
                
                # Display SendOnBehalf permissions
                $sendOnBehalfPermissions | ForEach-Object {
                    Write-Host "$delegateIndex. $_ SendOnBehalf" -ForegroundColor White
                    $delegateMap[$delegateIndex] = @{
                        Email = $_
                        Type = "SendOnBehalf"
                    }
                    $delegateIndex++
                }
            } else {
                Write-Host "No delegates found" -ForegroundColor White
            }
            
            # Delegate management options
            Write-Host "`nOptions:" -ForegroundColor Cyan
            Write-Host "A: Add Delegate" -ForegroundColor Yellow
            Write-Host "R: Remove Delegate (enter R followed by number, e.g., R1)" -ForegroundColor Yellow
            Write-Host "C: Change Mailbox" -ForegroundColor Yellow
            Write-Host "B: Go Back" -ForegroundColor Yellow
            
            $choice = Read-Host "`nEnter choice"
            
            switch -Regex ($choice) {
                '^[Aa]$' {
                    # Add new delegate
                    $delegateEmail = Read-Host "Enter delegate's email address"
                    Write-Host "`nAvailable Access Levels:" -ForegroundColor Yellow
                    Write-Host "1: Full Access (can access entire mailbox)" -ForegroundColor White
                    Write-Host "2: Send As (can send as this mailbox)" -ForegroundColor White
                    Write-Host "3: Send On Behalf (can send on behalf of this mailbox)" -ForegroundColor White
                    Write-Host "4: All Access (combines all above permissions)" -ForegroundColor White
                    
                    $accessChoice = Read-Host "Enter access level (1-4)"
                    
                    try {
                        switch ($accessChoice) {
                            '1' {
                                Add-MailboxPermission -Identity $mailboxEmail -User $delegateEmail -AccessRights FullAccess -InheritanceType All
                            }
                            '2' {
                                Add-RecipientPermission -Identity $mailboxEmail -Trustee $delegateEmail -AccessRights SendAs -Confirm:$false
                            }
                            '3' {
                                Set-Mailbox -Identity $mailboxEmail -GrantSendOnBehalfTo @{Add=$delegateEmail}
                            }
                            '4' {
                                Add-MailboxPermission -Identity $mailboxEmail -User $delegateEmail -AccessRights FullAccess -InheritanceType All
                                Add-RecipientPermission -Identity $mailboxEmail -Trustee $delegateEmail -AccessRights SendAs -Confirm:$false
                                Set-Mailbox -Identity $mailboxEmail -GrantSendOnBehalfTo @{Add=$delegateEmail}
                            }
                        }
                        Write-Host "Delegate access granted successfully" -ForegroundColor Green
                        Start-Sleep -Seconds 2
                    }
                    catch {
                        Write-Host "Error granting delegate access: $_" -ForegroundColor Red
                        Start-Sleep -Seconds 3
                    }
                }
                '^[Rr](\d+)$' {
                    # Remove delegate based on number
                    $index = $matches[1]
                    if ($delegateMap.ContainsKey([int]$index)) {
                        $delegateInfo = $delegateMap[[int]$index]
                        try {
                            switch ($delegateInfo.Type) {
                                "FullAccess" {
                                    Remove-MailboxPermission -Identity $mailboxEmail -User $delegateInfo.Email -AccessRights FullAccess -Confirm:$false
                                }
                                "SendAs" {
                                    Remove-RecipientPermission -Identity $mailboxEmail -Trustee $delegateInfo.Email -AccessRights SendAs -Confirm:$false
                                }
                                "SendOnBehalf" {
                                    Set-Mailbox -Identity $mailboxEmail -GrantSendOnBehalfTo @{Remove=$delegateInfo.Email}
                                }
                            }
                            Write-Host "Delegate access removed successfully" -ForegroundColor Green
                            Start-Sleep -Seconds 2
                        }
                        catch {
                            Write-Host "Error removing delegate access: $_" -ForegroundColor Red
                            Start-Sleep -Seconds 3
                        }
                    }
                    else {
                        Write-Host "Invalid delegate number" -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
                '^[Cc]$' {
                    $mailboxEmail = $null  # Reset mailbox to prompt for new one
                }
                '^[Bb]$' {
                    return
                }
                default {
                    Write-Host "Invalid choice" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                }
            }
        }
        catch {
            Write-Host "Error accessing mailbox: $_" -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
    } while ($true)
}
function ManageCalendarPermission {
    do {
        Clear-Host
        Write-Host "`nCalendar Permissions Management" -ForegroundColor Cyan
        
        # Get mailbox to manage (only on first run)
        if (-not $mailboxEmail) {
            $mailboxEmail = Read-Host "Enter email address of calendar to manage"
            
            # Validate email format
            if (-not ($mailboxEmail -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')) {
                Write-Host "Invalid email format" -ForegroundColor Red
                Start-Sleep -Seconds 2
                continue
            }
        }
        
        try {
            # First verify the mailbox exists
            Write-Host "Verifying mailbox..." -ForegroundColor Yellow
            $mailbox = Get-EXOMailbox -Identity $mailboxEmail -ErrorAction Stop
            if (-not $mailbox) {
                throw "Mailbox not found"
            }

            # Construct the calendar path correctly
            $calendarPath = "$($mailboxEmail):\Calendar"  # Simplified path construction
            Write-Host "Using calendar path: $calendarPath" -ForegroundColor Gray
            
            # Show current permissions with numbered list for easy removal
            Write-Host "`nCurrent Calendar Permissions:" -ForegroundColor Yellow
            $currentPermissions = Get-EXOMailboxFolderPermission -Identity $calendarPath
            
            if ($currentPermissions) {
                $permissionIndex = 1
                $permissionMap = @{}
                
                $currentPermissions | ForEach-Object {
                    Write-Host "$permissionIndex. $($_.User.DisplayName): $($_.AccessRights -join ', ')" -ForegroundColor White
                    $permissionMap[$permissionIndex] = @{
                        User = $_.User.DisplayName
                        AccessRights = $_.AccessRights
                    }
                    $permissionIndex++
                }
            } else {
                Write-Host "No custom permissions found" -ForegroundColor White
            }
            
            # Permission management options
            Write-Host "`nOptions:" -ForegroundColor Cyan
            Write-Host "A: Add Calendar Permission" -ForegroundColor Yellow
            Write-Host "R: Remove Permission (enter R followed by number, e.g., R1)" -ForegroundColor Yellow
            Write-Host "C: Change Calendar" -ForegroundColor Yellow
            Write-Host "B: Go Back" -ForegroundColor Yellow
            
            $choice = Read-Host "`nEnter choice"
            
            switch -Regex ($choice) {
                '^[Aa]$' {
                    $userEmail = Read-Host "Enter user's email address"
                    
                    # Display detailed access level descriptions
                    Write-Host "`nAvailable Access Levels:" -ForegroundColor Yellow
                    Write-Host "1: Owner" -ForegroundColor White
                    Write-Host "   - Full control of the calendar" -ForegroundColor Gray
                    Write-Host "   - Can create, read, edit, and delete all items" -ForegroundColor Gray
                    Write-Host "   - Can create and delete subfolders" -ForegroundColor Gray
                    
                    Write-Host "`n2: Publishing Editor" -ForegroundColor White
                    Write-Host "   - Can create, read, edit, and delete all items" -ForegroundColor Gray
                    Write-Host "   - Cannot change permissions" -ForegroundColor Gray
                    
                    Write-Host "`n3: Editor" -ForegroundColor White
                    Write-Host "   - Can create, read, edit, and delete all items" -ForegroundColor Gray
                    
                    Write-Host "`n4: Publishing Author" -ForegroundColor White
                    Write-Host "   - Can create and read items" -ForegroundColor Gray
                    Write-Host "   - Can edit and delete only items they create" -ForegroundColor Gray
                    Write-Host "   - Can create subfolders" -ForegroundColor Gray
                    
                    Write-Host "`n5: Author" -ForegroundColor White
                    Write-Host "   - Can create and read items" -ForegroundColor Gray
                    Write-Host "   - Can edit and delete only items they create" -ForegroundColor Gray
                    
                    Write-Host "`n6: Reviewer" -ForegroundColor White
                    Write-Host "   - Can read items only" -ForegroundColor Gray
                    
                    $accessLevel = Read-Host "`nEnter access level (1-6)"
                    
                    $accessRight = switch ($accessLevel) {
                        '1' { 'Owner' }
                        '2' { 'PublishingEditor' }
                        '3' { 'Editor' }
                        '4' { 'PublishingAuthor' }
                        '5' { 'Author' }
                        '6' { 'Reviewer' }
                        default { 'Editor' }
                    }
                    
                    try {
                        # Check if permission already exists
                        $existingPermission = Get-EXOMailboxFolderPermission -Identity $calendarPath -User $userEmail -ErrorAction SilentlyContinue
                        
                        if ($existingPermission) {
                            Set-MailboxFolderPermission -Identity $calendarPath -User $userEmail -AccessRights $accessRight
                            Write-Host "Calendar permission updated successfully" -ForegroundColor Green
                        } else {
                            Add-MailboxFolderPermission -Identity $calendarPath -User $userEmail -AccessRights $accessRight
                            Write-Host "Calendar permission granted successfully" -ForegroundColor Green
                        }
                        Start-Sleep -Seconds 2
                    }
                    catch {
                        Write-Host "Error setting calendar permission: $_" -ForegroundColor Red
                        Start-Sleep -Seconds 3
                    }
                }
                '^[Rr](\d+)$' {
                    # Remove permission based on number
                    $index = $matches[1]
                    if ($permissionMap.ContainsKey([int]$index)) {
                        $permissionInfo = $permissionMap[[int]$index]
                        try {
                            Remove-MailboxFolderPermission -Identity $calendarPath -User $permissionInfo.User -Confirm:$false
                            Write-Host "Calendar permission removed successfully" -ForegroundColor Green
                            Start-Sleep -Seconds 2
                        }
                        catch {
                            Write-Host "Error removing calendar permission: $_" -ForegroundColor Red
                            Start-Sleep -Seconds 3
                        }
                    }
                    else {
                        Write-Host "Invalid permission number" -ForegroundColor Red
                        Start-Sleep -Seconds 2
                    }
                }
                '^[Cc]$' {
                    $mailboxEmail = $null  # Reset mailbox to prompt for new one
                }
                '^[Bb]$' {
                    return
                }
                default {
                    Write-Host "Invalid choice" -ForegroundColor Red
                    Start-Sleep -Seconds 2
                }
            }
        }
        catch {
            Write-Host "Error accessing calendar: $_" -ForegroundColor Red
            Start-Sleep -Seconds 3
            return
        }
    } while ($true)
}
function Convert-ToSharedMailbox {
    Clear-Host
    Write-Host "`nConvert to Shared Mailbox" -ForegroundColor Cyan
    
    # Get mailbox to convert
    $mailboxEmail = Read-Host "Enter email address of mailbox to convert"
    
    try {
        $mailbox = Get-EXOMailbox -Identity $mailboxEmail -ErrorAction Stop
        
        if ($mailbox.RecipientTypeDetails -eq "SharedMailbox") {
            Write-Host "This mailbox is already a shared mailbox" -ForegroundColor Yellow
            return
        }
        
        # Confirm conversion
        Write-Host "`nWARNING: Converting to a shared mailbox will:" -ForegroundColor Yellow
        Write-Host "- Remove the need for a license" -ForegroundColor Yellow
        Write-Host "- Keep all existing email and calendar items" -ForegroundColor Yellow
        Write-Host "- Maintain all existing permissions" -ForegroundColor Yellow
        
        $confirm = Read-Host "`nDo you want to proceed? (Y/N)"
        
        if ($confirm -eq 'Y') {
            Set-Mailbox -Identity $mailboxEmail -Type Shared
            Write-Host "Mailbox converted to shared successfully" -ForegroundColor Green
            
            # Option to add delegate
            $addDelegate = Read-Host "`nDo you want to add a delegate to this shared mailbox? (Y/N)"
            if ($addDelegate -eq 'Y') {
                $delegateEmail = Read-Host "Enter delegate's email address"
                Add-MailboxPermission -Identity $mailboxEmail -User $delegateEmail -AccessRights FullAccess -InheritanceType All
                Write-Host "Delegate access granted successfully" -ForegroundColor Green
            }
        }
    }
    catch {
        Write-Host "Error converting mailbox: $_" -ForegroundColor Red
    }
}
function Show-EmailCalendarMenu {
    Write-Host "`nEmail & Calendar Management Options:" -ForegroundColor Cyan
    Write-Host "1: Manage Email Delegation" -ForegroundColor Yellow
    Write-Host "2: Manage Calendar Permissions" -ForegroundColor Yellow
    Write-Host "3: Convert to Shared Mailbox" -ForegroundColor Yellow
    Write-Host "4: Return to Main Menu" -ForegroundColor Yellow
    Write-Host "════════════════════════" -ForegroundColor Cyan
    
    do {
        $choice = Read-Host "`nEnter your choice (1-4)"
    } while ($choice -notmatch '^[1-4]$')
    
    return $choice
}
function Start-EmailCalendarManagement {
    do {
        Clear-Host
        
        # Show welcome banner
        Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║    Email & Calendar Management Tool    ║" -ForegroundColor Cyan
        Write-Host "║             Version 3.1                ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host

        try {
            # Select authentication method
            $authMethod = Select-AuthenticationMethod
            if ($authMethod -eq 'B') { return }
            
            if ($authMethod.Method -eq 'Interactive') {
                # Select tenant and connect interactively
                $adminEmail = Select-TenantEnvironment
                if ($adminEmail -eq 'B') { return }
                
                if (-not (Connect-MicrosoftServices -AdminEmail $adminEmail)) {
                    throw "Failed to establish required connections"
                }
            }
            else {
                # Connect using saved authentication config
                if (-not (Connect-MicrosoftServices -OrganizationName $authMethod.Config.OrganizationName -Environment $authMethod.Config.Environment)) {
                    throw "Failed to establish required connections"
                }
            }

            # Show submenu for email and calendar management
            Clear-Host
            $choice = Show-EmailCalendarMenu
            
            switch ($choice) {
                '1' { EmailDelegation }
                '2' { ManageCalendarPermission }
                '3' { Convert-ToSharedMailbox }
                '4' { return } # Return to main menu
            }
        }
        catch {
            Write-Host "Critical error: $_" -ForegroundColor Red
            Show-NavigationOptions -AllowMainMenu
            $choice = Get-NavigationChoice -AllowMainMenu
            if ($choice -eq 'M') { return }
            if ($choice -eq 'X') { exit }
        }
        finally {
            # Cleanup connections
            Invoke-Cleanup
        }
    } while ($true)
}
#EndRegion

#Region Offboarding functions

# Main function to handle user offboarding
function Remove-UserAccess {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName
    )
    
    try {
        # Initialize offboarding
        Write-ProgressStatus -Message "Starting offboarding process for $UserPrincipalName" -Status Progress -Color Cyan
        
        # Get user object
        $user = Invoke-WithRetry { 
            Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        }
        
        if (-not $user) {
            throw "User not found: $UserPrincipalName"
        }

        # Process group memberships
        Write-ProgressStatus -Message "Processing group memberships..." -Status Progress -Color Yellow
        $groupResults = Remove-UserGroupMemberships -UserId $user.Id

        # Process licenses
        Write-ProgressStatus -Message "Processing licenses..." -Status Progress -Color Yellow
        $licenseResults = Remove-UserLicenses -UserId $user.Id

        # Process mailbox
        Write-ProgressStatus -Message "Processing mailbox, please wait..." -Status Progress -Color Yellow
        $mailboxResults = Convert-UserMailbox -UserPrincipalName $UserPrincipalName

        # Disable user account and revoke sessions
        Write-ProgressStatus -Message "Disabling user account..." -Status Progress -Color Yellow
        $accountResults = Disable-UserAccount -UserId $user.Id

        # Remove user from GAL
        Write-ProgressStatus -Message "Removing user from Global Address List..." -Status Progress -Color Yellow
        $galResults = Remove-UserFromGAL -UserId $user.Id -UserPrincipalName $UserPrincipalName

        # Construct the summary
        $summary = @{
            DisplayName        = $user.DisplayName
            Email              = $user.UserPrincipalName
            GroupsRemoved      = if ($groupResults.RemovedGroups.Count -gt 0) { $groupResults.RemovedGroups -join ", " } else { "None" }
            FailedGroups       = if ($groupResults.FailedGroups.Count -gt 0) { $groupResults.FailedGroups -join ", " } else { "None" }
            LicensesRemoved    = if ($licenseResults.RemovedLicenses.Count -gt 0) { $licenseResults.RemovedLicenses -join ", " } else { "None" }
            MailboxConverted   = $mailboxResults.Status
            DelegateAssigned   = if ($mailboxResults.DelegateAccess) { $mailboxResults.Delegate } else { "None" }
            AccountDisabled    = if ($accountResults.Disabled) { "Yes" } else { "No" }
            SessionsRevoked    = if ($accountResults.SessionsRevoked) { "Yes" } else { "No" }
            OffboardingStatus  = "Completed"
            Timestamp          = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }

        # Save the summary to JSON
        $basePath = Join-Path -Path $PSScriptRoot -ChildPath "Reports/Offboarding"
        if (-not (Test-Path -Path $basePath)) {
            New-Item -ItemType Directory -Path $basePath -Force | Out-Null
        }

        $jsonFilePath = Join-Path -Path $basePath -ChildPath "$($user.UserPrincipalName)_OffboardingReport.json"
        $summary | ConvertTo-Json -Depth 3 | Set-Content -Path $jsonFilePath

        # Output the summary to the terminal
        Write-Host "`n═══════ Offboarding Summary ═══════" -ForegroundColor Cyan
        Write-Host "Name: $($summary.DisplayName)"
        Write-Host "Email: $($summary.Email)"
        Write-Host "Groups Removed: $($summary.GroupsRemoved)"
        Write-Host "Skipped Groups: $($summary.FailedGroups)"
        Write-Host "Licenses Removed: $($summary.LicensesRemoved)"
        Write-Host "Mailbox Converted: $($summary.MailboxConverted)"
        Write-Host "Delegate Assigned: $($summary.DelegateAssigned)"
        Write-Host "Account Disabled: $($summary.AccountDisabled)"
        Write-Host "Sessions Revoked: $($summary.SessionsRevoked)"
        Write-Host "Offboarding Status: $($summary.OffboardingStatus)"
        Write-Host "Timestamp: $($summary.Timestamp)"
        Write-Host "════════════════════════════════════" -ForegroundColor Cyan

        Write-Host "Offboarding completed successfully. Summary exported to $jsonFilePath" -ForegroundColor Green

        return $summary
    }
    catch {
        Write-ProgressStatus -Message "Critical error during offboarding: $_" -Status Failed -Color Red
        return $false
    }
}

# Function to remove group memberships
function Remove-UserGroupMemberships {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserId
    )

    $removedGroups = @()
    $failedGroups = @()

    try {
        # Get the user object
        $user = Get-MgUser -UserId $UserId
        if (-not $user) {
            throw "User not found"
        }

        # Get all groups the user is a member of
        $userGroups = Get-MgUserMemberOf -UserId $UserId -All
        if (-not $userGroups) {
            Write-Host "User is not a member of any groups." -ForegroundColor Yellow
            return @{
                RemovedGroups = $removedGroups
                FailedGroups  = $failedGroups
            }
        }

        # Process groups (Microsoft 365, Security, and Distribution Lists)
        foreach ($group in $userGroups) {
            $groupName = $group.AdditionalProperties['displayName']
            $isUnified = $group.AdditionalProperties['groupTypes'] -contains "Unified"
            $isDynamic = $group.AdditionalProperties['groupTypes'] -contains "DynamicMembership"
            $isSecurity = $group.AdditionalProperties['securityEnabled'] -eq $true
            $isMailEnabled = $group.AdditionalProperties['mailEnabled'] -eq $true

            try {
                # Skip dynamic groups
                if ($isDynamic) {
                    Write-Host "Skipping dynamic group: $groupName" -NoNewLine
                    $failedGroups += $groupName
                    continue
                }

                # Remove from Microsoft 365 or Security Groups
                if ($isUnified -or $isSecurity) {
                    Write-Host "Removing from group: $groupName" -NoNewline
                    Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $UserId -ErrorAction Stop
                    Write-Host " - Success" -ForegroundColor Green
                    $removedGroups += $groupName
                }
                # Remove from Distribution Groups
                elseif ($isMailEnabled -and (-not $isSecurity)) {
                    Write-Host "Removing from distribution group: $groupName" -NoNewline
                    $mailNickname = $group.AdditionalProperties['mailNickname']
                    Remove-DistributionGroupMember -Identity $mailNickname -Member $UserId -Confirm:$false -ErrorAction Stop
                    Write-Host " - Success" -ForegroundColor Green
                    $removedGroups += $groupName
                }
            }
            catch {
                Write-Host " - Failed" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                $failedGroups += $groupName
            }
        }
        Write-Host "`n"

        # Return a summary of results
        return @{
            RemovedGroups = $removedGroups
            FailedGroups  = $failedGroups
        }
    }
    catch {
        Write-Host "Critical error occurred: $_" -ForegroundColor Red
        throw
    }
}

# Function to remove licenses
function Remove-UserLicenses {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserId
    )
    
    $results = @{
        TotalLicenses   = 0
        RemovedLicenses = @()
        FailedLicenses  = @()
    }

    try {
        $licenses = Get-MgUserLicenseDetail -UserId $UserId
        $results.TotalLicenses = $licenses.Count

        foreach ($license in $licenses) {
            try {
                $skuPartNumber = $license.SkuPartNumber
                Set-MgUserLicense -UserId $UserId -BodyParameter @{
                    RemoveLicenses = @($license.SkuId)
                    AddLicenses    = @()
                }
                $results.RemovedLicenses += $skuPartNumber
            }
            catch {
                $results.FailedLicenses += $license.SkuPartNumber
            }
        }
    }
    catch {
        Write-Host "Error removing licenses: $_" -ForegroundColor Red
    }

    return $results
}

# Function to convert mailbox
function Convert-UserMailbox {
    param (
        [string]$UserPrincipalName
    )
    
    $results = @{
        Status = "Unknown"
        DelegateAccess = $false
        Delegate = $null
    }

    try {
        $mailbox = Get-EXOMailbox -Identity $UserPrincipalName -ErrorAction Stop
        
        if ($mailbox.RecipientTypeDetails -ne "SharedMailbox") {
            Set-Mailbox -Identity $UserPrincipalName -Type Shared -Confirm:$false
            $results.Status = "Converted"
        }
        else {
            $results.Status = "AlreadyShared"
        }

        $addDelegate = Read-Host "Do you want to assign delegate access? (Y/N)"
        if ($addDelegate -eq 'Y') {
            do {
                $delegate = Read-Host "Enter delegate email"
                try {
                    Add-MailboxPermission -Identity $UserPrincipalName -User $delegate -AccessRights FullAccess -InheritanceType All
                    $results.DelegateAccess = $true
                    $results.Delegate = $delegate
                    break
                }
                catch {
                    Write-Host "Invalid delegate email. Please try again." -ForegroundColor Red
                }
            } while ($true)
        }
    }
    catch {
        $results.Status = "Failed"
        Write-ProgressStatus -Message "Error processing mailbox: $_" -Status Failed -Color Red
    }

    return $results
}

# Function to disable user account
function Disable-UserAccount {
    param (
        [string]$UserId
    )
    
    $results = @{
        Disabled = $false
        SessionsRevoked = $false
    }

    try {
        # Revoke all sessions
        Revoke-MgUserSignInSession -UserId $UserId
        $results.SessionsRevoked = $true

        # Disable account
        Update-MgUser -UserId $UserId -BodyParameter @{
            AccountEnabled = $false
        }
        $results.Disabled = $true
    }
    catch {
        Write-ProgressStatus -Message "Error disabling account: $_" -Status Failed -Color Red
    }

    return $results
}

# Function to hide user from GAL
function Remove-UserFromGAL {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserId,
        [string]$UserPrincipalName
    )
    
    $results = @{
        HiddenFromGAL = $false
        ProcessingStatus = "Not Started"
    }

    try {
        
        # Update mailbox to hide from GAL
        Set-Mailbox -Identity $UserPrincipalName -HiddenFromAddressListsEnabled $true
        
        # Verify the change
        $mailbox = Get-Mailbox -Identity $UserPrincipalName
        if ($mailbox.HiddenFromAddressListsEnabled) {
            $results.HiddenFromGAL = $true
            $results.ProcessingStatus = "Completed"
            Write-ProgressStatus -Message "Successfully removed $UserPrincipalName from GAL" -Status Success -Color Green
        } else {
            $results.ProcessingStatus = "Failed"
            Write-ProgressStatus -Message "Failed to verify GAL removal for $UserPrincipalName" -Status Failed -Color Red
        }
    }
    catch {
        $results.ProcessingStatus = "Error"
        Write-ProgressStatus -Message "Error removing user from GAL: $_" -Status Failed -Color Red
    }

    return $results
}

#EndRegion

#Region Onboarding Functions
# Main function for handling user generation
function Add-UserAccess {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$UserInfo
    )

    try {
        # Create user account
        Write-Host "Creating user account..." -ForegroundColor Yellow
        $newUser = New-AzureUserAccount -UserInfo $UserInfo
        if (-not $newUser) {
            throw "Failed to create user account"
        }
        $initialPassword = $newUser.Password
        $userObj = $newUser.User

        # Process license
        Write-Host "Processing licenses..." -ForegroundColor Yellow
        $licenseResults = Set-AzureUserLicense -UserId $userObj.Id

        # Process groups
        Write-Host "Processing group memberships..." -ForegroundColor Yellow
        $groupResults = Add-AzureUserToGroups -UserId $userObj.Id

        # Prepare summary
        $summary = @{
            DisplayName    = $userObj.DisplayName
            Email          = $userObj.UserPrincipalName
            InitialPassword = $initialPassword
            AssignedLicense = $licenseResults -join ", "
            AssignedGroups  = $groupResults -join ", "
            OfficeLocation  = $UserInfo.OfficeLocation
        }

        # Define folder path and create if it doesn't exist
        $basePath = Join-Path -Path $PSScriptRoot -ChildPath "Reports/Onboarding"
        if (-not (Test-Path -Path $basePath)) {
            New-Item -ItemType Directory -Path $basePath -Force | Out-Null
        }

        # Define JSON file path and export summary
        $jsonFilePath = Join-Path -Path $basePath -ChildPath "$($userObj.UserPrincipalName)_OnboardingReport.json"
        $summary | ConvertTo-Json -Depth 3 | Set-Content -Path $jsonFilePath

        # Output summary to terminal
        Write-Host "`n═══════ Onboarding Summary ═══════" -ForegroundColor Cyan
        Write-Host "Name: $($summary.DisplayName)"
        Write-Host "Email: $($summary.Email)"
        Write-Host "Initial Password: $($summary.InitialPassword)"
        Write-Host "Assigned License(s): $($summary.AssignedLicense)"
        Write-Host "Assigned Group(s): $($summary.AssignedGroups)"
        Write-Host "Office Location: $($summary.OfficeLocation)"
        Write-Host "════════════════════════════════════" -ForegroundColor Cyan

        Write-Host "Onboarding completed successfully. Summary exported to $jsonFilePath" -ForegroundColor Green

        return $summary
    }
    catch {
        Write-Host "Critical error during onboarding: $_" -ForegroundColor Red
        return $false
    }
}

# Function to generate user
function New-AzureUserAccount {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$UserInfo
    )
    
    try {

        # Display header for email format requirements
        Write-Host "NOTE: Ensure organizational email formats are maintained" -ForegroundColor Yellow

        # Generate password
        Add-Type -AssemblyName System.Web
        $passwordProfile = @{
            forceChangePasswordNextSignIn = $true
            password = [System.Web.Security.Membership]::GeneratePassword(16, 3)
        }

        # Create user parameters
        $userParams = @{
            AccountEnabled = $true
            DisplayName = "$($UserInfo.GivenName) $($UserInfo.Surname)"
            MailNickname = "$($UserInfo.GivenName).$($UserInfo.Surname)".ToLower()
            UserPrincipalName = $UserInfo.Email
            GivenName = $UserInfo.GivenName
            Surname = $UserInfo.Surname
            JobTitle = $UserInfo.JobTitle
            UsageLocation = "AU"
            PasswordProfile = $passwordProfile
            BusinessPhones = $UserInfo.BusinessPhones
            OfficeLocation = $UserInfo.OfficeLocation
            StreetAddress = $UserInfo.StreetAddress
            PostalCode = $UserInfo.PostalCode
            Country = $UserInfo.Country
        }

        # Validate and log parameters
        foreach ($key in $userParams.Keys) {
            if (-not $userParams[$key]) {
                Write-Host "Warning: Parameter '$key' is empty or null." -ForegroundColor Yellow
            }
        }
        
        # Create user
        $newUser = New-MgUser @userParams
        
        if ($newUser) {
            Write-Host "User account created successfully" -ForegroundColor Green
            return @{
                User = $newUser
                Password = $passwordProfile.password
            }
        } else {
            throw "User creation returned null"
        }
    }
    catch {
        Write-Host "Error creating user: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Function to assign licenses
function Set-AzureUserLicense {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        if ([string]::IsNullOrEmpty($UserId)) {
            Write-Host "UserId is empty, cannot process licenses" -ForegroundColor Red
            return $null
        }

        # Get all available licenses with their details
        $allLicenses = Get-MgSubscribedSku
        if (-not $allLicenses) {
            Write-Host "No licenses found in the tenant" -ForegroundColor Yellow
            return $null
        }

        $selectedLicenses = @()
        
        do {
            Clear-Host
            # Get and display current licenses
            $currentLicenses = Get-MgUser -UserId $UserId -Property AssignedLicenses | 
                Select-Object -ExpandProperty AssignedLicenses

            Write-Host "`nCurrently Assigned Licenses:" -ForegroundColor Cyan
            if ($currentLicenses.Count -gt 0) {
                foreach ($license in $currentLicenses) {
                    $licenseName = ($allLicenses | Where-Object { $_.SkuId -eq $license.SkuId }).SkuPartNumber
                    Write-Host "- $licenseName" -ForegroundColor Green
                }
            } else {
                Write-Host "None" -ForegroundColor Yellow
            }

            # Display available licenses in a single column
            Write-Host "`nAvailable Licenses:" -ForegroundColor Yellow
            
            # Find the longest license name for padding
            $maxNameLength = ($allLicenses | ForEach-Object { $_.SkuPartNumber.Length } | Measure-Object -Maximum).Maximum
            
            # Display each license on its own line
            for ($idx = 0; $idx -lt $allLicenses.Count; $idx++) {
                $license = $allLicenses[$idx]
                $availableUnits = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
                
                # Pad the index number to 3 digits
                $paddedIndex = "{0:D3}" -f ($idx + 1)
                
                # Create padded license name
                $paddedName = $license.SkuPartNumber.PadRight($maxNameLength)
                
                # Format the complete line
                Write-Host ("{0}. {1} (Available: {2})" -f $paddedIndex, $paddedName, $availableUnits) -ForegroundColor White
            }

            Write-Host "`nCommands:" -ForegroundColor Yellow
            Write-Host "- Enter a number to add that license" -ForegroundColor Green
            Write-Host "- Type 'r' followed by a number to remove that license (e.g., 'r012')" -ForegroundColor Red
            Write-Host "- Type 'done' to finish" -ForegroundColor Cyan

            $choice = Read-Host "`nEnter command"
            
            if ($choice -eq 'done') {
                break
            }
            
            # Handle remove command
            if ($choice -match '^r(\d+)$') {
                $index = [int]$matches[1] - 1
                if ($index -ge 0 -and $index -lt $allLicenses.Count) {
                    $licenseToRemove = $allLicenses[$index]
                    if ($currentLicenses.SkuId -contains $licenseToRemove.SkuId) {
                        Set-MgUserLicense -UserId $UserId -RemoveLicenses @($licenseToRemove.SkuId) -AddLicenses @()
                        Write-Host "Removed license: $($licenseToRemove.SkuPartNumber)" -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                    } else {
                        Write-Host "This license is not currently assigned" -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                    }
                }
                continue
            }

            # Handle add command
            if ([int]::TryParse($choice, [ref]$null)) {
                $index = [int]$choice - 1
                if ($index -ge 0 -and $index -lt $allLicenses.Count) {
                    $selectedLicense = $allLicenses[$index]
                    
                    # Check if license is already assigned
                    if ($currentLicenses.SkuId -contains $selectedLicense.SkuId) {
                        Write-Host "License $($selectedLicense.SkuPartNumber) is already assigned" -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                    } else {
                        # Add license
                        Set-MgUserLicense -UserId $UserId -AddLicenses @(@{ SkuId = $selectedLicense.SkuId }) -RemoveLicenses @()
                        Write-Host "Added license: $($selectedLicense.SkuPartNumber)" -ForegroundColor Green
                        Start-Sleep -Seconds 1
                    }
                }
            }
        } while ($true)

        return $true
    }
    catch {
        Write-Host "Error in license management: $_" -ForegroundColor Red
        Start-Sleep -Seconds 2
        return $false
    }
}

# Function to delegate user groups
function Add-AzureUserToGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )

    try {
        # Initialize array to track assigned groups
        $assignedGroups = @()
        
        # Get initial group assignments
        $currentGroups = Get-MgUserMemberOf -UserId $UserId | 
            Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
        $assignedGroups = $currentGroups | Select-Object -ExpandProperty DisplayName
        
        # Get all available groups and sort them
        $groups = Get-MgGroup -Top 100 | Select-Object DisplayName, Id | Sort-Object DisplayName
        
        do {
            Clear-Host
            # Display current group assignments
            Write-Host "`nCurrently Selected Groups:" -ForegroundColor Cyan
            if ($assignedGroups.Count -gt 0) {
                Write-Host ($assignedGroups -join ", ") -ForegroundColor Green
                }
             else {
                Write-Host "None" -ForegroundColor Yellow
            }

            # Display available groups
            Write-Host "`nAvailable Groups:" -ForegroundColor Yellow
            $groupMap = @{}
            
            # Find the longest group name for padding
            $maxLength = ($groups | Measure-Object -Property DisplayName -Maximum).Maximum.Length
            
            # Calculate number of columns
            $columnWidth = 40
            $windowWidth = $Host.UI.RawUI.WindowSize.Width
            $numColumns = [math]::Max(1, [math]::Floor($windowWidth / $columnWidth))
            
            # Display groups in columns
            for ($idx = 0; $idx -lt $groups.Count; $idx++) {
                $group = $groups[$idx]
                $paddedIndex = "{0:D3}" -f ($idx + 1)
                $entry = "{0}. {1}" -f $paddedIndex, $group.DisplayName
                $paddedEntry = $entry.PadRight($columnWidth)
                
                # Add to groupMap
                $groupMap[$idx + 1] = @{
                    Id = $group.Id
                    Name = $group.DisplayName
                }
                
                # Write the entry
                Write-Host $paddedEntry -NoNewline -ForegroundColor White
                
                if (($idx + 1) % $numColumns -eq 0 -or $idx -eq ($groups.Count - 1)) {
                    Write-Host ""
                }
            }

            Write-Host "`nCommands:" -ForegroundColor Yellow
            Write-Host "- Enter a number to add that group" -ForegroundColor Green
            Write-Host "- Type 'r' followed by a number to remove from that group (e.g., 'r1')" -ForegroundColor Red
            Write-Host "- Type 'done' to finish" -ForegroundColor Cyan

            $choice = Read-Host "`nEnter command"
            
            if ($choice -eq 'done') {
                break
            }
            
            # Handle remove command
            if ($choice -match '^r(\d+)$') {
                $index = [int]$matches[1] - 1
                if ($index -ge 0 -and $index -lt $groups.Count) {
                    $groupToRemove = $groups[$index]
                    if ($assignedGroups -contains $groupToRemove.DisplayName) {
                        try {
                            Remove-MgGroupMemberByRef -GroupId $groupToRemove.Id -DirectoryObjectId $UserId
                            Write-Host "Removed from group: $($groupToRemove.DisplayName)" -ForegroundColor Yellow
                            $assignedGroups = $assignedGroups | Where-Object { $_ -ne $groupToRemove.DisplayName }
                            Start-Sleep -Seconds 1
                        }
                        catch {
                            Write-Host "Error removing from group: $_" -ForegroundColor Red
                            Start-Sleep -Seconds 1
                        }
                    } else {
                        Write-Host "User is not a member of this group: $($groupToRemove.DisplayName)" -ForegroundColor Yellow
                        Start-Sleep -Seconds 1
                    }
                }
                continue
            }

            # Handle add command
            if ([int]::TryParse($choice, [ref]$null)) {
                $index = [int]$choice - 1
                if ($index -ge 0 -and $index -lt $groups.Count) {
                    $selectedGroup = $groups[$index]
                    if ($assignedGroups -contains $selectedGroup.DisplayName) {
                        Write-Host "User is already a member of $($selectedGroup.DisplayName)" -ForegroundColor Yellow
                    } else {
                        New-MgGroupMember -GroupId $selectedGroup.Id -DirectoryObjectId $UserId
                        Write-Host "Added to $($selectedGroup.DisplayName) successfully" -ForegroundColor Green
                        $assignedGroups += $selectedGroup.DisplayName
                    }
                }
            }
        } while ($true)
        
        return $assignedGroups
    }
    catch {
        Write-Host "Error managing groups: $_" -ForegroundColor Red
        throw
    }
}

#EndRegion

#Region Boarding Execution
# Main offboarding function
function Get-UserForOffboarding {
    param (
        [int]$MaxAttempts = 5,
        [int]$DelaySeconds = 2
    )
    
    $attempts = 0
    
    do {
        $attempts++
        $confirmation = Read-Host "Ensure backups are removed before continuing [Y/N]"
        if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
            Write-Host "Proceeding...`n`n" -ForegroundColor Green
        } else {
            Write-Host "Operation cancelled" -ForegroundColor Yellow
            Show-NavigationOptions -AllowMainMenu -AllowContinue
                $choice = Get-NavigationChoice -AllowMainMenu -AllowContinue
                if ($choice -eq 'M') { return }
        }
        $userEmail = Read-Host "Enter email address to offboard (or 'B' to go back)"
        
        if ($userEmail -eq 'B') {
            return $null
        }
        
        if (-not ($userEmail -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$')) {
            Write-Host "Invalid email format. Please try again." -ForegroundColor Red
            continue
        }
        
        try {
            Write-Host "`nSearching for user..." -ForegroundColor Yellow
            $user = Get-MgUser -UserId $userEmail -ErrorAction Stop
            
            # Display comprehensive user information
            Write-Host "`nUser Details:" -ForegroundColor Cyan
            Write-Host "----------------------------------------" -ForegroundColor Cyan
            Write-Host "Name: $($user.DisplayName)" -ForegroundColor White
            Write-Host "Email: $($user.UserPrincipalName)" -ForegroundColor White
            Write-Host "Title: $($user.JobTitle)" -ForegroundColor White
            Write-Host "Department: $($user.Department)" -ForegroundColor White
            Write-Host "Office Location: $($user.OfficeLocation)" -ForegroundColor White
            Write-Host "----------------------------------------" -ForegroundColor Cyan
            
            # Confirm offboarding
            Write-Host "`nWARNING: This will remove the user's access and licenses." -ForegroundColor Yellow
            $confirm = Read-Host "Are you sure you want to offboard this user? (Y/N)"
            
            if ($confirm -eq 'Y') {
                return $user
            }
            else {
                Write-Host "Operation cancelled by user." -ForegroundColor Yellow
                return $null
            }
        }
        catch {
            $errorMsg = if ($_.Exception.Message -like "*Resource '*' does not exist*") {
                "User not found in Azure AD. Please verify the email address."
            } else {
                "Error accessing user information: $_"
            }
            Write-Host "`n$errorMsg" -ForegroundColor Red
            
            if ($attempts -lt $MaxAttempts) {
                Write-Host "Attempts remaining: $($MaxAttempts - $attempts)" -ForegroundColor Yellow
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    } while ($attempts -lt $MaxAttempts)
    
    Write-Host "`nMaximum attempts exceeded. Returning to main menu." -ForegroundColor Red
    return $null
}
function Start-OffboardingProcess {
    do {
        Clear-Host
        
        # Show welcome banner
        Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║    Azure User Offboarding Assistant    ║" -ForegroundColor Cyan
        Write-Host "║             Version 3.1                ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host

        try {
            # Select authentication method
            $authMethod = Select-AuthenticationMethod
            if ($authMethod -eq 'B') { return }
            
            if ($authMethod.Method -eq 'Interactive') {
                # Select tenant and connect interactively
                $adminEmail = Select-TenantEnvironment
                if ($adminEmail -eq 'B') { return }
                
                if (-not (Connect-MicrosoftServices -AdminEmail $adminEmail)) {
                    throw "Failed to establish required connections"
                }
            }
            else {
                # Connect using saved authentication config
                if (-not (Connect-MicrosoftServices -OrganizationName $authMethod.Config.OrganizationName -Environment $authMethod.Config.Environment)) {
                    throw "Failed to establish required connections"
                }
            }

            # Get user with validation
            $user = Get-UserForOffboarding
            if (-not $user) {
                Write-Host "`nOffboarding process cancelled." -ForegroundColor Yellow
                Show-NavigationOptions -AllowMainMenu -AllowContinue
                $choice = Get-NavigationChoice -AllowMainMenu -AllowContinue
                if ($choice -eq 'M') { return }
                if ($choice -eq 'C') { continue }                
            }
            else {                
                if (Remove-UserAccess -UserPrincipalName $user.UserPrincipalName) {
                    Write-Host "`nOffboarding completed successfully!" -ForegroundColor Green
                }
                else {
                    Write-Host "`nOffboarding completed with errors. Please check the logs." -ForegroundColor Yellow
                }
            }
            
            Show-NavigationOptions -AllowMainMenu -AllowContinue
            $choice = Get-NavigationChoice -AllowMainMenu -AllowContinue
            if ($choice -eq 'M') { return }
            if ($choice -eq 'C') { continue }
        }
        catch {
            Write-Host "`nCritical error during offboarding process:" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
            
            Show-NavigationOptions -AllowMainMenu
            $choice = Get-NavigationChoice -AllowMainMenu
            if ($choice -eq 'M') { return }
            if ($choice -eq 'X') { exit }
        }
        finally {
            # Cleanup connections
            Invoke-Cleanup
        }
    } while ($true)
}

#Main onboarding function
function Get-UserForOnboarding {
    param (
        [int]$MaxAttempts = 5
    )
    
    $attempts = 0
    do {
        $attempts++
        
        # Define generic locations with placeholder data
        $locations = @{
            "1" = @{
                Name = "Primary Office"
                Phone = "+1 (555) 123-4567"
                Fax = "+1 (555) 123-4568"
                Address = "123 Corporate Drive"
                Country = "United States"
                UsageLocation = "USA"
                PostalCode = "12345"
            }
            "2" = @{
                Name = "Secondary Office"
                Phone = "+1 (555) 987-6543"
                Fax = "+1 (555) 987-6544"
                Address = "456 Business Boulevard"
                Country = "Canada"
                UsageLocation = "CA"
                PostalCode = "67890"
            }
        }

        # Location selection
        Write-Host "`nSelect Location:" -ForegroundColor Yellow
        foreach ($key in $locations.Keys | Sort-Object) {
            Write-Host "$key. $($locations[$key].Name)" -ForegroundColor White
        }
        Write-Host "M. Enter Location Manually" -ForegroundColor Yellow
        Write-Host "B. Go Back" -ForegroundColor Yellow

        $locationChoice = Read-Host "`nEnter location number"
        if ($locationChoice -eq 'B') { return $null }
        
        # Location selection logic
        if ($locationChoice -eq 'M') {
            $selectedLocation = @{
                Name = Read-Host "Enter location name"
                Phone = Read-Host "Enter phone number"
                Fax = Read-Host "Enter fax number"
                Address = Read-Host "Enter street address"
                Country = Read-Host "Enter country"
                UsageLocation = Read-Host "Enter usage location (2-letter country code)"
                PostalCode = Read-Host "Enter postal code"
            }
        }
        elseif ($locationChoice -in $locations.Keys) {
            $selectedLocation = $locations[$locationChoice]
        }
        else {
            Write-Host "Invalid location choice. Please try again." -ForegroundColor Red
            continue
        }

        # Collect user information
        Write-Host "`nEnter User Details:" -ForegroundColor Yellow
        
        # Create user hashtable with all required fields
        $userInfo = @{
            GivenName = Read-Host "Enter first name"
            Surname = Read-Host "Enter last name"
            JobTitle = Read-Host "Enter job title"
            Email = Read-Host "Enter email address"
            # Ensure BusinessPhones is an array
            Fax = @($selectedLocation.Fax)
            BusinessPhones = @($selectedLocation.Phone)
            OfficeLocation = $selectedLocation.Name
            StreetAddress = $selectedLocation.Address
            PostalCode = $selectedLocation.PostalCode
            Country = $selectedLocation.Country
            UsageLocation = $selectedLocation.UsageLocation
            # Add required fields for Azure
            AccountEnabled = $true
        }

        # Add derived fields
        $userInfo.DisplayName = "$($userInfo.GivenName) $($userInfo.Surname)"
        $userInfo.MailNickname = "$($userInfo.GivenName).$($userInfo.Surname)".ToLower() -replace '\s',''
        $userInfo.UserPrincipalName = $userInfo.Email

        # Validate email format
        if (-not ($userInfo.Email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')) {
            Write-Host "Invalid email format. Please try again." -ForegroundColor Red
            continue
        }

        # Validate all required fields have values
        $requiredFields = @(
            'GivenName', 'Surname', 'JobTitle', 'Email', 'BusinessPhones',
            'OfficeLocation', 'StreetAddress', 'PostalCode', 'Country', 'UsageLocation'
        )

        $missingFields = $requiredFields.Where({
            [string]::IsNullOrWhiteSpace($userInfo[$_])
        })

        if ($missingFields) {
            Write-Host "Missing required fields: $($missingFields -join ', ')" -ForegroundColor Red
            continue
        }

        # Debug output with privacy masking
        Write-Host "`nCollected User Information:" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor Cyan
        $userInfo.GetEnumerator() | Sort-Object Name | ForEach-Object {
            $displayValue = if ($_.Key -match 'Email|Phone') { 
                $_.Value -replace '(?<=.{3}).+(?=.{3})', '*****'
            } else { 
                $_.Value 
            }
            Write-Host "$($_.Key): $displayValue" -ForegroundColor White
        }
        Write-Host "----------------------------------------" -ForegroundColor Cyan

        # Confirm information
        $confirm = Read-Host "`nIs this information correct? (Y/N)"
        if ($confirm -eq 'Y') {
            return $userInfo
        }
    } while ($attempts -lt $MaxAttempts)

    Write-Host "Maximum attempts exceeded." -ForegroundColor Red
    return $null
}

function Start-OnboardingProcess {
    do {
        Clear-Host
        
        # Show welcome banner
        Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║    Azure User Onboarding Assistant     ║" -ForegroundColor Cyan
        Write-Host "║             Version 3.1                ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host

        try {
            # Select authentication method
            $authMethod = Select-AuthenticationMethod
            if ($authMethod -eq 'B') { return }
            
            if ($authMethod.Method -eq 'Interactive') {
                # Select tenant and connect interactively
                $adminEmail = Select-TenantEnvironment
                if ($adminEmail -eq 'B') { return }
                
                if (-not (Connect-MicrosoftServices -AdminEmail $adminEmail)) {
                    throw "Failed to establish required connections"
                }
            }
            else {
                # Connect using saved authentication config
                if (-not (Connect-MicrosoftServices -OrganizationName $authMethod.Config.OrganizationName -Environment $authMethod.Config.Environment)) {
                    throw "Failed to establish required connections"
                }
            }

            # Get user info with validation
            $userInfo = Get-UserForOnboarding
            if ($userInfo -eq 'MainMenu') {
                return
            }
            if ($userInfo -eq 'Back') {
                continue
            }
            if (-not $userInfo) {
                Write-Host "Onboarding process cancelled." -ForegroundColor Yellow
                Show-NavigationOptions -AllowMainMenu -AllowContinue
                $choice = Get-NavigationChoice -AllowMainMenu -AllowContinue
                if ($choice -eq 'M') { return }
                if ($choice -eq 'C') { continue }    
            }
        
            # Proceed with onboarding
            Write-Host "Proceeding with onboarding for: $($userInfo.Email)" -ForegroundColor Green
            
            $result = Add-UserAccess -UserInfo $userInfo
            
            Show-NavigationOptions -AllowMainMenu -AllowContinue
            $choice = Get-NavigationChoice -AllowMainMenu -AllowContinue
            if ($choice -eq 'M') { return }
            if ($choice -eq 'C') { continue }
        }
        catch {
            Write-Host "Critical error: $_" -ForegroundColor Red
            Show-NavigationOptions -AllowMainMenu
            $choice = Get-NavigationChoice -AllowMainMenu
            if ($choice -eq 'M') { return }
            if ($choice -eq 'X') { exit }
        }
        finally {
            # Cleanup connections
            Invoke-Cleanup
        }
    } while ($true)
}
#EndRegion

#Region Authentication Functions
function SetupAuthentication {
    # Initialize and start menu
    Initialize-AuthDirectories
    Show-AuthMenu
}
function Initialize-AuthDirectories {
    # Define root config path relative to script location
    $script:ConfigRootPath = Join-Path $PSScriptRoot "Config"
    
    if (-not (Test-Path $ConfigRootPath)) {
        New-Item -ItemType Directory -Path $ConfigRootPath -Force | Out-Null
        Write-Host "Created configuration directory: $ConfigRootPath" -ForegroundColor Yellow
    }
}
function Show-AuthMenu {
    while ($true) {
        Clear-Host
        Write-Host "====================================" -ForegroundColor Cyan
        Write-Host "    Azure Authentication Manager     " -ForegroundColor Cyan
        Write-Host "====================================" -ForegroundColor Cyan
        Write-Host
        Write-Host "1. Create New Configuration" -ForegroundColor Yellow
        Write-Host "2. List Available Configurations" -ForegroundColor Yellow
        Write-Host "3. Test Authentication" -ForegroundColor Yellow
        Write-Host "4. Return to Main Menu" -ForegroundColor Yellow
        Write-Host
        Write-Host "Select an option (1-4): " -NoNewline
        
        $choice = Read-Host
        
        switch ($choice) {
            "1" { Invoke-NewAuthConfigMenu }
            "2" { Invoke-ListAuthConfigs }
            "3" { Invoke-TestAuthMenu }
            "4" { return }
            default { 
                Write-Host "`nInvalid option. Press any key to continue..." -ForegroundColor Red
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
        }
    }
}
function Invoke-NewAuthConfigMenu {
    Clear-Host
    Write-Host "=== Create New Configuration ===" -ForegroundColor Cyan
    Write-Host

    $orgName = Read-Host "Enter Organization Name"
    $environments = @("prod", "dev", "test")
    
    Write-Host "`nSelect Environment:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $environments.Count; $i++) {
        Write-Host "$($i + 1). $($environments[$i])"
    }
    
    do {
        $envChoice = Read-Host "`nEnter choice (1-$($environments.Count))"
    } while ([int]$envChoice -lt 1 -or [int]$envChoice -gt $environments.Count)
    
    $environment = $environments[[int]$envChoice - 1]
    
    try {
        New-SecureAuthConfig -OrganizationName $orgName -Environment $environment
        Write-Host "`nPress any key to return to main menu..." -ForegroundColor Green
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    catch {
        Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Press any key to return to main menu..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}
function New-SecureAuthConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OrganizationName,

        [Parameter(Mandatory = $false)]
        [string]$Environment = "prod"
    )

    try {
        # Create config name from organization and environment
        $configName = "$($OrganizationName.ToLower())_$($Environment.ToLower()).encrypted"
        $fullConfigPath = Join-Path $ConfigRootPath $configName

        # Collect configuration values
        Write-Host "`nEntering configuration for: $OrganizationName ($Environment)" -ForegroundColor Cyan
        $config = @{
            OrganizationName = $OrganizationName
            Environment = $Environment
            ClientId = Read-Host "Enter Azure AD Application ID"
            TenantId = Read-Host "Enter Tenant ID"
            ClientSecret = Read-Host "Enter Client Secret" 
            CertThumbprint = Read-Host "Enter Certificate Thumbprint"
            TenantDomain = Read-Host "Enter Tenant Domain (e.g., contoso.onmicrosoft.com)"
        }

        # Convert and save encrypted
        ConvertTo-Json $config | ConvertTo-SecureString -AsPlainText -Force | 
        ConvertFrom-SecureString | Set-Content $fullConfigPath

        Write-Host "`nConfiguration saved successfully!" -ForegroundColor Green
        Write-Host "Location: $fullConfigPath" -ForegroundColor Green
        
        return $fullConfigPath
    }
    catch {
        Write-Host "Error creating secure configuration: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}
function Get-SecureAuthConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OrganizationName,

        [Parameter(Mandatory = $false)]
        [string]$Environment = "prod"
    )

    try {
        $configName = "$($OrganizationName.ToLower())_$($Environment.ToLower()).encrypted"
        $configPath = Join-Path $ConfigRootPath $configName

        if (-not (Test-Path $configPath)) {
            throw "Configuration file not found for $OrganizationName ($Environment) at: $configPath"
        }

        $secureConfig = Get-Content $configPath | ConvertTo-SecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureConfig)
        $config = ConvertFrom-Json ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))
        
        return $config
    }
    catch {
        Write-Host "Error reading secure configuration: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}
function Invoke-ListAuthConfigs {
    Clear-Host
    Write-Host "=== Available Configurations ===" -ForegroundColor Cyan
    Write-Host

    $configs = Get-AvailableAuthConfigs
    
    if ($configs.Count -eq 0) {
        Write-Host "No configurations found." -ForegroundColor Yellow
    }
    else {
        $configs | Format-Table @{
            Label = "Organization"
            Expression = {$_.Organization}
            Width = 20
        },
        @{
            Label = "Environment"
            Expression = {$_.Environment}
            Width = 15
        },
        @{
            Label = "Last Modified"
            Expression = {$_.LastModified.ToString("yyyy-MM-dd HH:mm:ss")}
            Width = 20
        }
    }
    
    Write-Host "`nPress any key to return to main menu..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
function Get-AvailableAuthConfigs {
    [CmdletBinding()]
    param()

    try {
        if (-not (Test-Path $ConfigRootPath)) {
            return @()
        }

        $configs = Get-ChildItem -Path $ConfigRootPath -Filter "*.encrypted" | 
                  ForEach-Object {
                      $nameParts = $_.BaseName -split '_'
                      [PSCustomObject]@{
                          Organization = $nameParts[0]
                          Environment = $nameParts[1]
                          FullPath = $_.FullName
                          LastModified = $_.LastWriteTime
                      }
                  }

        return $configs
    }
    catch {
        Write-Host "Error listing configurations: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}
function Invoke-TestAuthMenu {
    Clear-Host
    Write-Host "=== Test Authentication ===" -ForegroundColor Cyan
    Write-Host

    try {
        $configs = @(Get-AvailableAuthConfigs) # Force array
        
        if ($null -eq $configs -or $configs.Count -eq 0) {
            Write-Host "`nNo configurations found. Please create a configuration first." -ForegroundColor Red
            Write-Host "`nPress any key to return to main menu..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }

        Write-Host "Available Configurations:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $configs.Count; $i++) {
            Write-Host "$($i + 1). $($configs[$i].Organization) ($($configs[$i].Environment))"
        }
        
        Write-Host # Add a blank line
        $prompt = "Select configuration (1-$($configs.Count)), or 'q' to return to main menu"
        do {
            $choice = Read-Host $prompt
            if ($choice -eq 'q') {
                return
            }
            
            $validChoice = $false
            if ([int]::TryParse($choice, [ref]$null)) {
                $choiceNum = [int]$choice
                if ($choiceNum -ge 1 -and $choiceNum -le $configs.Count) {
                    $validChoice = $true
                    $selectedConfig = $configs[$choiceNum - 1]
                    
                    Write-Host "`nTesting authentication for $($selectedConfig.Organization) ($($selectedConfig.Environment))..."
                    Test-AzureAppAuthentication -OrganizationName $selectedConfig.Organization -Environment $selectedConfig.Environment
                    break
                }
            }
            if (-not $validChoice) {
                Write-Host "Invalid selection. Please try again." -ForegroundColor Red
            }
        } while ($true)
    }
    catch {
        Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        Write-Host "`nPress any key to return to main menu..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}
function Test-AzureAppAuthentication {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OrganizationName,

        [Parameter(Mandatory = $false)]
        [string]$Environment = "prod"
    )
    
    try {
        Write-Host "Testing Azure App Authentication for $OrganizationName ($Environment)..." -ForegroundColor Yellow
        
        # Get configuration
        $config = Get-SecureAuthConfig -OrganizationName $OrganizationName -Environment $Environment
        
        # Disconnect any existing sessions
        Write-Host "Disconnecting existing sessions..." -ForegroundColor Yellow
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        
        # Connect to Microsoft Graph
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        $secureSecret = ConvertTo-SecureString -String $config.ClientSecret -AsPlainText -Force
        $credentials = New-Object System.Management.Automation.PSCredential ($config.ClientId, $secureSecret)
        Connect-MgGraph -ClientSecretCredential $credentials -TenantId $config.TenantId -ErrorAction Stop
        
        # Test Graph connection
        $org = Get-MgOrganization
        Write-Host "Successfully connected to Graph: $($org.DisplayName)" -ForegroundColor Green
        
        # Connect to Exchange Online
        Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -AppId $config.ClientId `
                              -CertificateThumbprint $config.CertThumbprint `
                              -Organization $config.TenantDomain `
                              -ShowBanner:$false
        
        # Test Exchange connection
        Write-Host "Testing Exchange Online connection..." -ForegroundColor Yellow
        $mailboxes = Get-Mailbox -ResultSize 3
        Write-Host "Successfully retrieved mailboxes:" -ForegroundColor Green
        $mailboxes | Format-Table DisplayName, UserPrincipalName
        
        return $true
    }
    catch {
        Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        return $false
    }
    finally {
        Write-Host "`nCleaning up connections..." -ForegroundColor Yellow
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}
#EndRegion

#Region Module Management
# Function to validate and install required modules
function Install-RequiredModules {
    param (
        [switch]$Force
    )
    
    # Check for administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-ProgressStatus -Message "Administrator privileges required for module installation" -Status Warning -Color Yellow
        Write-Host "Please restart the script as Administrator to install modules" -ForegroundColor Yellow
        return $false
    }

    # Required modules with their minimum versions
    $requiredModules = @{
        'Microsoft.Graph' = '2.0.0'
        'ExchangeOnlineManagement' = '3.0.0'
    }

    $installNeeded = $false
    $modulesToInstall = @()

    Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║        Module Validation Check         ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host

    # Check each required module
    foreach ($moduleName in $requiredModules.Keys) {
        $minVersion = [version]$requiredModules[$moduleName]
        $module = Get-Module -ListAvailable -Name $moduleName | 
                 Sort-Object Version -Descending | 
                 Select-Object -First 1

        Write-ProgressStatus -Message "Checking $moduleName..." -Status Progress -Color Yellow
        
        if (-not $module) {
            Write-ProgressStatus -Message "$moduleName is not installed" -Status Warning -Color Yellow
            $installNeeded = $true
            $modulesToInstall += $moduleName
        }
        elseif ($Force -or $module.Version -lt $minVersion) {
            Write-ProgressStatus -Message "$moduleName version $($module.Version) needs update to $minVersion" -Status Warning -Color Yellow
            $installNeeded = $true
            $modulesToInstall += $moduleName
        }
        else {
            Write-ProgressStatus -Message "$moduleName version $($module.Version) is up to date" -Status Success -Color Green
        }
    }

    # Install or update modules if needed
    if ($installNeeded) {
        Write-Host "`nModules requiring installation/update:" -ForegroundColor Yellow
        $modulesToInstall | ForEach-Object { Write-Host "- $_" -ForegroundColor Yellow }

        $proceed = Read-Host "`nDo you want to proceed with module installation/update? (Y/N)"
        if ($proceed -eq 'Y') {
            foreach ($moduleName in $modulesToInstall) {
                try {
                    Write-ProgressStatus -Message "Installing $moduleName..." -Status Progress -Color Yellow
                    
                    # Remove existing module if updating
                    if (Get-Module -ListAvailable -Name $moduleName) {
                        Uninstall-Module -Name $moduleName -AllVersions -Force
                    }

                    # Install latest version
                    Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser
                    Write-ProgressStatus -Message "$moduleName installed successfully" -Status Success -Color Green
                }
                catch {
                    Write-ProgressStatus -Message "Failed to install $moduleName $_" -Status Failed -Color Red
                    return $false
                }
            }
        }
        else {
            Write-Host "Module installation cancelled" -ForegroundColor Yellow
            return $false
        }
    }

    Write-Host "`nAll required modules are ready" -ForegroundColor Green
    return $true
}
#EndRegion

#Region Main Menu Interface
# Function to display animated banner
function Show-AnimatedBanner {
    $colors = @('Red', 'Yellow', 'Green', 'Cyan', 'Blue', 'Magenta')
    $colorIndex = 0
    
    $banner = @"
    ╔═══════════════════════════════════════════╗
    ║        _    _____ _    _ ____  _____      ║
    ║       / \  |__  /| |  | |  _ \| ____|     ║
    ║      / _ \   / / | |  | | |_) |  _|       ║
    ║     / ___ \ / /_ | |__| |  _ <| |___      ║
    ║    /_/   \_/____| \____/|_| \_\_____|     ║
    ║                                           ║
    ║              Management Suite             ║
    ║                Version 4.1                ║
    ║                                           ║
    ╚═══════════════════════════════════════════╝
"@

    $banner.Split("`n") | ForEach-Object {
        Write-Host $_ -ForegroundColor $colors[$colorIndex]
        $colorIndex = ($colorIndex + 1) % $colors.Length
    }
}

# Function to show main menu
function Show-MainMenu {
    Write-Host "`n═══════ Main Menu ═══════" -ForegroundColor Cyan
    Write-Host "1: Start User Onboarding" -ForegroundColor Yellow
    Write-Host "2: Start User Offboarding" -ForegroundColor Yellow
    Write-Host "3: Manage Email & Calendar" -ForegroundColor Yellow
    Write-Host "4: Validate Environment" -ForegroundColor Yellow
    Write-Host "5: Configure Authentication" -ForegroundColor Yellow
    Write-Host "6: View Documentation" -ForegroundColor Yellow
    Write-Host "7: Exit" -ForegroundColor Yellow
    Write-Host "════════════════════════" -ForegroundColor Cyan
}

# Function to display documentation
function Show-Documentation {
    # Configuration
    $script:CONFIG = @{
        InitialMoney = 100
        MinimumBet = 1
        DealerStandThreshold = 17
        ShuffleThreshold = 40
        DelayBetweenActions = 0.5
        DelayAfterDeal = 0.3
        DelayAfterResult = 1.5
    }

    # Card display settings
    $script:CARD_SYMBOLS = @{
        Hearts = '♥'
        Diamonds = '♦'
        Clubs = '♣'
        Spades = '♠'
        Hidden = '??'
    }

    # Rest of the helper functions remain the same until Show-Hand
    function New-Deck {
        $suits = @($CARD_SYMBOLS.Hearts, $CARD_SYMBOLS.Diamonds, $CARD_SYMBOLS.Clubs, $CARD_SYMBOLS.Spades)
        $values = @('2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K', 'A')
        $deck = @()
        
        foreach ($suit in $suits) {
            foreach ($value in $values) {
                $deck += @{
                    'Suit' = $suit
                    'Value' = $value
                    'Display' = "$value$suit"
                    'Color' = if ($suit -in @($CARD_SYMBOLS.Hearts, $CARD_SYMBOLS.Diamonds)) { 'Red' } else { 'White' }
                }
            }
        }
        return $deck | Sort-Object {Get-Random}
    }

    function Get-CardValue {
        param (
            [Parameter(Mandatory=$true)]
            [hashtable]$Card,
            [int]$CurrentTotal = 0
        )
        
        switch ($Card.Value) {
            {'J', 'Q', 'K' -contains $_} { return 10 }
            'A' { return if ($CurrentTotal + 11 -gt 21) { 1 } else { 11 } }
            default { return [int]$Card.Value }
        }
    }

    function Get-HandTotal {
        param (
            [Parameter(Mandatory=$true)]
            [array]$Hand
        )
        
        $total = 0
        $aces = 0
        
        foreach ($card in $Hand) {
            if ($card.Value -eq 'A') {
                $aces++
                $total += 11
            } else {
                $total += Get-CardValue $card
            }
        }
        
        while ($total -gt 21 -and $aces -gt 0) {
            $total -= 10
            $aces--
        }
        
        return $total
    }

    function Get-HandDisplay {
        param (
            [Parameter(Mandatory=$true)]
            [array]$Hand,
            [bool]$HideSecondCard = $false
        )
        
        $displayParts = @()
        if ($HideSecondCard) {
            $displayParts += @(
                @{Text = $Hand[0].Display; Color = $Hand[0].Color},
                @{Text = $CARD_SYMBOLS.Hidden; Color = 'White'}
            )
        } else {
            $total = Get-HandTotal $Hand
            foreach ($card in $Hand) {
                $displayParts += @{Text = $card.Display; Color = $card.Color}
            }
            $displayParts += @{Text = "(Total: $total)"; Color = 'White'}
        }
        return $displayParts
    }

    function Show-GameState {
        param (
            [int]$Money,
            [int]$CurrentBet,
            [array]$PlayerHand,
            [array]$DealerHand,
            [bool]$HideDealerCard = $true,
            [string]$Message = "",
            [string]$MessageColor = 'White'
        )
        
        Clear-Host
        Write-Host "`n=== Blackjack ===" -ForegroundColor Cyan
        Write-Host "Current Balance: " -NoNewline
        Write-Host "`$$Money" -ForegroundColor Green
        Write-Host "Current Bet: " -NoNewline
        Write-Host "`$$CurrentBet" -ForegroundColor Yellow
        Write-Host "==================`n" -ForegroundColor Cyan

        # Display player's hand
        Write-Host "Your hand: " -NoNewline
        $playerDisplay = Get-HandDisplay $PlayerHand
        foreach ($part in $playerDisplay) {
            Write-Host $part.Text -NoNewline -ForegroundColor $part.Color
            Write-Host " " -NoNewline
        }
        Write-Host ""

        # Display dealer's hand
        Write-Host "Dealer's hand: " -NoNewline
        $dealerDisplay = Get-HandDisplay $DealerHand -HideSecondCard $HideDealerCard
        foreach ($part in $dealerDisplay) {
            Write-Host $part.Text -NoNewline -ForegroundColor $part.Color
            Write-Host " " -NoNewline
        }
        Write-Host ""

        if ($Message) {
            Write-Host "`n$Message" -ForegroundColor $MessageColor
        }
    }

    function Can-DoubleDown {
        param (
            [Parameter(Mandatory=$true)]
            [array]$Hand,
            [int]$CurrentBet,
            [int]$PlayerMoney
        )
        
        return ($Hand.Count -eq 2) -and ($PlayerMoney -ge $CurrentBet * 2)
    }

    function Start-BlackjackGame {
        $deck = New-Deck
        $deckIndex = 0
        $playerMoney = $CONFIG.InitialMoney
        $gameStats = @{
            Wins = 0
            Losses = 0
            Pushes = 0
        }
    
        Clear-Host
        Write-Host "`nWelcome to Blackjack!" -ForegroundColor Cyan
        Write-Host "Press Enter to start..." -ForegroundColor Cyan
        $null = Read-Host
        
        while ($playerMoney -ge $CONFIG.MinimumBet) {
            Clear-Host
            Write-Host "`n=== Blackjack ===" -ForegroundColor Cyan
            Write-Host "Current Balance: " -NoNewline
            Write-Host "`$$playerMoney" -ForegroundColor Green
            Write-Host "==================`n" -ForegroundColor Cyan
            
            $bet = Read-Host "Enter your bet (or 'quit' to exit)"
            
            if ($bet -eq 'quit') { break }
            if ($bet -notmatch '^\d+$' -or [int]$bet -gt $playerMoney -or [int]$bet -lt $CONFIG.MinimumBet) {
                Write-Host "`nInvalid bet. Please enter a number between $($CONFIG.MinimumBet) and $playerMoney" -ForegroundColor Red
                Write-Host "Press Enter to continue..." -ForegroundColor Cyan
                $null = Read-Host
                continue
            }
            
            $bet = [int]$bet
            $playerHand = @($deck[$deckIndex++], $deck[$deckIndex++])
            $dealerHand = @($deck[$deckIndex++], $deck[$deckIndex++])
            
            Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand
            
            # Player's turn
            $playerStand = $false
            while ((Get-HandTotal $playerHand) -lt 21 -and -not $playerStand) {
                $canDoubleDown = Can-DoubleDown -Hand $playerHand -CurrentBet $bet -PlayerMoney $playerMoney
                $options = if ($canDoubleDown) { "(H)it, (S)tand, or (D)ouble down?" } else { "(H)it or (S)tand?" }
                $action = Read-Host "`nDo you want to $options"
                
                switch ($action.ToUpper()) {
                    'H' {
                        $playerHand += $deck[$deckIndex++]
                        Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand
                    }
                    'D' {
                        if ($canDoubleDown) {
                            $bet *= 2
                            $playerHand += $deck[$deckIndex++]
                            Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand -Message "Doubled down! New bet: `$$bet" -MessageColor 'Yellow'
                            $playerStand = $true
                        } else {
                            Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand -Message "Double down not available!" -MessageColor 'Red'
                            Write-Host "`nPress Enter to continue..." -ForegroundColor Cyan
                            $null = Read-Host
                        }
                    }
                    'S' {
                        $playerStand = $true
                    }
                    default {
                        Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand -Message "Invalid action!" -MessageColor 'Red'
                        Write-Host "`nPress Enter to continue..." -ForegroundColor Cyan
                        $null = Read-Host
                        continue
                    }
                }
            }
            
            $playerTotal = Get-HandTotal $playerHand
            if ($playerTotal -gt 21) {
                Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand -HideDealerCard $false -Message "Bust! You lose `$$bet" -MessageColor 'Red'
                $playerMoney -= $bet
                $gameStats.Losses++
                Write-Host "`nPress Enter to continue..." -ForegroundColor Cyan
                $null = Read-Host
            } else {
                # Dealer's turn
                Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand -HideDealerCard $false
                
                while ((Get-HandTotal $dealerHand) -lt $CONFIG.DealerStandThreshold) {
                    Write-Host "`nDealer takes a card..." -ForegroundColor Cyan
                    Write-Host "Press Enter to continue..." -ForegroundColor Cyan
                    $null = Read-Host
                    $dealerHand += $deck[$deckIndex++]
                    Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand -HideDealerCard $false
                }
                
                $dealerTotal = Get-HandTotal $dealerHand
                $resultMessage = if ($dealerTotal -gt 21) {
                    $playerMoney += $bet
                    $gameStats.Wins++
                    @{Text = "Dealer busts! You win `$$bet!"; Color = 'Green'}
                } elseif ($dealerTotal -gt $playerTotal) {
                    $playerMoney -= $bet
                    $gameStats.Losses++
                    @{Text = "Dealer wins! You lose `$$bet"; Color = 'Red'}
                } elseif ($playerTotal -gt $dealerTotal) {
                    $playerMoney += $bet
                    $gameStats.Wins++
                    @{Text = "You win `$$bet!"; Color = 'Green'}
                } else {
                    $gameStats.Pushes++
                    @{Text = "Push! Bet returned"; Color = 'Yellow'}
                }
                
                Show-GameState -Money $playerMoney -CurrentBet $bet -PlayerHand $playerHand -DealerHand $dealerHand -HideDealerCard $false -Message $resultMessage.Text -MessageColor $resultMessage.Color
                Write-Host "`nPress Enter to continue..." -ForegroundColor Cyan
                $null = Read-Host
            }
            
            if ($deckIndex -gt $CONFIG.ShuffleThreshold) {
                Write-Host "`nShuffling deck..." -ForegroundColor Cyan
                Write-Host "Press Enter to continue..." -ForegroundColor Cyan
                $null = Read-Host
                $deck = New-Deck
                $deckIndex = 0
            }
        }
        
        # Game over summary
        Clear-Host
        Write-Host "`n=== Game Over ===" -ForegroundColor Cyan
        Write-Host "Final Balance: " -NoNewline
        Write-Host "`$$playerMoney" -ForegroundColor $(if ($playerMoney -gt $CONFIG.InitialMoney) { 'Green' } else { 'Red' })
        Write-Host "`nGame Statistics:"
        Write-Host "Wins: " -NoNewline
        Write-Host $gameStats.Wins -ForegroundColor Green
        Write-Host "Losses: " -NoNewline
        Write-Host $gameStats.Losses -ForegroundColor Red
        Write-Host "Pushes: " -NoNewline
        Write-Host $gameStats.Pushes -ForegroundColor Yellow
        Write-Host "`nThanks for playing!" -ForegroundColor Cyan
    }

    # Start the game
    Start-BlackjackGame
}

function Initialize-ScriptDirectories {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ScriptRoot
    )

    # Directories to create
    $requiredDirectories = @(
        "Config",
        "Reports",
        "Reports\Onboarding",
        "Reports\Offboarding"
    )

    foreach ($dir in $requiredDirectories) {
        $fullPath = Join-Path -Path $ScriptRoot -ChildPath $dir
        
        if (-not (Test-Path -Path $fullPath)) {
            try {
                New-Item -ItemType Directory -Path $fullPath -Force | Out-Null
            }
            catch {
                Write-Host "Failed to create directory: $dir" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Directory exists: $dir" -ForegroundColor Cyan
        }
    }
}
#EndRegion

#Region Script Entry Point
# Main script execution
Clear-Host
$scriptRoot = $PSScriptRoot
Initialize-ScriptDirectories -ScriptRoot $scriptRoot
Clear-Host
Show-AnimatedBanner

do {
    Show-MainMenu
    $choice = Read-Host "`nEnter your choice (1-6)"
    
    switch ($choice) {
        '1' {
            Clear-Host
            Start-OnboardingProcess
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        '2' {
            Clear-Host
            Start-OffboardingProcess
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        '3' {
            Clear-Host
            Start-EmailCalendarManagement
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        '4' {
            Clear-Host
            Install-RequiredModules
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        '5' {
            Clear-Host
            SetupAuthentication
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
        '6' {
            Clear-Host
            Show-Documentation
        }
        '7' {
            Write-Host "Cleaning up and exiting..." -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-Host "Goodbye!" -ForegroundColor Cyan
            return
        }
        default {
            Write-Host "Invalid choice. Please select 1-6." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Clear-Host
            Show-AnimatedBanner
        }
    }
    Clear-Host
    Show-AnimatedBanner
} while ($true)

#EndRegion

# Export functions for module usage
Export-ModuleMember -Function @(
    'Start-OffboardingProcess',
    'Start-AzureUserOnboarding',
    'Start-EmailCalendarManagement',
    'Install-RequiredModules',
    'Test-Environment',
    'Show-Documentation'
)
