<##########################################################################################################################################
.SYNOPSIS
    This script provides simplified tools to search and clean the Recoverable Items folder for Exchange users.

.DESCRIPTION
    This PowerShell script offers two main functions:
    1. Search for recoverable items on user accounts - Displays detailed information about the size and contents of hidden recoverable items folders.
    2. Purge recoverable items folder - Uses Exchange Management Shell commands to permanently delete content from the Recoverable Items folder.

.NOTES
    File Name      : Clean-RecoverableItems.ps1
    Author         : Con Moore
    Prerequisite   : PowerShell V.7.2.18 or later, Exchange Online Management Module, Compliance Search Module (implicitly used via Connect-IPPSSession)
    Version        : 4.1
    Created Date   : 03/10/2024
    Last Modified  : 2025-04-15 (Refactored for UI/UX and minimal fixes)

    Version History:
    - 1.0: Initial script - Trash🗑️....work in progress....2024/03/10
    - 2.0: 2024/11/10
        -Refactored Powershell version checking to include installing Powershell 7.
        -Refactored folder size checking to prevent hangs.
        -Updated positional parameters
    - 3.0: 2025/03/17
        -Added menu-based interface with switch statement
        -Added credential collection at startup
        -Added Show-Logo and Show-MainMenu functions
    - 4.0: 2025/04/14
        -Simplified menu to focus on two core functions
        -Added Search-RecoverableItems function to quickly view folder sizes
        -Refactored Script to be function driven instead of procedural for easier management
    - 4.1: 2025-04-15
        - Improved CLI aesthetics (logo, menu, output formatting).
        - Standardized prompts and user feedback.
        - Added pauses for better readability.
        - Minor fixes (redundant module import, unused variable check).

.EXAMPLE
    .\Clean-RecoverableItems.ps1

    This command executes the script which will prompt for Admin credentials and present a simplified menu with two options:
    1. Search for recoverable items on user account - Shows the size and item count of recoverable items folders.
    2. Purge recoverable items folder - Permanently removes all items from the recoverable items folder.

    *Note: To execute this command, it is necessary to grant the admin email account permissions for conducting searches and making modifications to accounts within Exchange.
     This requires Exchange and Compliance Center admin permissions. Ensure PIMs is setup to check out the required roles.

############################################################################################################################################>

### Parameters
# Although parameters are defined, the script primarily uses an interactive prompt model via global variables.
param (
    [Parameter(Mandatory = $false)]
    [securestring]$CredentialsParam, # Renamed to avoid conflict with $Global:Credentials if used
    [Parameter(Mandatory = $false)]
    [string]$TechEmailParam,         # Renamed
    [Parameter(Mandatory = $false)]
    [string]$EmployeeEmailParam      # Renamed
)

### Set Execution Policy (Consider if this is always necessary/desirable)
# Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force # Force prevents prompts

### Global Variables
$Global:Credentials = $null
$Global:TechEmail = $null
$Global:EmployeeEmail = $null
$Global:SearchName = $null # Initialize SearchName
$Global:RecoverableItemFolderId = $null
$Global:DeletionsFolderId = $null
$Global:DiscoveryHoldsFolderId = $null
$Global:SearchDiscoveryHoldsFolderId = $null
$Global:SubstrateHoldsFolderId = $null

# --- UI Helper ---
function Write-Separator {
    param([string]$Character = '-', [int]$Length = 80)
    Write-Host ($Character * $Length) -ForegroundColor DarkGray
}

#Region-Functions
# Dependency Check Functions
function Check-PowerShell7 {
    Write-Separator
    Write-Host "Checking PowerShell Version..." -ForegroundColor Cyan
    # Check if running in PowerShell 7
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        $scriptPath = $MyInvocation.MyCommand.Path
        $ps7Path = Join-Path $env:ProgramFiles "PowerShell\7\pwsh.exe" # More robust path construction
        if (-not (Test-Path $ps7Path)) {
            # PowerShell 7 is not installed, download and install it
            Write-Host "PowerShell 7 is required but not installed." -ForegroundColor Yellow
            Write-Host "Attempting to download and install PowerShell 7.2.0..." -ForegroundColor Yellow

            $installerPath = Join-Path $env:TEMP "PowerShell-7.2.0-win-x64.msi"
            $downloadUri = "https://github.com/PowerShell/PowerShell/releases/download/v7.2.0/PowerShell-7.2.0-win-x64.msi"

            try {
                Invoke-WebRequest -Uri $downloadUri -OutFile $installerPath -ErrorAction Stop
                Write-Host "Installer downloaded to '$installerPath'." -ForegroundColor Green

                Write-Host "Starting PowerShell 7 installation (requires administrator privileges)..." -ForegroundColor Yellow
                # Note: Start-Process with msiexec might require elevation depending on UAC settings.
                Start-Process msiexec.exe -ArgumentList "/i `"$installerPath`" /quiet /norestart" -Wait -Verb RunAs

                # Basic check if installation likely succeeded
                if (Test-Path $ps7Path) {
                    Write-Host "PowerShell 7 installation seems complete." -ForegroundColor Green
                } else {
                    Write-Warning "PowerShell 7 installation might have failed or requires a system restart."
                    Write-Warning "Please install PowerShell 7 manually and re-run the script."
                    Read-Host "Press Enter to exit"
                    Exit 1
                }
                Remove-Item $installerPath -ErrorAction SilentlyContinue
            }
            catch {
                Write-Error "Failed to download or install PowerShell 7. Error: $($_.Exception.Message)"
                Write-Error "Please install PowerShell 7 manually (version 7.2 or later recommended) and re-run the script."
                Read-Host "Press Enter to exit"
                Exit 1
            }
        }

        # Restart the script in PowerShell 7
        Write-Host "Restarting script in PowerShell 7..." -ForegroundColor Cyan
        try {
            Start-Process -FilePath $ps7Path -ArgumentList "-File `"$scriptPath`"" -Wait
            exit # Exit the current (older PS) session
        }
        catch {
            Write-Error "Failed to restart script in PowerShell 7. Error: $($_.Exception.Message)"
            Read-Host "Press Enter to exit"
            Exit 1
        }
    } else {
        Write-Host "Script is running in PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor) (version 7+)." -ForegroundColor Green
    }
    Write-Separator
    # Pause briefly
    # Start-Sleep -Seconds 1
}

function Install-RequiredModules {
    Write-Separator
    Write-Host "Checking Required PowerShell Modules..." -ForegroundColor Cyan
    # Define the module names
    $exchangeModuleName = "ExchangeOnlineManagement"
    # $complianceModuleName = "ComplianceSearch" # This module is implicitly used via Connect-IPPSSession, not installed separately.

    # Check for Exchange Online Management module
    if (-not (Get-Module -ListAvailable -Name $exchangeModuleName)) {
        Write-Host "Module '$exchangeModuleName' is not installed. Attempting to install..." -ForegroundColor Yellow

        # Attempt to install the module
        try {
            # Install the module from the PowerShell Gallery
            Install-Module -Name $exchangeModuleName -Force -AllowClobber -Scope CurrentUser -Confirm:$false -ErrorAction Stop
            Write-Host "Module '$exchangeModuleName' installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install module '$exchangeModuleName'. Error: $($_.Exception.Message)"
            Write-Error "Please try installing it manually: Install-Module $exchangeModuleName -Scope CurrentUser"
            Read-Host "Press Enter to exit"
            Exit 1
        }
    }
    else {
        Write-Host "Module '$exchangeModuleName' is already installed." -ForegroundColor Green
    }

    # Import Exchange Online Management module
    try {
        # Check if already imported in this session to avoid redundant messages
        if (-not (Get-Module -Name $exchangeModuleName)) {
             Import-Module $exchangeModuleName -ErrorAction Stop
             Write-Host "Module '$exchangeModuleName' imported successfully." -ForegroundColor Green
        } else {
             Write-Host "Module '$exchangeModuleName' is already imported in this session." -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to import module '$exchangeModuleName'. Error: $($_.Exception.Message)"
        Read-Host "Press Enter to exit"
        Exit 1
    }
    Write-Separator
    # Pause briefly
    # Start-Sleep -Seconds 1
}


#Menu Functions
function Show-Logo {
    # Using a slightly more compact logo for better fit in standard terminals
    Write-Host @"
                                 ███╗   ███╗ ██████╗  ██████╗ ███╗   ██╗██╗███╗   ██╗██╗████████╗███████╗    ████████╗ ██████╗  ██████╗ ██╗     ██╗███╗   ██╗ ██████╗ 
                                 ████╗ ████║██╔═══██╗██╔═══██╗████╗  ██║██║████╗  ██║██║╚══██╔══╝██╔════╝    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██║████╗  ██║██╔════╝ 
                                 ██╔████╔██║██║   ██║██║   ██║██╔██╗ ██║██║██╔██╗ ██║██║   ██║   █████╗         ██║   ██║   ██║██║   ██║██║     ██║██╔██╗ ██║██║  ███╗
                                 ██║╚██╔╝██║██║   ██║██║   ██║██║╚██╗██║██║██║╚██╗██║██║   ██║   ██╔══╝         ██║   ██║   ██║██║   ██║██║     ██║██║╚██╗██║██║   ██║
                                 ██║ ╚═╝ ██║╚██████╔╝╚██████╔╝██║ ╚████║██║██║ ╚████║██║   ██║   ███████╗       ██║   ╚██████╔╝╚██████╔╝███████╗██║██║ ╚████║╚██████╔╝
                                 ╚═╝     ╚═╝ ╚═════╝  ╚═════╝ ╚═╝  ╚═══╝╚═╝╚═╝  ╚═══╝╚═╝   ╚═╝   ╚══════╝       ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚═╝╚═╝  ╚═══╝ ╚═════╝ 
"@ -ForegroundColor Magenta -NoNewline
    Write-Host "   TOOL" -ForegroundColor Magenta # Append TOOL to the logo line

    Write-Host ("=" * 75) -ForegroundColor Cyan
    Write-Host (" " * 18 + "EXCHANGE RECOVERABLE ITEMS CLEANUP") -ForeccgroundColor Yellow
    Write-Host ("=" * 75) -ForegroundColor Cyan
    Write-Host (" " * 32 + "Version 4.1") -ForegroundColor Green
    Write-Host # Blank line for spacing
}

function Show-MainMenu {
    Write-Separator '=' 75
    Write-Host " MAIN MENU" -ForegroundColor Green
    Write-Separator '=' 75
    Write-Host "[1] Search Recoverable Items" -ForegroundColor White
    Write-Host "    (View sizes and item counts for a user)"
    Write-Host "[2] Purge Recoverable Items" -ForegroundColor Yellow
    Write-Host "    (Permanently delete items for a user)"
    Write-Host "[Q] Quit" -ForegroundColor Red
    Write-Separator '-' 75

    # Loop until a valid choice is entered
    while ($true) {
        $choice = Read-Host "Enter your choice [1, 2, Q]"
        if ($choice -in '1', '2', 'Q', 'q') {
            return $choice.ToUpper() # Return uppercase Q for consistency
        } else {
            Write-Warning "Invalid choice. Please enter 1, 2, or Q."
        }
    }
}

# --- Input Functions ---
function Get-AdminCredentials {
    Clear-Host
    Show-Logo
    Write-Host "Administrator Credentials Required" -ForegroundColor Cyan
    Write-Separator
    Write-Host "Before proceeding, please enter the credentials for an account with" -ForegroundColor White
    Write-Host "Exchange Online Administrator and Compliance Administrator roles." -ForegroundColor White
    Write-Host "(Example format: YourDomain\AdminUsername or admin.user@yourdomain.com)" -ForegroundColor Gray
    Write-Separator
    $Global:Credentials = Get-Credential
    Clear-Host
    Show-Logo
    Write-Host "Credentials captured for user: $($Global:Credentials.UserName)" -ForegroundColor Green
    Write-Separator
    Read-Host "Press Enter to continue"
}

function Set-AdminEmail {
    Write-Separator
    Write-Host "Enter your Admin Email Address" -ForegroundColor Cyan
    Write-Host "(Must have permissions for Exchange/Compliance tasks)" -ForegroundColor Gray
    $Global:TechEmail = Read-Host "Admin Email"
    Write-Host "Admin email set to: $Global:TechEmail" -ForegroundColor Green
    Write-Separator
}

function Set-EmployeeEmail {
     Write-Separator
    Write-Host "Enter the Target Employee Email Address" -ForegroundColor Cyan
    Write-Host "(The user whose recoverable items you want to manage)" -ForegroundColor Gray
    $Global:EmployeeEmail = Read-Host "Employee Email"
    Write-Host "Target employee email set to: $Global:EmployeeEmail" -ForegroundColor Green
    Write-Separator
}

# --- Core Exchange Functions ---
function Prepare-UserMailbox {
    if ([string]::IsNullOrEmpty($Global:TechEmail) -or [string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Write-Warning "Admin and Employee email addresses must be set first."
        return $false # Indicate failure
    }

    try {
        Write-Host "Connecting to Exchange Online (Admin: $Global:TechEmail)..." -ForegroundColor Cyan
        # Ensure connection uses the provided admin UPN, credentials might be cached or prompted if needed by Connect-ExchangeOnline
        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $Global:TechEmail -ErrorAction Stop
        Write-Separator
        Write-Host "Preparing mailbox for '$($Global:EmployeeEmail)'..." -ForegroundColor Yellow

        Write-Host "  - Disabling Client Access Services (EWS, ActiveSync, MAPI, OWA, IMAP, POP)..." -ForegroundColor Gray
        Set-CASMailbox $Global:EmployeeEmail -EwsEnabled $false -ActiveSyncEnabled $false -MAPIEnabled $false -OWAEnabled $false -ImapEnabled $false -PopEnabled $false -ErrorAction SilentlyContinue # Keep SilentlyContinue as original

        Write-Host "  - Setting Deleted Item Retention to 0 days..." -ForegroundColor Gray
        Set-Mailbox $Global:EmployeeEmail -RetainDeletedItemsFor 0 -ErrorAction SilentlyContinue

        Write-Host "  - Disabling Single Item Recovery..." -ForegroundColor Gray
        Set-Mailbox $Global:EmployeeEmail -SingleItemRecoveryEnabled $false -ErrorAction SilentlyContinue

        Write-Host "  - Disabling Managed Folder Assistant processing..." -ForegroundColor Gray
        Set-Mailbox $Global:EmployeeEmail -ElcProcessingDisabled $true -ErrorAction SilentlyContinue

        Write-Host "  - Disabling Litigation Hold..." -ForegroundColor Gray
        Set-Mailbox $Global:EmployeeEmail -LitigationHoldEnabled $false -ErrorAction SilentlyContinue

        Write-Host "User mailbox prepared successfully." -ForegroundColor Green
        Write-Separator
        return $true # Indicate success
    }
    catch {
        Write-Error "Failed to prepare mailbox '$($Global:EmployeeEmail)'. Error: $($_.Exception.Message)"
        Write-Separator
        return $false # Indicate failure
    }
}

function Find-FolderIDs {
    if ([string]::IsNullOrEmpty($Global:TechEmail) -or [string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Write-Warning "Admin and Employee email addresses must be set first."
        return $false
    }

    try {
        # Ensure connected to Exchange Online for Get-MailboxFolderStatistics
        # Reconnect if necessary, or rely on existing connection from Prepare-UserMailbox
        if (-not (Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' })) {
             Write-Host "Reconnecting to Exchange Online for folder statistics..." -ForegroundColor Cyan
             Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $Global:TechEmail -ErrorAction Stop
        }

        Write-Host "Retrieving folder statistics for '$($Global:EmployeeEmail)'..." -ForegroundColor Cyan
        Write-Host "(This may take a moment for mailboxes with many folders)" -ForegroundColor Gray

        # Clear previous IDs
        $Global:RecoverableItemFolderId = $null
        $Global:DeletionsFolderId = $null
        $Global:DiscoveryHoldsFolderId = $null
        $Global:SearchDiscoveryHoldsFolderId = $null
        $Global:SubstrateHoldsFolderId = $null

        # Pull all Exchange Folders for User
        $folderStatistics = Get-MailboxFolderStatistics $Global:EmployeeEmail -ErrorAction Stop
        Write-Host "Processing $($folderStatistics.Count) folders..." -ForegroundColor Gray

        # --- Original Folder ID Calculation Logic ---
        # This complex logic generates a 'folderid:' query string based on byte manipulation.
        # Keeping it as per minimal change request, though simpler methods might exist.
        $folderQueries = @()
        foreach ($folderStatistic in $folderStatistics) {
            $folderId = $folderStatistic.FolderId;
            $folderPath = $folderStatistic.FolderPath;
            $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
            $nibbler = $encoding.GetBytes("0123456789ABCDEF");
            $folderIdBytes = [Convert]::FromBase64String($folderId);
            $indexIdBytes = New-Object byte[] 48;
            $indexIdIdx = 0;
            # This specific byte manipulation likely targets a part of the ID relevant for querying
            $folderIdBytes | Select-Object -Skip 23 -First 24 | ForEach-Object {
                $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; # High nibble
                $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF] # Low nibble
            }
            $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";
            $folderStat = [PSCustomObject]@{
                FolderPath = $folderPath
                FolderQuery = $folderQuery
            }
            $folderQueries += $folderStat
        }
        # --- End Original Folder ID Calculation Logic ---

        Write-Host "Searching for specific Recoverable Items folder IDs..." -ForegroundColor Cyan

        # Define patterns (using single quotes for consistency)
        $patternBase = 'folderid:([A-Za-z0-9+/=]+)' # Base64-like pattern

        # Find Recoverable Items Folder ID
        $RecoverableItemsString = $folderQueries | Where-Object { $_.FolderPath -eq '/Recoverable Items' } | Select-Object -ExpandProperty FolderQuery -First 1
        $match = [regex]::Match($RecoverableItemsString, $patternBase)
        if ($match.Success) { $Global:RecoverableItemFolderId = $match.Value; Write-Host "  [✓] Recoverable Items:" -ForegroundColor Green -NoNewline; Write-Host " $Global:RecoverableItemFolderId" -ForegroundColor DarkGray }
        else { Write-Warning "Could not find Recoverable Items folder ID." }

        # Find Deletions Folder ID
        $DeletionsString = $folderQueries | Where-Object { $_.FolderPath -eq '/Recoverable Items/Deletions' } | Select-Object -ExpandProperty FolderQuery -First 1
        $match = [regex]::Match($DeletionsString, $patternBase)
        if ($match.Success) { $Global:DeletionsFolderId = $match.Value; Write-Host "  [✓] Deletions:" -ForegroundColor Green -NoNewline; Write-Host "         $Global:DeletionsFolderId" -ForegroundColor DarkGray }
        else { Write-Warning "Could not find Deletions folder ID." }

        # Find DiscoveryHolds Folder ID
        $DiscoveryHoldsString = $folderQueries | Where-Object { $_.FolderPath -eq '/Recoverable Items/DiscoveryHolds' } | Select-Object -ExpandProperty FolderQuery -First 1
        $match = [regex]::Match($DiscoveryHoldsString, $patternBase)
        if ($match.Success) { $Global:DiscoveryHoldsFolderId = $match.Value; Write-Host "  [✓] DiscoveryHolds:" -ForegroundColor Green -NoNewline; Write-Host "    $Global:DiscoveryHoldsFolderId" -ForegroundColor DarkGray }
        else { Write-Warning "Could not find DiscoveryHolds folder ID." }

        # Find SearchDiscoveryHoldsFolder Folder ID (Note: Path might vary)
        $SearchDiscoveryHoldsFolderString = $folderQueries | Where-Object { $_.FolderPath -like '/Recoverable Items/DiscoveryHolds/SearchDiscoveryHoldsFolder*' } | Select-Object -ExpandProperty FolderQuery -First 1
        $match = [regex]::Match($SearchDiscoveryHoldsFolderString, $patternBase)
        if ($match.Success) { $Global:SearchDiscoveryHoldsFolderId = $match.Value; Write-Host "  [✓] SearchDiscoveryHolds:" -ForegroundColor Green -NoNewline; Write-Host "$Global:SearchDiscoveryHoldsFolderId" -ForegroundColor DarkGray }
        else { Write-Warning "Could not find SearchDiscoveryHoldsFolder ID." } # This might be expected if not used

        # Find SubstrateHolds
        $SubstrateHoldsString = $folderQueries | Where-Object { $_.FolderPath -eq '/Recoverable Items/SubstrateHolds' } | Select-Object -ExpandProperty FolderQuery -First 1
        $match = [regex]::Match($SubstrateHoldsString, $patternBase)
        if ($match.Success) { $Global:SubstrateHoldsFolderId = $match.Value; Write-Host "  [✓] SubstrateHolds:" -ForegroundColor Green -NoNewline; Write-Host "    $Global:SubstrateHoldsFolderId" -ForegroundColor DarkGray }
        else { Write-Warning "Could not find SubstrateHolds folder ID." } # This might be expected if not used

        Write-Separator
        # Check if essential IDs were found
        if ($Global:RecoverableItemFolderId -and $Global:DeletionsFolderId) {
             Write-Host "Essential folder IDs found successfully." -ForegroundColor Green
             return $true
        } else {
             Write-Error "Failed to find one or more essential folder IDs (Recoverable Items, Deletions). Cannot proceed with purge."
             return $false
        }
    }
    catch {
        Write-Error "Failed to find folder IDs for '$($Global:EmployeeEmail)'. Error: $($_.Exception.Message)"
        Write-Separator
        return $false
    }
}

function Run-ComplianceSearch {
    if ([string]::IsNullOrEmpty($Global:TechEmail) -or [string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Write-Warning "Admin and Employee email addresses must be set first."
        return $false
    }
    # Check if required folder IDs are present
    if (-not ($Global:RecoverableItemFolderId -and $Global:DeletionsFolderId)) { # Only check essential ones found previously
        Write-Warning "Essential folder IDs (Recoverable Items, Deletions) are missing. Cannot run compliance search."
        return $false
    }

    $Global:SearchName = "$($Global:EmployeeEmail.Split('@')[0])-RecoverableItemsPurge-$(Get-Date -Format 'yyyyMMddHHmmss')" # More unique name
    Write-Host "Preparing Compliance Search..." -ForegroundColor Cyan
    Write-Host "Search Name: $Global:SearchName" -ForegroundColor Gray

    # Construct the query using only the found folder IDs
    $queryParts = @($Global:RecoverableItemFolderId, $Global:DeletionsFolderId, $Global:DiscoveryHoldsFolderId, $Global:SearchDiscoveryHoldsFolderId, $Global:SubstrateHoldsFolderId) | Where-Object { -not [string]::IsNullOrEmpty($_) }
    $contentQuery = $queryParts -join ' OR '

    if ([string]::IsNullOrEmpty($contentQuery)) {
        Write-Error "Could not construct a valid content query from found folder IDs."
        return $false
    }
    Write-Host "Content Query: $contentQuery" -ForegroundColor DarkGray

    try {
        Write-Host "Connecting to Security & Compliance Center PowerShell..." -ForegroundColor Cyan
        Connect-IPPSSession -ShowBanner:$false -UserPrincipalName $Global:TechEmail -ErrorAction Stop

        # Remove existing search with the same name, if any (e.g., from a failed previous run)
        Write-Host "Checking for existing compliance search '$Global:SearchName'..." -ForegroundColor Gray
        $existingSearch = Get-ComplianceSearch -Identity $Global:SearchName -ErrorAction SilentlyContinue
        if ($existingSearch) {
            Write-Host "Removing existing search..." -ForegroundColor Yellow
            Remove-ComplianceSearch -Identity $Global:SearchName -Confirm:$false -ErrorAction Stop
        }

        Write-Host "Creating new compliance search..." -ForegroundColor Cyan
        New-ComplianceSearch -Name $Global:SearchName -ExchangeLocation $Global:EmployeeEmail -ContentMatchQuery $contentQuery -ErrorAction Stop

        Write-Host "Starting compliance search..." -ForegroundColor Cyan
        Start-ComplianceSearch -Identity $Global:SearchName -ErrorAction Stop

        Write-Host "Waiting for search completion (this can take several minutes)..." -ForegroundColor Cyan
        $startTime = Get-Date
        $timeoutSeconds = 600 # 10 minutes timeout for the search itself
        $checkIntervalSeconds = 10

        Do {
            Start-Sleep -Seconds $checkIntervalSeconds
            $ComplianceSearchStatus = Get-ComplianceSearch -Identity $Global:SearchName -ErrorAction Stop
            $elapsedTime = (Get-Date) - $startTime
            Write-Host "  - Status: $($ComplianceSearchStatus.Status) (Elapsed: $($elapsedTime.ToString('hh\:mm\:ss')))" -ForegroundColor Gray

            if ($elapsedTime.TotalSeconds -gt $timeoutSeconds) {
                Write-Error "Compliance search timed out after $timeoutSeconds seconds."
                # Consider attempting to remove the timed-out search
                # Remove-ComplianceSearch -Identity $Global:SearchName -Confirm:$false -ErrorAction SilentlyContinue
                return $false
            }

        } Until ($ComplianceSearchStatus.Status -in 'Completed', 'Failed')

        Write-Separator

        if ($ComplianceSearchStatus.Status -eq 'Failed') {
             Write-Error "Compliance search failed. Status Details: $($ComplianceSearchStatus.StatusDetails)"
             # Consider removing the failed search
             # Remove-ComplianceSearch -Identity $Global:SearchName -Confirm:$false -ErrorAction SilentlyContinue
             return $false
        }

        # Status is 'Completed'
        Write-Host "Compliance Search Completed Successfully!" -ForegroundColor Green
        $SearchDetails = Get-ComplianceSearch -Identity $Global:SearchName
        $sizeInGB = [math]::Round($SearchDetails.Size / 1GB, 2)
        $itemCount = $SearchDetails.Items

        Write-Host "Results:" -ForegroundColor Cyan
        Write-Host "  - Items Found: $itemCount"
        Write-Host "  - Total Size: $sizeInGB GB"

        # Basic check if items were found
        if ($itemCount -eq 0) {
            Write-Host "No items found matching the query. No purge action needed for this search." -ForegroundColor Yellow
            # Clean up the search as it's not needed for purge
            Write-Host "Removing completed (empty) compliance search '$Global:SearchName'..." -ForegroundColor Gray
            Remove-ComplianceSearch -Identity $Global:SearchName -Confirm:$false -ErrorAction SilentlyContinue
            $Global:SearchName = $null # Clear search name as no purge is happening
            return $true # Return true, as the process didn't fail, just found nothing
        }

        return $true # Indicate search success and items found
    }
    catch {
        Write-Error "An error occurred during the compliance search process. Error: $($_.Exception.Message)"
        # Attempt cleanup if search name exists
        if ($Global:SearchName -and (Get-ComplianceSearch -Identity $Global:SearchName -ErrorAction SilentlyContinue)) {
            Write-Warning "Attempting to remove potentially incomplete compliance search '$Global:SearchName'..."
            Remove-ComplianceSearch -Identity $Global:SearchName -Confirm:$false -ErrorAction SilentlyContinue
        }
        $Global:SearchName = $null # Clear search name on failure
        Write-Separator
        return $false
    }
}

function Run-PurgeOperation {
    if ([string]::IsNullOrEmpty($Global:SearchName)) {
        Write-Warning "No active compliance search name found. Please run a successful compliance search first."
        return $false
    }

    $purgeActionIdentity = "${Global:SearchName}_Purge"
    $purgeSuccess = $false # Flag to track overall success

    try {
        # Loop for purging, especially if size remains large
        $maxPurgeAttempts = 5 # Limit attempts to prevent infinite loops
        $currentAttempt = 1
        $sizeThresholdGB = 20 # Original threshold

        do {
            Write-Separator '=' 75
            Write-Host "Starting Purge Attempt $currentAttempt of $maxPurgeAttempts for Search '$Global:SearchName'" -ForegroundColor Yellow
            Write-Separator '=' 75

            # Check for existing purge action for this specific search name (might exist from previous loop iteration)
            $existingAction = Get-ComplianceSearchAction -Identity $purgeActionIdentity -ErrorAction SilentlyContinue
            if ($existingAction) {
                Write-Host "Removing previous purge action '$purgeActionIdentity' before starting new one..." -ForegroundColor Yellow
                Remove-ComplianceSearchAction -Identity $purgeActionIdentity -Confirm:$false -ErrorAction Stop
                Start-Sleep -Seconds 5 # Give Azure time to process removal
            }

            Write-Host "Creating new purge action (HardDelete)..." -ForegroundColor Cyan
            New-ComplianceSearchAction -SearchName $Global:SearchName -Purge -PurgeType HardDelete -Confirm:$false -ErrorAction Stop

            Write-Host "Waiting for purge action completion..." -ForegroundColor Cyan
            $startTime = Get-Date
            $timeoutSeconds = 1800 # 30 minutes timeout for the purge action
            $checkIntervalSeconds = 15

            Do {
                Start-Sleep -Seconds $checkIntervalSeconds
                $ActionStatusResult = Get-ComplianceSearchAction -Identity $purgeActionIdentity -ErrorAction Stop
                $ComplianceSearchActionStatus = $ActionStatusResult.Status
                $elapsedTime = (Get-Date) - $startTime
                Write-Host "  - Status: $ComplianceSearchActionStatus (Elapsed: $($elapsedTime.ToString('hh\:mm\:ss')))" -ForegroundColor Gray

                if ($elapsedTime.TotalSeconds -gt $timeoutSeconds) {
                    Write-Error "Purge action '$purgeActionIdentity' timed out after $timeoutSeconds seconds."
                    # State is unknown, might need manual check. Don't assume failure/success.
                    # Consider *not* removing the action automatically here.
                    return $false # Exit function due to timeout
                }

            } Until ($ComplianceSearchActionStatus -in 'Completed', 'Failed')

            Write-Separator '-' 75

            if ($ComplianceSearchActionStatus -eq 'Failed') {
                Write-Error "Purge action '$purgeActionIdentity' failed."
                $PurgeErrors = ($ActionStatusResult | Select-Object -ExpandProperty Errors) -join '; '
                Write-Error "Errors: $PurgeErrors"
                # Don't automatically remove the failed action, might need investigation.
                return $false # Exit function due to failure
            }

            # Purge action completed
            Write-Host "Purge Action Completed Successfully!" -ForegroundColor Green
            $PurgeDetails = Get-ComplianceSearchAction -Identity $purgeActionIdentity | Select-Object SearchName, Name, Action, Results, Errors
            Write-Host "Purge Details:" -ForegroundColor Cyan
            $PurgeDetails | Format-List | Out-String | Write-Host

            # --- Check Size Post-Purge ---
            Write-Separator
            Write-Host "Running Managed Folder Assistant (may take time to reflect changes)..." -ForegroundColor Cyan
            Start-ManagedFolderAssistant -Identity $Global:EmployeeEmail
            Write-Host "Waiting briefly for MFA to potentially process..." -ForegroundColor Gray
            Start-Sleep -Seconds 30 # Give MFA some time

            Write-Host "Checking remaining Recoverable Items size..." -ForegroundColor Cyan
            $folderStats = Get-MailboxFolderStatistics -Identity $Global:EmployeeEmail -FolderScope RecoverableItems -ErrorAction SilentlyContinue
            if (-not $folderStats) {
                 Write-Warning "Could not retrieve folder statistics after purge. Unable to verify size."
                 # Assume success for this attempt, but don't loop again.
                 $sizeInGB = 0
            } else {
                $totalSizeBytes = ($folderStats | Measure-Object -Property FolderSizeRaw -Sum).Sum
                $sizeInGB = [math]::Round($totalSizeBytes / 1GB, 2)
                Write-Host "Current Estimated Size: $sizeInGB GB" -ForegroundColor ($sizeInGB -ge $sizeThresholdGB ? 'Yellow' : 'Green')
            }

            # --- Loop Condition ---
            if ($sizeInGB -ge $sizeThresholdGB -and $currentAttempt -lt $maxPurgeAttempts) {
                Write-Host "Size ($sizeInGB GB) is still >= $sizeThresholdGB GB. Preparing for next purge attempt." -ForegroundColor Yellow
                $currentAttempt++
                # The loop will continue, removing the completed action and starting a new one.
            } else {
                # Size is below threshold OR max attempts reached
                if ($sizeInGB -ge $sizeThresholdGB) {
                     Write-Warning "Maximum purge attempts ($maxPurgeAttempts) reached, but size ($sizeInGB GB) is still >= $sizeThresholdGB GB."
                     Write-Warning "Manual investigation might be required."
                } else {
                     Write-Host "Folder size is now below $sizeThresholdGB GB." -ForegroundColor Green
                }
                $purgeSuccess = $true # Mark overall purge as successful
                break # Exit the Do-While loop
            }

        } While ($true) # Loop controlled by break or return

        return $purgeSuccess

    }
    catch {
        Write-Error "An error occurred during the purge operation. Error: $($_.Exception.Message)"
        # State is uncertain, avoid automatic cleanup of action/search unless sure.
        Write-Separator
        return $false
    }
    finally {
         # --- Cleanup ---
         # Clean up the final purge action and the compliance search if the purge was marked successful
         if ($purgeSuccess) {
             Write-Separator
             Write-Host "Cleaning up purge resources..." -ForegroundColor Cyan
             if ($purgeActionIdentity -and (Get-ComplianceSearchAction -Identity $purgeActionIdentity -ErrorAction SilentlyContinue)) {
                 Write-Host "Removing final purge action '$purgeActionIdentity'..." -ForegroundColor Gray
                 Remove-ComplianceSearchAction -Identity $purgeActionIdentity -Confirm:$false -ErrorAction SilentlyContinue
             }
             if ($Global:SearchName -and (Get-ComplianceSearch -Identity $Global:SearchName -ErrorAction SilentlyContinue)) {
                 Write-Host "Removing compliance search '$Global:SearchName'..." -ForegroundColor Gray
                 Remove-ComplianceSearch -Identity $Global:SearchName -Confirm:$false -ErrorAction SilentlyContinue
             }
             $Global:SearchName = $null # Clear search name after successful cleanup
         } else {
             Write-Warning "Purge operation did not complete successfully or timed out."
             Write-Warning "Compliance search '$($Global:SearchName)' and action '$($purgeActionIdentity)' may still exist."
             Write-Warning "Manual cleanup might be required in the Security & Compliance Center."
         }
    }
}

function Restore-MailboxSettings {
    if ([string]::IsNullOrEmpty($Global:TechEmail) -or [string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Write-Warning "Admin and Employee email addresses must be set first."
        return $false
    }

    try {
        Write-Separator '=' 75
        Write-Host "Restoring Mailbox Settings for '$($Global:EmployeeEmail)'..." -ForegroundColor Cyan
        Write-Separator '=' 75

        # Ensure connected to Exchange Online
        if (-not (Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' })) {
             Write-Host "Reconnecting to Exchange Online..." -ForegroundColor Cyan
             Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $Global:TechEmail -ErrorAction Stop
        }

        Write-Host "  - Enabling Client Access Services (EWS, ActiveSync, MAPI, OWA, IMAP, POP)..." -ForegroundColor Gray
        Set-CASMailbox $Global:EmployeeEmail -EwsEnabled $true -ActiveSyncEnabled $true -MAPIEnabled $true -OWAEnabled $true -ImapEnabled $true -PopEnabled $true -ErrorAction SilentlyContinue

        Write-Host "  - Setting Deleted Item Retention back to 30 days..." -ForegroundColor Gray
        Set-Mailbox $Global:EmployeeEmail -RetainDeletedItemsFor 30 -ErrorAction SilentlyContinue

        Write-Host "  - Enabling Single Item Recovery..." -ForegroundColor Gray
        Set-Mailbox $Global:EmployeeEmail -SingleItemRecoveryEnabled $true -ErrorAction SilentlyContinue

        Write-Host "  - Enabling Managed Folder Assistant processing..." -ForegroundColor Gray
        Set-Mailbox $Global:EmployeeEmail -ElcProcessingDisabled $false -ErrorAction SilentlyContinue

        # Note: Litigation Hold is not re-enabled automatically. This should be a conscious decision.
        Write-Host "Mailbox settings restored (Litigation Hold remains OFF)." -ForegroundColor Green
        Write-Separator
        return $true
    }
    catch {
        Write-Error "Failed to restore settings for mailbox '$($Global:EmployeeEmail)'. Error: $($_.Exception.Message)"
        Write-Warning "Manual verification of mailbox settings is recommended."
        Write-Separator
        return $false
    }
}

# --- Main Workflow Functions ---
function Search-RecoverableItems {
    Clear-Host
    Show-Logo
    Write-Host "Search Recoverable Items" -ForegroundColor Cyan
    Write-Separator '=' 75

    # Dependency Checks
    Check-PowerShell7
    Install-RequiredModules # Ensures EXO module is present and imported

    # Get Inputs if needed
    if ([string]::IsNullOrEmpty($Global:TechEmail)) { Set-AdminEmail }
    if ([string]::IsNullOrEmpty($Global:EmployeeEmail)) { Set-EmployeeEmail }

    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    try {
        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $Global:TechEmail -ErrorAction Stop
    } catch {
        Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
        Read-Host "Press Enter to return to the main menu"
        return
    }

    Write-Host "Searching for recoverable items in '$($Global:EmployeeEmail)' mailbox..." -ForegroundColor Cyan
    Write-Host "(This might take a moment)" -ForegroundColor Gray
    Write-Separator

    try {
        $folderStats = Get-MailboxFolderStatistics -Identity $Global:EmployeeEmail -FolderScope RecoverableItems -ErrorAction Stop

        if (-not $folderStats) {
            Write-Warning "No Recoverable Items folders found for '$Global:EmployeeEmail'."
        } else {
            Write-Host "Recoverable Items Folders for '$Global:EmployeeEmail':" -ForegroundColor Green
            Write-Separator '-' 75

            $totalSizeGB = 0
            $totalItems = 0
            $outputData = @()

            foreach ($folder in $folderStats) {
                # Use raw size for accurate calculation
                $sizeGB = [math]::Round($folder.FolderSizeRaw / 1GB, 3) # Use 3 decimal places for GB
                $totalSizeGB += $sizeGB
                $totalItems += $folder.ItemsInFolder

                $outputData += [PSCustomObject]@{
                    Path        = $folder.FolderPath
                    Items       = $folder.ItemsInFolder
                    Size        = $folder.FolderSize # Display friendly size
                    SizeGB      = $sizeGB
                    # Created     = $folder.CreationTime # Optional
                }
            }

            # Display as a formatted table
            $outputData | Format-Table -AutoSize -Wrap `
                @{ Label = 'Folder Path'; Expression = { $_.Path }; Alignment = 'Left' }, `
                @{ Label = 'Item Count'; Expression = { $_.Items }; Alignment = 'Right'; Width = 12 }, `
                @{ Label = 'Size (Friendly)'; Expression = { $_.Size }; Alignment = 'Right'; Width = 18 }, `
                @{ Label = 'Size (GB)'; Expression = { $_.SizeGB.ToString('N3') }; Alignment = 'Right'; Width = 12 }


            Write-Separator '-' 75
            Write-Host "Summary:" -ForegroundColor Yellow
            Write-Host "  - Total Folders Found: $($folderStats.Count)"
            Write-Host "  - Total Items Found:   $totalItems"
            Write-Host "  - Total Size (GB):     $($totalSizeGB.ToString('N3')) GB"
            Write-Separator '-' 75
        }
        Write-Host "Search completed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to retrieve folder statistics. Error: $($_.Exception.Message)"
    }

    Write-Separator
    Read-Host "Press Enter to return to the main menu"
}

function Purge-RecoverableItems {
    Clear-Host
    Show-Logo
    Write-Host "Purge Recoverable Items" -ForegroundColor Yellow
    Write-Separator '=' 75

    # Dependency Checks
    Check-PowerShell7
    Install-RequiredModules

    # Get Inputs if needed
    if ([string]::IsNullOrEmpty($Global:TechEmail)) { Set-AdminEmail }
    if ([string]::IsNullOrEmpty($Global:EmployeeEmail)) { Set-EmployeeEmail }

    # --- Confirmation ---
    Write-Host "`n⚠️ WARNING ⚠️" -ForegroundColor Red
    Write-Host "This action will attempt to PERMANENTLY DELETE items from the" -ForegroundColor Red
    Write-Host "Recoverable Items folders for '$Global:EmployeeEmail'." -ForegroundColor Red
    Write-Host "This includes Deletions, Purges, DiscoveryHolds, etc." -ForegroundColor Red
    Write-Host "THIS ACTION CANNOT BE UNDONE." -ForegroundColor Red
    Write-Separator '-' 75
    $confirm = Read-Host "Are you absolutely sure you want to proceed? (Type 'YES' to confirm)"

    if ($confirm -ne "YES") {
        Write-Warning "Operation cancelled by user."
        Read-Host "Press Enter to return to the main menu"
        return
    }
    # --- End Confirmation ---

    Clear-Host
    Show-Logo
    Write-Host "Starting Purge Operation for '$Global:EmployeeEmail'..." -ForegroundColor Green
    Write-Separator '=' 75

    # --- Execute Steps ---
    $stepSuccess = $true # Track overall success

    # 1. Prepare Mailbox
    if ($stepSuccess) {
        Write-Host "[Step 1/5] Preparing Mailbox..." -ForegroundColor Cyan
        $stepSuccess = Prepare-UserMailbox
    }

    # 2. Find Folder IDs
    if ($stepSuccess) {
        Write-Host "[Step 2/5] Finding Folder IDs..." -ForegroundColor Cyan
        $stepSuccess = Find-FolderIDs
    }

    # 3. Run Compliance Search
    if ($stepSuccess) {
        Write-Host "[Step 3/5] Running Compliance Search..." -ForegroundColor Cyan
        $stepSuccess = Run-ComplianceSearch
        # Check if search found items, if not, skip purge but don't mark as failure
        if ($stepSuccess -and [string]::IsNullOrEmpty($Global:SearchName)) {
             Write-Host "Compliance search found no items. Skipping Purge and Restore steps." -ForegroundColor Green
             # Reset stepSuccess to true because finding nothing isn't a failure of the *purge* goal
             $stepSuccess = $true
             # Skip steps 4 and 5
             goto EndPurgeProcess
        }
    }

    # 4. Run Purge Operation
    if ($stepSuccess) {
        Write-Host "[Step 4/5] Running Purge Operation..." -ForegroundColor Cyan
        $stepSuccess = Run-PurgeOperation
    }

    # 5. Restore Mailbox Settings (Always attempt restore unless preparation failed)
    # We attempt restore even if purge failed, as the mailbox was still prepared.
    Write-Host "[Step 5/5] Restoring Mailbox Settings..." -ForegroundColor Cyan
    # Don't overwrite $stepSuccess here, just report if restore fails
    if (-not (Restore-MailboxSettings)) {
         Write-Warning "Failed to fully restore mailbox settings. Manual check recommended."
    }

    :EndPurgeProcess # Label for goto

    # --- Final Summary ---
    Write-Separator '=' 75
    if ($stepSuccess) {
        Write-Host "Purge Operation Workflow Completed Successfully!" -ForegroundColor Green
    } else {
        Write-Error "Purge Operation Workflow Encountered Errors."
        Write-Error "Please review the logs above for details. Manual intervention may be required."
    }
    Write-Separator '=' 75
    Read-Host "Press Enter to return to the main menu"
}
#EndRegion-Functions



### Main Script Execution ###
# Initial Credential Check
if ($null -eq $Global:Credentials) {
    Get-AdminCredentials
}

# Main Menu Loop
while ($true) {
    Clear-Host
    Show-Logo
    $choice = Show-MainMenu

    # Process menu choice
    switch ($choice) {
        '1' { Search-RecoverableItems }
        '2' { Purge-RecoverableItems }
        'Q' {
            Write-Host "Exiting script. Goodbye!" -ForegroundColor Cyan
            # Consider disconnecting sessions if they are active
            # Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
            exit
        }
        # Default case handled within Show-MainMenu validation loop
    }
}
