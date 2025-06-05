
<# Region-Synopsis
.SYNOPSIS
    This script provides simplified tools to search and clean the Recoverable Items folder for Exchange users.

.DESCRIPTION
    This PowerShell script offers two main functions:
    1. Search for recoverable items on user accounts - Displays detailed information about the size and contents of hidden recoverable items folders
    2. Purge recoverable items folder - Uses Exchange Management Shell commands to permanently delete content from the Recoverable Items folder

.NOTES
    File Name      : Clean-RecoverableItems.ps1
    Author         : Con Moore
    Prerequisite   : PowerShell V.7.2.18 or later, Exchange Online Managment Module, Compliance Managment Module
    Version        : 5.0
    Created Date   : 03/10/2024
    Last Modified  : 05/24/2025

    Version History:
    - 1.0: Initial script - Trash🗑️....work in progress....2024/03/10
    - 2.0: 2024/11/10
        -Refactored Powershell version checking to inlude insatlling Powershell 7.
        -Refactored folder size checking to prevent hangs.
        -Updated positional parameters
    - 3.0: 2025/03/17
        -Added menu-based interface with switch statement
        -Added credential collection at startup
        -Added Show-Logo and Show-MainMenu functions
    - 4.0: 2025/04/14
        -Simplified menu to focus on two core functions
        -Added Search-RecoverableItems function to quickly view folder sizes
        -Refactored Script to be function driven instead of procedural for easier managment
    - 5.0: 2025/05/24
        -Refactored script to use new syntax that is complinant with the new Purview Compliance Center (I.E FolderID:"FolderId" to FolderID="FolderId")

.EXAMPLE
    .\Clean-RecoverableItems.ps1

    This command executes the script which will prompt for Admin credentials and present a simplified menu with two options:
    1. Search for recoverable items on user account - Shows the size and item count of recoverable items folders
    2. Purge recoverable items folder - Permanently removes all items from the recoverable items folder

    *Note: To execute this command, it is necessary to grant the admin email account permissions for conducting searches and making modifications to accounts within Exchange.
     This requires Exchange and Compliance Center admin permissions. Ensure PIMs is setup with the required roles.

#EndRegion-Synopsis#>


#Region- Var and Script Config

param (
    [Parameter(Mandatory = $false)]
    [securestring]$Credentials,
    [Parameter(Mandatory = $false)]
    [string]$TechEmail,
    [Parameter(Mandatory = $false)]
    [string]$EmployeeEmail
)

### --- Set Execution Policy ---
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser

### Global Variables
$Global:Credentials = $null
$Global:TechEmail = $null
$Global:EmployeeEmail = $null
#EndRegion-Var Intiluzation


#Region-Functions

# --- Dependency Functions ---
function Check-PowerShell7 {
    # Check if running in PowerShell 7
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        $scriptPath = $MyInvocation.MyCommand.Path
        $ps7Path = "C:\Program Files\PowerShell\7\pwsh.exe"
        if (-not (Test-Path $ps7Path)) {
            # PowerShell 7 is not installed, download and install it
            Write-Host "PowerShell 7 is not installed. Downloading and installing..." -ForegroundColor Yellow

            $installerPath = "$env:TEMP\PowerShell-7.2.0-win-x64.msi"
            Invoke-WebRequest -Uri "https://github.com/PowerShell/PowerShell/releases/download/v7.2.0/PowerShell-7.2.0-win-x64.msi" -OutFile $installerPath

            Write-Host "Installing PowerShell 7..." -ForegroundColor Yellow
            Start-Process msiexec.exe -ArgumentList "/i `"$installerPath`" /quiet /norestart" -Wait

            Remove-Item $installerPath
        }

        # Restart the script in PowerShell 7
        Write-Host "Restarting script in PowerShell 7..." -ForegroundColor Cyan
        Start-Process -FilePath $ps7Path -ArgumentList "-File `"$scriptPath`"" -Wait
        exit
    } else {
        Write-Host "Script is running in PowerShell 7." -ForegroundColor Green
        Write-Host "Press any key to continue..."
        #[void][System.Console]::ReadKey($true)
    }
}
function Install-RequiredModules {
    # Define the module names
    $exchangeModuleName = "ExchangeOnlineManagement"
    $complianceModuleName = "ComplianceSearch"

    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force

    # Check for Exchange Online Management module
    if (-not (Get-Module -ListAvailable -Name $exchangeModuleName)) {
        Write-Host "Module $exchangeModuleName is not installed. Attempting to install..."

        # Attempt to install the module
        try {
            # Install the module from the PowerShell Gallery
            Install-Module -Name $exchangeModuleName -Force -AllowClobber -Scope CurrentUser -Confirm:$false
            Write-Host "Module $exchangeModuleName installed successfully."
        }
        catch {
            Write-Error "Failed to install module $exchangeModuleName. Error: $_"
        }
    }
    else {
        Write-Host "Module $exchangeModuleName is already installed."
    }

    # Import Exchange Online Management module
    try {
        Import-Module $exchangeModuleName -ErrorAction Stop
        Write-Host "Module $exchangeModuleName imported successfully."
    }
    catch {
        Write-Error "Failed to import module $exchangeModuleName. Error: $_"
    }

    Import-Module ExchangeOnlineManagement
    Write-Host "Press any key to continue..."
    #[void][System.Console]::ReadKey($true)
}


#--- Menu Functions ---
function Show-Logo {
    Write-Host "
                                 ███╗   ███╗ ██████╗  ██████╗ ███╗   ██╗██╗███╗   ██╗██╗████████╗███████╗    ████████╗ ██████╗  ██████╗ ██╗     ██╗███╗   ██╗ ██████╗
                                 ████╗ ████║██╔═══██╗██╔═══██╗████╗  ██║██║████╗  ██║██║╚══██╔══╝██╔════╝    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██║████╗  ██║██╔════╝
                                 ██╔████╔██║██║   ██║██║   ██║██╔██╗ ██║██║██╔██╗ ██║██║   ██║   █████╗         ██║   ██║   ██║██║   ██║██║     ██║██╔██╗ ██║██║  ███╗
                                 ██║╚██╔╝██║██║   ██║██║   ██║██║╚██╗██║██║██║╚██╗██║██║   ██║   ██╔══╝         ██║   ██║   ██║██║   ██║██║     ██║██║╚██╗██║██║   ██║
                                 ██║ ╚═╝ ██║╚██████╔╝╚██████╔╝██║ ╚████║██║██║ ╚████║██║   ██║   ███████╗       ██║   ╚██████╔╝╚██████╔╝███████╗██║██║ ╚████║╚██████╔╝
                                 ╚═╝     ╚═╝ ╚═════╝  ╚═════╝ ╚═╝  ╚═══╝╚═╝╚═╝  ╚═══╝╚═╝   ╚═╝   ╚══════╝       ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚═╝╚═╝  ╚═══╝ ╚═════╝ " -ForegroundColor Magenta

    Write-Host "                                                                =======================================================" -ForegroundColor Cyan
    Write-Host "                                                                       EXCHANGE RECOVERABLE ITEMS CLEANUP TOOL         "-ForegroundColor Yellow
    Write-Host "                                                                =======================================================" -ForegroundColor Cyan
    Write-Host "                                                                                                            Version 4.0" -ForegroundColor Green
}
function Show-MainMenu {
    Write-Host "Main Menu:" -ForegroundColor Green
    Write-Host "1. Search for recoverable items on user account" -ForegroundColor White
    Write-Host "2. Purge recoverable items folder" -ForegroundColor Yellow
    Write-Host "Q. Quit" -ForegroundColor Red

    $choice = Read-Host "Enter your choice"
    return $choice
}


# --- Exchange Functions ---
function Set-AdminEmail {
    Clear-Host
    Write-Host "Enter your work email address that has the necessary permissions to conduct the compliance search below:" -Foreground Yellow
    $Global:TechEmail = Read-Host
    Write-Host "Admin email set to: $Global:TechEmail" -ForegroundColor Green
}
function Set-EmployeeEmail {
    write-host "Enter the email address of the UPT Employee you are wanting to query or edit below:" -ForegroundColor Yellow
    $Global:EmployeeEmail = Read-Host
    Write-Host "Employee email set to: $Global:EmployeeEmail" -ForegroundColor Green
    Write-Host ""
}
function Prepare-UserMailbox {
    if ([string]::IsNullOrEmpty($Global:TechEmail) -or [string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Write-Host "Please set both Admin and Employee email addresses first." -ForegroundColor Red
        Write-Host ""
        #[void][System.Console]::ReadKey($true)
        return
    }

    Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $Global:TechEmail
    write-host "Step [1]: Prepare '$Global:EmployeeEmail' for processing......" -ForegroundColor Green
    Write-Host "————————————————————————————————————————————————————————————————————————————————————————————————"
    Write-Host ""
    Write-Host "Disabling WS, Active Sync, MAPI, OWA, IMAP, and POP for $Global:EmployeeEmail........" -ForegroundColor Yellow
    Set-CASMailbox $Global:EmployeeEmail -EwsEnabled $false -ActiveSyncEnabled $false -MAPIEnabled $false -OWAEnabled $false -ImapEnabled $false -PopEnabled $false -ErrorAction SilentlyContinue
    Write-Host ""

    Write-Host "Increasing the retention window to 30-days for $Global:EmployeeEmail......" -ForegroundColor Yellow
    Set-Mailbox $Global:EmployeeEmail -RetainDeletedItemsFor 0 -ErrorAction SilentlyContinue
    Write-Host ""

    Write-Host "Disabling single-item-recovery for $Global:EmployeeEmail......" -ForegroundColor Yellow
    Set-Mailbox $Global:EmployeeEmail -SingleItemRecoveryEnabled $false -ErrorAction SilentlyContinue
    Write-Host ""

    Write-Host "Disabling Manage Folder Assistant for $Global:EmployeeEmail......" -ForegroundColor Yellow
    Set-Mailbox $Global:EmployeeEmail -ElcProcessingDisabled $true -ErrorAction SilentlyContinue
    Write-Host ""

    Write-Host "Removing all holds on $Global:EmployeeEmail......" -ForegroundColor Yellow
    Set-Mailbox $Global:EmployeeEmail -LitigationHoldEnabled $false -ErrorAction SilentlyContinue
    Write-Host ""

    Write-Host "User mailbox prepared successfull!!!" -ForegroundColor Green
    Write-Host "`n`n`n`n`n"
}
function Find-FolderIDs {
    if ([string]::IsNullOrEmpty($Global:TechEmail) -or [string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Write-Host "Please set both Admin and Employee email addresses first." -ForegroundColor Red
        Write-Host ""
        #[void][System.Console]::ReadKey($true)
        return
    }

    # Pull all Exchange Folders for User
    if (!$ExoSession) {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $Global:TechEmail -CommandName Get-MailboxFolderStatistics
    }

    $folderQueries = @()

    # Check if mailbox exists first
    try {
        $mailboxCheck = Get-Mailbox -Identity $Global:EmployeeEmail -ErrorAction Stop
        Write-Host "Mailbox found: $($mailboxCheck.DisplayName)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: Unable to find mailbox '$Global:EmployeeEmail'" -ForegroundColor Red
        Write-Host "Please verify the email address is correct and the mailbox exists." -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    try {
        $folderStatistics = Get-MailboxFolderStatistics $Global:EmployeeEmail -FolderScope RecoverableItems -ErrorAction Stop
    }
    catch {
        Write-Host "Error: Unable to retrieve mailbox folder statistics for '$Global:EmployeeEmail'" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if ($null -eq $folderStatistics -or $folderStatistics.Count -eq 0) {
        Write-Host "No recoverable items folders found for '$Global:EmployeeEmail'" -ForegroundColor Yellow
        return
    }
    foreach ($folderStatistic in $folderStatistics) {
        $folderId = $folderStatistic.FolderId;
        $folderPath = $folderStatistic.FolderPath;
        $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
        $nibbler = $encoding.GetBytes("0123456789ABCDEF");
        $folderIdBytes = [Convert]::FromBase64String($folderId);
        $indexIdBytes = New-Object byte[] 48;
        $indexIdIdx = 0;
        $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
        $folderQuery = "folderid=$($encoding.GetString($indexIdBytes))";
        $folderStat = New-Object PSObject
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
        $folderQueries += $folderStat
    }

    write-host "Step [2]: Find all folder IDs for '$Global:EmployeeEmail' and their respective IDs." -ForegroundColor Green
    Write-Host " ————————————————————————————————————————————————————————————————————————————————————————————————"
    Write-Host ""
    # Assign various folder ID to VAR using Regular Expression
    Write-Host "Looking for Folder ID's, please be patient.........." -foreground Yellow
    Write-Host ""

    # Find Recoverable Items Folder ID
    $RecoverableItemsString = $folderQueries | Select-String "/Recoverable Items"
    if ($RecoverableItemsString) {
        $pattern = 'folderid=([A-Za-z0-9+/=]+)'
        $match = [regex]::Match($RecoverableItemsString, $pattern)

        if ($match.Success) {
            $Global:RecoverableItemFolderId = $match.value
            Write-host "Found Recoverable Items Folder - $Global:RecoverableItemFolderId" -ForegroundColor Cyan
        }
    } else {
        Write-Host "Warning: Recoverable Items folder not found" -ForegroundColor Yellow
    }

    # Find Deletions Folder ID
    $DeletionsString = $folderQueries | Select-String "/Deletions"
    if ($DeletionsString) {
        $pattern = "folderid=([A-Za-z0-9\+/=]+)"
        $match = [regex]::Match($DeletionsString, $pattern)

        if ($match.Success) {
            $Global:DeletionsFolderId = $match.value
            Write-host "Found Deletions Folder         - $Global:DeletionsFolderId" -ForegroundColor Cyan
        }
    } else {
        Write-Host "Warning: Deletions folder not found" -ForegroundColor Yellow
    }

    # Find DiscoveryHolds Folder ID
    $DiscoveryHoldsString = $folderQueries | Select-String "/DiscoveryHolds"
    if ($DiscoveryHoldsString) {
        $pattern = "folderid=([A-Za-z0-9+/=]+)"
        $match = [regex]::Match($DiscoveryHoldsString, $pattern)

        if ($match.Success) {
            $Global:DiscoveryHoldsFolderId = $match.value
            Write-host "Found DiscoveryHolds Folder    - $Global:DiscoveryHoldsFolderId" -ForegroundColor Cyan
        }
    } else {
        Write-Host "Warning: DiscoveryHolds folder not found" -ForegroundColor Yellow
    }

    # Find SearchDiscoveryHoldsFolder Folder ID
    $SearchDiscoveryHoldsFolderString = $folderQueries | Select-String "/DiscoveryHolds/SearchDiscoveryHoldsFolder"
    if ($SearchDiscoveryHoldsFolderString) {
        $pattern = "folderid=([A-Za-z0-9+/=]+)"
        $match = [regex]::Match($SearchDiscoveryHoldsFolderString, $pattern)

        if ($match.Success) {
            $Global:SearchDiscoveryHoldsFolderId = $match.value
            Write-host "Found Search Discovery Holds Folder    - $Global:SearchDiscoveryHoldsFolderId" -ForegroundColor Cyan
        }
    } else {
        Write-Host "Warning: SearchDiscoveryHoldsFolder not found" -ForegroundColor Yellow
    }

    # Find SubstrateHolds
    $SubstrateHoldsString = $folderQueries | Select-String "/SubstrateHolds"
    if ($SubstrateHoldsString) {
        $pattern = "folderid=([A-Za-z0-9+/=]+)"
        $match = [regex]::Match($SubstrateHoldsString, $pattern)

        if ($match.Success) {
            $Global:SubstrateHoldsFolderId = $match.value
            Write-host "Found SubstrateHolds Folder    - $Global:SubstrateHoldsFolderId" -ForegroundColor Cyan
        }
    } else {
        Write-Host "Warning: SubstrateHolds folder not found" -ForegroundColor Yellow
    }

    # Find Purges Folder ID
    $PurgesString = $folderQueries | Select-String "/Purges"
    if ($PurgesString) {
        $pattern = "folderid=([A-Za-z0-9+/=]+)"
        $match = [regex]::Match($PurgesString, $pattern)

        if ($match.Success) {
            $Global:PurgesFolderId = $match.value
            Write-host "Found Purges Folder            - $Global:PurgesFolderId" -ForegroundColor Cyan
        }
    } else {
        Write-Host "Warning: Purges folder not found" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Folder IDs found successfully." -ForegroundColor Green
    Write-Host "`n`n`n`n`n"
}
function Run-ComplianceSearch {
    Write-Host "Step [3]: Create Compliance Search for '$Global:EmployeeEmail'" -ForegroundColor Green
    Write-Host "————————————————————————————————————————————————————————————————————————————————————————————————"
    Write-Host ""
    Write-Host "Starting Compliance Search..." -ForegroundColor Yellow
    Write-Host ""

    if ([string]::IsNullOrEmpty($Global:TechEmail) -or [string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Write-Host "Please set both Admin and Employee email addresses first." -ForegroundColor Red
        #Write-Host ""
        ###[void][System.Console]::ReadKey($true)
        return
    }

    if ([string]::IsNullOrEmpty($Global:RecoverableItemFolderId) -or
        [string]::IsNullOrEmpty($Global:DeletionsFolderId) -or
        [string]::IsNullOrEmpty($Global:DiscoveryHoldsFolderId) -or
        [string]::IsNullOrEmpty($Global:SearchDiscoveryHoldsFolderId) -or
        [string]::IsNullOrEmpty($Global:SubstrateHoldsFolderId)) {
        Write-Host "Please find folder IDs first." -ForegroundColor Red
        #Write-Host ""
        ###[void][System.Console]::ReadKey($true)
        return
    }

    $Global:SearchName = "$Global:EmployeeEmail-Purge"

    Connect-IPPSSession -ShowBanner:$false -UserPrincipalName $Global:TechEmail
    Connect-ExchangeOnline  -ShowBanner:$false -UserPrincipalName $Global:TechEmail
    New-ComplianceSearch -Name $Global:SearchName -ExchangeLocation $Global:EmployeeEmail -ContentMatchQuery "$Global:RecoverableItemFolderId OR $Global:DeletionsFolderId OR $Global:DiscoveryHoldsFolderId OR $Global:SearchDiscoveryHoldsFolderId OR $Global:SubstrateHoldsFolderId OR $Global:PurgesFolderId"
    Start-Sleep -Seconds 5
    Start-ComplianceSearch -Identity $Global:SearchName

    Do {
        Start-Sleep -Seconds 5
        $ComplianceSearchStatus = Get-ComplianceSearch -Identity $Global:SearchName
        If (($ComplianceSearchStatus).Status -eq "NotStarted" -or ($ComplianceSearchStatus).Status -eq "Starting") {
            Write-Host "Search job is still Running, please be patient ............" -ForegroundColor Cyan
        }
    } Until (($ComplianceSearchStatus).Status -eq "Completed")

    Write-Host "`n Search complete!!!!!" -ForegroundColor Cyan
    Write-Host ""

    $SearchDetails = Get-ComplianceSearch -Identity $Global:SearchName

    if (($SearchDetails).Status -eq "Completed") {
        Write-Host "Pulling Compliance Search Results...`n" -ForegroundColor Cyan
        Write-Host "Search Name: $($SearchDetails.Name)" -ForegroundColor DarkGreen
        Write-Host "Status: $($SearchDetails.Status)" -ForegroundColor DarkGreen
        $sizeInGB = [math]::Round($SearchDetails.Size / 1GB, 2)
        Write-Host "Search Content Size: $($sizeInGB) GB`n" -ForegroundColor DarkGreen
    }

    Write-Host "Compliance search completed successfully." -ForegroundColor Green
    write-host "`n`n`n`n`n"
    #Write-Host ""
    ###[void][System.Console]::ReadKey($true)
}
function Run-PurgeOperation {
    if ([string]::IsNullOrEmpty($Global:SearchName)) {
        Write-Host "Please run compliance search first." -ForegroundColor Red
        #Write-Host ""
        ##[void][System.Console]::ReadKey($true)
        return
    }

    Do {

        Write-Host "Step [4]: Start Compliance Search Action to purge E-Mails" -ForegroundColor Green
        Write-Host "————————————————————————————————————————————————————————————————————————————————————————————————" -ForegroundColor DarkYellow
        Write-Host "Search Name   -   Purge Name  -  Purge Type         -             Running As      -      Purge Status" -ForegroundColor Green
        Write-Host "————————————————————————————————————————————————————————————————————————————————————————————————" -ForegroundColor DarkYellow
        New-ComplianceSearchAction -SearchName "${Global:SearchName}" -Purge -PurgeType HardDelete -Confirm:$false
        $ComplianceSearchActionStatus = (Get-ComplianceSearchAction -Identity "${Global:SearchName}_Purge").Status

        while ($ComplianceSearchActionStatus -eq "Starting" -or $ComplianceSearchActionStatus -eq "InProgress") {
            Write-Host "Search Action is still running, please be patient......" -ForegroundColor Cyan
            Start-Sleep -Seconds 10
            $ComplianceSearchActionStatus = (Get-ComplianceSearchAction -Identity "${Global:SearchName}_Purge").Status
        }

        $ComplianceSearchActionStatus = (Get-ComplianceSearchAction -Identity "${Global:SearchName}_Purge").Status
        if ($ComplianceSearchActionStatus -eq "Completed") {
            Write-Host ""
            Write-Host "Compliance Search Action completed. See the purge details below..." -ForegroundColor Cyan
        }

        $SearchName  = (Get-ComplianceSearchAction -Identity "${Global:SearchName}_Purge" | select *).SearchName
        $PurgeName   = (Get-ComplianceSearchAction -Identity "${Global:SearchName}_Purge" | select *).Name
        $ItemsPurged = (Get-ComplianceSearchAction -Identity "${Global:SearchName}_Purge" | select *).Results
        $PurgeAction = (Get-ComplianceSearchAction -Identity "${Global:SearchName}_Purge" | select *).Action
        $PurgeErrors = (Get-ComplianceSearchAction -Identity "${Global:SearchName}_Purge" | select *).Errors

        Write-Host ""
        Write-Host "######################### Purge Job Details #######################################################################################################" -ForegroundColor Green
        Write-Host ""
        Write-Host "Search Name:       | $SearchName" -ForegroundColor Green
        Write-Host "Purge Name:        | $PurgeName" -ForegroundColor Green
        Write-Host "Purge Errors:      | $PurgeErrors" -ForegroundColor Green
        Write-Host ""
        Write-Host "Results:           | $ItemsPurged" -ForegroundColor Green
        Write-Host ""
        Write-Host ""
        Write-Host ""
        Write-Host "####################################################################################################################################################" -ForegroundColor Green

        Write-host ""
        Write-host ""
        Start-ManagedFolderAssistant -Identity $Global:EmployeeEmail
        Write-host ""
        Write-host ""
        Write-host "Checking folder size to see if additional search purges are needed......." -ForegroundColor yellow

        $folderSizes = Get-MailboxFolderStatistics -Identity $Global:EmployeeEmail -FolderScope RecoverableItems | Select-Object -ExpandProperty FolderSize
        $totalSizeGB = 0.0

        foreach ($folderSizeString in $folderSizes) {
            $sizeStringClean = $folderSizeString -replace '\s*\(.*\)' # Remove byte count
            $sizeToAddGB = 0.0

            if ($sizeStringClean -match '(\d+(\.\d+)?)\s*GB') {
                $sizeToAddGB = [double]$matches[1]
            }
            elseif ($sizeStringClean -match '(\d+(\.\d+)?)\s*MB') {
                $sizeToAddGB = ([double]$matches[1]) / 1024.0
            }
            elseif ($sizeStringClean -match '(\d+(\.\d+)?)\s*KB') {
                $sizeToAddGB = ([double]$matches[1]) / (1024.0 * 1024.0)
            }
             elseif ($sizeStringClean -match '(\d+)\s*B') {
                $sizeToAddGB = ([double]$matches[1]) / (1024.0 * 1024.0 * 1024.0)
            }
            $totalSizeGB += $sizeToAddGB
        }
        # Round Number
        $totalSizeGB_Rounded = [math]::Round($totalSizeGB, 2)
        Write-Host "Total Recoverable Items Size (Calculated from strings): $totalSizeGB_Rounded GB"
        $sizeInGB = $totalSizeGB_Rounded

        if ($sizeInGB -ge 20) {
            Write-Host ""
            Write-Host "The folder size is still $sizeInGB GB, restarting the job to continue clearing items." -ForegroundColor cyan
            get-mailboxFolderStatistics $Global:EmployeeEmail -FolderScope recoverableitems | Select identity, foldersize, hostname
            Write-Host "`n`n`n`n"
            Remove-ComplianceSearchAction -Identity "${Global:SearchName}_Purge" -Confirm:$false
            #Write-Host "———————————————————————————————————————————————————————————————————————————————————————————————————————————————————————" -ForegroundColor Green
            #$Results = Get-MailboxFolderStatistics -Identity $Global:EmployeeEmail -FolderScope RecoverableItems | Select-Object FolderPath, ItemsInFolder, FolderSize | Format-List
            #write-host $Results -ForegroundColor Green
            #Write-Host "———————————————————————————————————————————————————————————————————————————————————————————————————————————————————————" -ForegroundColor Green
        }

        Start-Sleep -Seconds 2

    } While ($sizeInGB -ge 20)

    Write-Host "Process completed, folder size is below 20GB." -ForegroundColor Green
    ##Write-Host ""
    ###[void][System.Console]::ReadKey($true)
}
function Restore-MailboxSettings {
    if ([string]::IsNullOrEmpty($Global:TechEmail) -or [string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Write-Host "Please set both Admin and Employee email addresses first." -ForegroundColor Red
        #Write-Host ""
        ##[void][System.Console]::ReadKey($true)
        return
    }

    Write-Host "Reverting changes implemented earlier in the User Mailbox Prep process "
    Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $Global:TechEmail

    Write-Host "Enabling WS, Active Sync, MAPI, OWA, IMAP, and POP for $Global:EmployeeEmail........" -ForegroundColor Cyan
    Set-CASMailbox $Global:EmployeeEmail -EwsEnabled $true -ActiveSyncEnabled $true -MAPIEnabled $true -OWAEnabled $true -ImapEnabled $true -PopEnabled $true -ErrorAction SilentlyContinue

    Write-Host "Reverting the retention window back to 30-days for $Global:EmployeeEmail......" -ForegroundColor Cyan
    Set-Mailbox $Global:EmployeeEmail -RetainDeletedItemsFor 30 -ErrorAction SilentlyContinue

    Write-Host "Enabling single-item-recovery for $Global:EmployeeEmail......" -ForegroundColor Cyan
    Set-Mailbox $Global:EmployeeEmail -SingleItemRecoveryEnabled $true -ErrorAction SilentlyContinue

    Write-Host "Enabling Manage Folder Assistant for $Global:EmployeeEmail......" -ForegroundColor Cyan
    Set-Mailbox $Global:EmployeeEmail -ElcProcessingDisabled $false -ErrorAction SilentlyContinue

    Write-Host "Recoverable Items are now cleared and mailbox settings restored." -ForegroundColor Cyan
    #Write-Host ""
    ##[void][System.Console]::ReadKey($true)
}


# --- Main Script Execution functions ---
function Search-RecoverableItems {
    Check-PowerShell7
    Install-RequiredModules

    if ([string]::IsNullOrEmpty($Global:TechEmail)) {
        Set-AdminEmail
    }

    if ([string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Set-EmployeeEmail
    }

    # Connect to Exchange Online
    Connect-ExchangeOnline -ShowBanner:$false -UserPrincipalName $Global:TechEmail

    # Get folder sizes for the recoverable items
    Write-Host "Searching for recoverable items in $Global:EmployeeEmail mailbox..." -ForegroundColor Cyan

    $folderStats = Get-MailboxFolderStatistics -Identity $Global:EmployeeEmail -FolderScope RecoverableItems

        # Display results
        Write-Host "`nRecoverable Items folders for $Global:EmployeeEmail:" -ForegroundColor Green
    Write-Host "=======================================================" -ForegroundColor Green

    $totalSize = 0
    $totalItems = 0

    foreach ($folder in $folderStats) {
        # Convert folder size to GB if possible
        $sizeInGB = 0
        $folderSize = $folder.FolderSize
        $size = $folderSize -replace '\s*\(.*\)'

        if ($size -match '(\d+(\.\d+)?)\s*GB') {
            $sizeInGB = [double]$matches[1]
        }
        elseif ($size -match '(\d+(\.\d+)?)\s*MB') {
            $sizeInGB = [double]$matches[1] / 1024
        }

        $totalSize += $sizeInGB
        $totalItems += $folder.ItemsInFolder

        Write-Host "Folder Path: $($folder.FolderPath)" -ForegroundColor White
        Write-Host "  Size: $($folder.FolderSize)" -ForegroundColor Cyan
        Write-Host "  Items: $($folder.ItemsInFolder)" -ForegroundColor Cyan
        Write-Host "  Created: $($folder.CreationTime)" -ForegroundColor Gray
        Write-Host "------------------------------------------------------" -ForegroundColor DarkGray
    }

    Write-Host "`nSummary:" -ForegroundColor Yellow
    Write-Host "  Total Folders: $($folderStats.Count)" -ForegroundColor Yellow
    Write-Host "  Total Items: $totalItems" -ForegroundColor Yellow
    Write-Host "  Total Size: $($totalSize.ToString('0.00')) GB" -ForegroundColor Yellow

    Write-Host "`nSearch completed successfully!" -ForegroundColor Green
    Write-Host "Press enter to conitnue......"
    [void][System.Console]::ReadKey($true)
}
function Purge-RecoverableItems {
    Check-PowerShell7
    Install-RequiredModules

    if ([string]::IsNullOrEmpty($Global:TechEmail)) {
        Set-AdminEmail
    }

    if ([string]::IsNullOrEmpty($Global:EmployeeEmail)) {
        Set-EmployeeEmail
    }

    # Confirm the purge operation
     clear-host
    Write-Host "`n⚠️ WARNING: This will permanently delete all items in the recoverable items folder for $Global:EmployeeEmail." -ForegroundColor Red
    Write-Host "This action cannot be undone and will remove all recovery options for deleted items." -ForegroundColor Red
    $confirm = Read-Host "Are you sure you want to proceed? (Y/N)"
      
    if ($confirm -ne "Y" -and $confirm -ne "y") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        #Write-Host ""
        ##[void][System.Console]::ReadKey($true)
        return
    }

    clear-host
    Show-Logo
    Write-Host "================================================================================"
    Write-Host "Starting purge operation for $Global:EmployeeEmail..." -ForegroundColor Green
    Write-Host "================================================================================"
    Write-Host ""
    # Prepare user mailbox for purge
    Prepare-UserMailbox

    # Find folder IDs
    clear-host
    Show-Logo
    Find-FolderIDs

    # Run compliance search
    clear-host
    Show-Logo
    Run-ComplianceSearch

    # Run purge operation
    clear-host
    Show-Logo
    Run-PurgeOperation

    # Restore mailbox settings
    clear-host
    Show-Logo
    Restore-MailboxSettings

    Write-Host "`nPurge operation completed successfully!" -ForegroundColor Green
    #Write-Host ""
    #[void][System.Console]::ReadKey($true)

}
#EndRegion-Functions


# --- Main Script Execution ---
while ($true) {

    # Collect credentials
    if ($null -eq $Global:Credentials) {
        Clear-Host
        Write-Host "Hello, before we proceed, domain administrator credentials are required (e.g., OTLS1\xadm_$($ENV:UserName))."  -ForegroundColor Blue
        Write-Host "Please enter your admin credentials below."  -ForegroundColor Blue
        $Global:Credentials = Get-Credential
        Write-Host "Awesome! You are running the script as $($Global:Credentials.username)!" -ForegroundColor Green
        write-host "`n"
    }


    Clear-Host
    Show-Logo
    Write-Host "`n"

    # Show Main Menu and get choice
    $choice = Show-MainMenu

    # Process menu choice
    switch ($choice) {
        "1" { Search-RecoverableItems }
        "2" { Purge-RecoverableItems }
        "Q" {
            Write-Host "Exiting script. Goodbye!" -ForegroundColor Cyan
            exit
        }
        "q" {
            Write-Host "Exiting script. Goodbye!" -ForegroundColor Cyan
            exit
        }
        default {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
}