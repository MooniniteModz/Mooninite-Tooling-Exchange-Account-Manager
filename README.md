## Mooninite Tooling - Exchange Recoverable Items Manager


PowerShell tool for searching and purging the Recoverable Items folder on Exchange Online mailboxes. Built to deal with oversized recoverable items folders that pile up from compliance holds, litigation holds, and general mailbox bloat across multi-org Exchange environments.



## What It Does - Two things, and it does them well:

- Search: Connects to Exchange Online, pulls Get-MailboxFolderStatistics scoped to RecoverableItems, and gives you a per-folder breakdown (path, size, item count, creation date) with a total summary. Quick way to - see how bad things are before you commit to a purge.
- Purge: The full pipeline. Preps the mailbox (disables EWS, ActiveSync, MAPI, OWA, IMAP, POP, turns off single-item recovery, kills holds), resolves all the hidden folder IDs (Deletions, DiscoveryHolds, SubstrateHolds, Purges, etc.), runs a Compliance Search against those folders, executes a HardDelete purge action, and loops until the recoverable items folder drops below 20GB. Then it puts everything back the way it was.



## The Purge Pipeline (step by step)
- Prepare Mailbox → Find Folder IDs → Compliance Search → HardDelete Purge (loop) → Restore Settings

- Prepare: Disables all client access protocols, sets retention to 0, disables single-item recovery, removes litigation holds, stops the Managed Folder Assistant
- Folder IDs: Converts Exchange folder IDs from Base64 to hex-encoded format compatible with Purview Compliance Search queries
- Search: Creates and runs a New-ComplianceSearch targeting all resolved recoverable item subfolders
- Purge: Runs New-ComplianceSearchAction -Purge -PurgeType HardDelete, checks remaining size, and loops if still over 20GB (Exchange caps purge operations at ~10 items per mailbox per batch, so large folders need multiple passes)
- Restore: Re-enables all protocols, sets retention back to 30 days, re-enables single-item recovery and the Managed Folder Assistant


# Requirements

- PowerShell 7.2+ — the script will auto-download and install PS7 if you're running it from Windows PowerShell 5.x
- ExchangeOnlineManagement module — auto-installed from PSGallery if missing
- Exchange Admin permissions — your admin account needs Compliance Search and Compliance Search Action roles
- PIM activation — if you're running Privileged Identity Management, make sure the required roles are active before launching



## Usage
# powershell.\Clean-RecoverableItems.ps1
- On first run, you'll be prompted for domain admin credentials (format: DOMAIN\xadm_username), then your Exchange admin email, and then the target employee mailbox. From there, the menu handles the rest.
Main Menu:
1. Search for recoverable items on user account
2. Purge recoverable items folder
Q. Quit
Granting Compliance Permissions

The admin account running the search/purge needs these roles assigned in the Purview Compliance Center:
- Compliance Search Administrator
- Compliance Administrator (or equivalent that grants New-ComplianceSearchAction)

- If you're using PIM, activate both roles before running the script.
- Purview Compliance Center Syntax (v5.0)
- As of v5.0, the script uses the updated Purview Compliance Center syntax for content match queries. The old FolderID:"value" format has been replaced with FolderID=value to match current Microsoft requirements.


The purge loop threshold is 20GB. If the recoverable items folder is under 20GB after a purge pass, the script considers the job done and restores mailbox settings.
Exchange Online caps HardDelete purge operations, so very large folders (50GB+) will take multiple passes. The script handles this automatically but it can take a while. Go get coffee.
The Managed Folder Assistant is kicked via Start-ManagedFolderAssistant between purge passes to help Exchange process the deletions faster.
Compliance Search creation and execution is async — the script polls status every 5 seconds until completion.


License
Do whatever you want with it. If it breaks your Exchange environment, that's between you and Microsoft support.
