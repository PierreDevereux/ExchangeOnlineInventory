<#
.SYNOPSIS
    Inventories all Exchange Online mailboxes (user, shared, room, equipment, etc.)
    and their in-place archives, and exports size, item count, and oldest/newest
    message dates to a CSV file for reconciliation against a Mimecast report.

.DESCRIPTION
    This script connects to Exchange Online, enumerates every mailbox in the tenant,
    and for each one collects:
        - Mailbox           : Primary SMTP address of the mailbox
        - MailboxType       : "Primary" or "Archive"
        - Messages          : Total item count (ItemCount)
        - Size (bytes)      : Total mailbox size in bytes
        - Size (GB)         : Total mailbox size in gigabytes (rounded to 2 decimals)
        - Oldest Message    : Received date of the oldest item across all folders
        - Newest Message    : Received date of the newest item across all folders

    Each mailbox produces one "Primary" row. If the mailbox has an active in-place
    archive, a second "Archive" row is produced with the same Mailbox value so the
    two can be matched during comparison.

    ================================================================================
    INSTRUCTIONS FOR THE PERSON RUNNING THIS SCRIPT
    ================================================================================

    PREREQUISITES
    -------------
    1. A Windows machine with PowerShell 5.1 or PowerShell 7+.

    2. An Exchange Online / Microsoft 365 account that has at least the
       "View-Only Recipients" role (the "Global Reader", "Exchange
       Administrator", or "Global Administrator" roles also work). This account
       is only used to READ mailbox statistics; the script makes no changes.

    3. The Exchange Online Management PowerShell module. To install it, open
       PowerShell and run the following command once:

           Install-Module ExchangeOnlineManagement -Scope CurrentUser

       If prompted to trust the PSGallery repository, answer "Yes" (Y).
       (If the machine has an older version, update it with:
           Update-Module ExchangeOnlineManagement )

    HOW TO RUN
    ----------
    1. Save this script somewhere convenient, e.g. C:\Temp\ExchangeOnlineInventory.ps1

    2. Open PowerShell and change to that folder, e.g.:

           cd C:\Temp

    3. If script execution is blocked, allow it for this session only by running:

           Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

    4. Run the script. You will be prompted to sign in to Microsoft 365 in a
       browser window (this supports multi-factor authentication):

           .\ExchangeOnlineInventory.ps1

       To choose where the CSV is written, supply the -OutputPath parameter:

           .\ExchangeOnlineInventory.ps1 -OutputPath "C:\Temp\Inventory.csv"

    5. When it finishes, the script prints the full path to the generated CSV
       file. Please send that CSV back to Matthew Levy.

    NOTES ON RUNTIME
    ----------------
    - The script may take a long time on large tenants because it reads folder
      statistics for every mailbox and every archive. This is expected.
    - A progress bar shows how far through the mailbox list the script is.
    - If a single mailbox fails to return statistics, the script logs a warning
      and continues with the remaining mailboxes.

.PARAMETER OutputPath
    Full path (including file name) for the CSV output. If omitted, the file is
    written to the current directory as:
        ExchangeOnlineInventory_<yyyyMMdd_HHmmss>.csv

.EXAMPLE
    .\ExchangeOnlineInventory.ps1

    Connects interactively, inventories all mailboxes, and writes a timestamped
    CSV to the current directory.

.EXAMPLE
    .\ExchangeOnlineInventory.ps1 -OutputPath "C:\Reports\EXO_Inventory.csv"

    Connects interactively and writes the CSV to the specified path.

.NOTES
    Author  : Matthew Levy (MVP)
    Purpose : Mailbox inventory for reconciliation against a Mimecast report.
    Version : 1.0

    This script only READS data from Exchange Online. It does not modify, move,
    or delete any mailbox or message.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("ExchangeOnlineInventory_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss')))
)

# Tracks whether THIS script opened the EXO connection, so we only disconnect
# a session we created and leave any pre-existing session intact.
$script:ConnectionOpenedByScript = $false

function Convert-SizeToBytes {
    # Parses the "12.34 GB (13,247,905,792 bytes)" string returned by the
    # statistics cmdlets into a plain [long] byte count. Returns 0 when empty.
    param($TotalItemSize)

    if ($null -eq $TotalItemSize) { return [long]0 }

    $text = $TotalItemSize.ToString()
    if ([string]::IsNullOrWhiteSpace($text)) { return [long]0 }

    if ($text -match '\(([\d,]+)\s*bytes\)') {
        return [long]($matches[1] -replace ',', '')
    }

    return [long]0
}

function Get-MailboxDateRange {
    # Returns a PSCustomObject with Oldest and Newest received dates across all
    # folders for a mailbox (or its archive). Nulls are ignored.
    param(
        [string]$Identity,
        [switch]$Archive
    )

    $oldest = $null
    $newest = $null

    $folderParams = @{
        Identity                    = $Identity
        IncludeOldestAndNewestItems = $true
        ErrorAction                 = 'Stop'
    }
    if ($Archive) { $folderParams['Archive'] = $true }

    $folders = Get-EXOMailboxFolderStatistics @folderParams

    foreach ($folder in $folders) {
        if ($folder.OldestItemReceivedDate) {
            if ($null -eq $oldest -or $folder.OldestItemReceivedDate -lt $oldest) {
                $oldest = $folder.OldestItemReceivedDate
            }
        }
        if ($folder.NewestItemReceivedDate) {
            if ($null -eq $newest -or $folder.NewestItemReceivedDate -gt $newest) {
                $newest = $folder.NewestItemReceivedDate
            }
        }
    }

    return [PSCustomObject]@{
        Oldest = $oldest
        Newest = $newest
    }
}

function New-InventoryRow {
    # Builds a single CSV row (primary or archive) for a mailbox.
    param(
        [string]$Mailbox,
        [ValidateSet('Primary', 'Archive')]
        [string]$MailboxType,
        [string]$Identity
    )

    $statParams = @{
        Identity    = $Identity
        Properties  = 'ItemCount', 'TotalItemSize'
        ErrorAction = 'Stop'
    }
    if ($MailboxType -eq 'Archive') { $statParams['Archive'] = $true }

    $stats = Get-EXOMailboxStatistics @statParams

    $bytes = Convert-SizeToBytes -TotalItemSize $stats.TotalItemSize
    $dates = Get-MailboxDateRange -Identity $Identity -Archive:($MailboxType -eq 'Archive')

    return [PSCustomObject]@{
        'Mailbox'         = $Mailbox
        'MailboxType'     = $MailboxType
        'Messages'        = [int]$stats.ItemCount
        'Size (bytes)'    = $bytes
        'Size (GB)'       = [math]::Round($bytes / 1GB, 2)
        'Oldest Message'  = $dates.Oldest
        'Newest Message'  = $dates.Newest
    }
}

# --- Ensure the Exchange Online Management module is available ----------------
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Error @"
The 'ExchangeOnlineManagement' module is not installed.
Install it by running the following command, then re-run this script:

    Install-Module ExchangeOnlineManagement -Scope CurrentUser
"@
    return
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

# --- Connect to Exchange Online (reuse an existing session if present) --------
try {
    $existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue

    if (-not $existingConnection) {
        Write-Host 'Connecting to Exchange Online. A sign-in window will open...' -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $script:ConnectionOpenedByScript = $true
    }
    else {
        Write-Host 'Using existing Exchange Online connection.' -ForegroundColor Cyan
    }
}
catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    return
}

# --- Inventory ----------------------------------------------------------------
$results = [System.Collections.Generic.List[object]]::new()

try {
    Write-Host 'Retrieving mailbox list...' -ForegroundColor Cyan
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties ArchiveStatus -ErrorAction Stop

    $total = @($mailboxes).Count
    Write-Host "Found $total mailbox(es). Collecting statistics..." -ForegroundColor Cyan

    $index = 0
    foreach ($mbx in $mailboxes) {
        $index++
        $address = $mbx.PrimarySmtpAddress

        Write-Progress -Activity 'Inventorying mailboxes' `
            -Status "$index of $total : $address" `
            -PercentComplete (($index / [math]::Max($total, 1)) * 100)

        # Primary mailbox
        try {
            $results.Add((New-InventoryRow -Mailbox $address -MailboxType 'Primary' -Identity $mbx.Identity))
        }
        catch {
            Write-Warning "Failed to read primary statistics for '$address': $($_.Exception.Message)"
        }

        # Archive mailbox (only if an active archive exists)
        if ($mbx.ArchiveStatus -eq 'Active') {
            try {
                $results.Add((New-InventoryRow -Mailbox $address -MailboxType 'Archive' -Identity $mbx.Identity))
            }
            catch {
                Write-Warning "Failed to read archive statistics for '$address': $($_.Exception.Message)"
            }
        }
    }

    Write-Progress -Activity 'Inventorying mailboxes' -Completed
}
catch {
    Write-Error "Failed while retrieving mailboxes: $($_.Exception.Message)"
}

# --- Export -------------------------------------------------------------------
if ($results.Count -gt 0) {
    try {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Done. {0} row(s) written to: {1}" -f $results.Count, (Resolve-Path -Path $OutputPath)) -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to write CSV to '$OutputPath': $($_.Exception.Message)"
    }
}
else {
    Write-Warning 'No mailbox data was collected; CSV was not created.'
}

# --- Disconnect (only the session this script created) ------------------------
if ($script:ConnectionOpenedByScript) {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host 'Disconnected from Exchange Online.' -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to disconnect cleanly: $($_.Exception.Message)"
    }
}
