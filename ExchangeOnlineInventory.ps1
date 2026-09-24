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
    - Exchange Online throttles heavy reporting workloads. The script detects
      transient throttling/timeout/connection errors and automatically retries
      each affected call with an increasing (exponential) back-off delay, and
      re-establishes the Exchange Online session if it drops mid-run.
    - If a mailbox still cannot be read after all retries, the script logs a
      warning, records the mailbox in a companion "*_failures.log" file next to
      the CSV, and continues with the remaining mailboxes. Always check that log
      after a run so you know the report is complete.
    - On very large tenants you can slow the script down with -ThrottleDelayMs
      to reduce the chance of being throttled in the first place.

.PARAMETER OutputPath
    Full path (including file name) for the CSV output. If omitted, the file is
    written to the current directory as:
        ExchangeOnlineInventory_<OrganisationName>_<yyyyMMdd_HHmmss>.csv

.PARAMETER MaxRetries
    Maximum number of automatic retries per failed statistics call when a
    transient (throttling/timeout/connection) error occurs. Default is 5.
    Set to 0 to disable retries.

.PARAMETER ThrottleDelayMs
    Optional pause, in milliseconds, inserted after each mailbox is processed.
    Use this on large tenants to pace the script and avoid triggering dynamic
    throttling. Default is 0 (no pause).

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
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 20)]
    [int]$MaxRetries = 5,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 60000)]
    [int]$ThrottleDelayMs = 0
)

# Tracks whether THIS script opened the EXO connection, so we only disconnect
# a session we created and leave any pre-existing session intact.
$script:ConnectionOpenedByScript = $false

function Test-TransientError {
    # Returns $true for errors that are worth retrying: throttling, timeouts,
    # service-busy, and dropped/expired connections.
    param($ErrorRecord)

    $msg = "$($ErrorRecord.Exception.Message)"
    return ($msg -match '(?i)(429|503|throttl|ServerBusy|TooManyRequests|timed out|timeout|operation has timed|service is unavailable|connection was closed|session .*expired|token .*expired|unable to connect)')
}

function Get-ServerBackoffSeconds {
    # Best-effort extraction of a server-suggested back-off from the error text.
    # Returns 0 when none is found.
    param($ErrorRecord)

    $msg = "$($ErrorRecord.Exception.Message)"
    if ($msg -match 'BackOffMilliseconds\D+(\d+)') {
        return [int][math]::Ceiling([int]$matches[1] / 1000)
    }
    if ($msg -match 'retry after\D+(\d+)\s*second') {
        return [int]$matches[1]
    }
    return 0
}

function Restore-EXOConnection {
    # Re-establishes the Exchange Online session if it has dropped. Uses cached
    # tokens when possible; may prompt for sign-in if the token has fully expired.
    try {
        $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if (-not $conn) {
            Write-Warning 'Exchange Online session lost; attempting to reconnect...'
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Reconnect attempt failed: $($_.Exception.Message)"
    }
}

function Invoke-WithRetry {
    # Runs a script block, retrying transient failures with exponential back-off
    # (capped at 60s) and reconnecting the session between attempts if needed.
    # Non-transient errors, or errors past MaxRetries, are re-thrown.
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$OperationName = 'operation',
        [int]$MaxRetries = 5
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $ScriptBlock
        }
        catch {
            $err = $_
            if (-not (Test-TransientError -ErrorRecord $err) -or $attempt -gt $MaxRetries) {
                throw
            }

            $backoff = [int][math]::Min([math]::Pow(2, $attempt), 60)
            $serverBackoff = Get-ServerBackoffSeconds -ErrorRecord $err
            if ($serverBackoff -gt $backoff) { $backoff = $serverBackoff }

            Write-Warning ("{0}: transient error (attempt {1} of {2}); retrying in {3}s -> {4}" -f `
                    $OperationName, $attempt, $MaxRetries, $backoff, $err.Exception.Message)

            Restore-EXOConnection
            Start-Sleep -Seconds $backoff
        }
    }
}

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

    $folders = Invoke-WithRetry -OperationName "Get-EXOMailboxFolderStatistics ($Identity)" -MaxRetries $script:MaxRetries -ScriptBlock {
        Get-EXOMailboxFolderStatistics @folderParams
    }

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

    $stats = Invoke-WithRetry -OperationName "Get-EXOMailboxStatistics ($MailboxType $Mailbox)" -MaxRetries $script:MaxRetries -ScriptBlock {
        Get-EXOMailboxStatistics @statParams
    }

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

# --- Build default output path using the organisation name --------------------
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $orgName = 'UnknownOrg'
    try {
        $orgName = (Get-OrganizationConfig -ErrorAction Stop).Name
    }
    catch {
        Write-Warning "Could not determine organisation name; using '$orgName'."
    }

    # Strip characters that are not valid in Windows file names.
    $safeOrg = ($orgName -replace '[\\/:*?"<>|]', '_')
    $fileName = 'ExchangeOnlineInventory_{0}_{1}.csv' -f $safeOrg, (Get-Date -Format 'yyyyMMdd_HHmmss')
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath $fileName
}

# Companion failure log lives next to the CSV.
$outDir = Split-Path -Parent $OutputPath
if ([string]::IsNullOrEmpty($outDir)) { $outDir = (Get-Location).Path }
$failureLogPath = Join-Path -Path $outDir -ChildPath (([System.IO.Path]::GetFileNameWithoutExtension($OutputPath)) + '_failures.log')

# --- Inventory ----------------------------------------------------------------
$results = [System.Collections.Generic.List[object]]::new()
$failures = [System.Collections.Generic.List[object]]::new()

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
            Write-Warning "Failed to read primary statistics for '$address' after retries: $($_.Exception.Message)"
            $failures.Add([PSCustomObject]@{
                    Timestamp   = (Get-Date)
                    Mailbox     = $address
                    MailboxType = 'Primary'
                    Error       = $_.Exception.Message
                })
        }

        # Archive mailbox (only if an active archive exists)
        if ($mbx.ArchiveStatus -eq 'Active') {
            try {
                $results.Add((New-InventoryRow -Mailbox $address -MailboxType 'Archive' -Identity $mbx.Identity))
            }
            catch {
                Write-Warning "Failed to read archive statistics for '$address' after retries: $($_.Exception.Message)"
                $failures.Add([PSCustomObject]@{
                        Timestamp   = (Get-Date)
                        Mailbox     = $address
                        MailboxType = 'Archive'
                        Error       = $_.Exception.Message
                    })
            }
        }

        if ($ThrottleDelayMs -gt 0) { Start-Sleep -Milliseconds $ThrottleDelayMs }
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

# --- Write failure log (if any mailboxes could not be read) -------------------
if ($failures.Count -gt 0) {
    try {
        $logLines = $failures | ForEach-Object {
            '{0:u}  [{1}]  {2}  ::  {3}' -f $_.Timestamp, $_.MailboxType, $_.Mailbox, $_.Error
        }
        $header = @(
            'Exchange Online Inventory - failure log',
            "Generated : $(Get-Date -Format 'u')",
            "Failures  : $($failures.Count)",
            ('-' * 60)
        )
        Set-Content -Path $failureLogPath -Value ($header + $logLines) -Encoding UTF8
        Write-Warning ("{0} mailbox/archive read(s) failed after retries. See: {1}" -f $failures.Count, $failureLogPath)
    }
    catch {
        Write-Warning "Failed to write failure log to '$failureLogPath': $($_.Exception.Message)"
    }
}
else {
    Write-Host 'No mailbox read failures.' -ForegroundColor Green
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
