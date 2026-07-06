<#
.SYNOPSIS
    Checks the local Windows machine against a company IT compliance checklist and
    writes a Markdown report.

.DESCRIPTION
    Test-ITCompliance is a CHECK / AUDIT script only. It reads, detects, and measures
    the state of the machine against nine IT checklist items (storage, username
    convention, corporate network, OS update freshness, battery health, software
    licensing indicators, bloatware detection, junk file sizing, and antivirus
    status) and writes the results to a single human-readable Markdown report.

    HARD CONSTRAINT: this script never deletes, uninstalls, moves, or modifies any
    file, program, or setting on the machine. It only reads/detects/measures and
    reports. Where the checklist talks about "removing" unused programs or junk
    files, this script only DETECTS and REPORTS candidates/sizes -- it explicitly
    states in the report that no removal was performed and that acting on the list
    is the IT team's decision. The only file this script creates is its own Markdown
    report; it may also use a short-lived temporary file for the battery report XML,
    which is always deleted again before the script exits (even on error).

    Every check is gathered independently and wrapped in its own function with a
    try/catch, so a failure in one check (for example, a permissions issue, a
    missing cmdlet, or a hung external tool) will not prevent the rest of the report
    from being produced. A check that could not be completed reports verdict
    UNKNOWN along with the reason instead of crashing the script.

.PARAMETER OutputPath
    Directory in which the generated Markdown report will be saved. Defaults to the
    current user's Documents folder.

.PARAMETER ExpectedUsernamePattern
    Regex pattern describing the company's Windows username naming convention. If
    not supplied, the username check is reported as SKIPPED (the current username
    is still shown as INFO) -- this script never guesses a naming convention.

.PARAMETER ExpectedDomain
    The Active Directory domain name the machine is expected to be joined to. Used
    by the Network check.

.PARAMETER ExpectedSSIDs
    One or more allowed corporate Wi-Fi SSID(s). Used by the Network check when the
    machine is currently connected via Wi-Fi.

.PARAMETER StorageFreeThresholdPercent
    Minimum acceptable free-space percentage on any fixed volume. Default 15.

.PARAMETER BatteryHealthThresholdPercent
    Minimum acceptable battery health percentage
    (FullChargeCapacity / DesignCapacity * 100). Default 70.

.PARAMETER JunkFilesWarnThresholdGB
    Reclaimable junk size, in GB, above which the Junk Files check becomes a WARN
    instead of INFO. Default 5. Nothing is ever deleted regardless of this value.

.PARAMETER BloatwareKeywords
    Keyword fragments matched case-insensitively as substrings against installed
    application DisplayNames to flag low-value / bloatware candidates. Override this
    to tune the list to your company's own policy. Defaults to a generic starter
    list (trial security suites, game promos, OEM trialware, toolbars, etc.).

.PARAMETER SkipWindowsUpdateCheck
    Skips the OS-update-freshness check entirely. Useful on offline or
    locked-down machines where Windows Update cannot be reached. The check's
    verdict becomes "N/A" when this switch is used.

.PARAMETER WindowsUpdateTimeoutSeconds
    Maximum time, in seconds, to wait for the Windows Update search (run inside a
    background job) before giving up and reporting UNKNOWN for that check. Default
    45.

.EXAMPLE
    .\Test-ITCompliance.ps1

    Runs the full checklist with all defaults and saves the report to the current
    user's Documents folder. Username and Network checks will be SKIPPED because no
    company policy values were supplied.

.EXAMPLE
    .\Test-ITCompliance.ps1 -ExpectedUsernamePattern '^[a-z]+\.[a-z]+$' -ExpectedDomain 'CORP' -ExpectedSSIDs 'CorpWiFi','CorpWiFi-5G' -OutputPath 'C:\Reports'

    Runs the full checklist against a company naming convention, AD domain, and
    approved Wi-Fi SSIDs, saving the report to C:\Reports.

.NOTES
    Target: Windows PowerShell 5.1+ / PowerShell 7+
    This is a CHECK/AUDIT script only -- it never deletes, uninstalls, moves, or
    modifies anything on the machine.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = [Environment]::GetFolderPath('MyDocuments'),

    [Parameter()]
    [string]$ExpectedUsernamePattern = $null,

    [Parameter()]
    [string]$ExpectedDomain = $null,

    [Parameter()]
    [string[]]$ExpectedSSIDs = $null,

    [Parameter()]
    [int]$StorageFreeThresholdPercent = 15,

    [Parameter()]
    [int]$BatteryHealthThresholdPercent = 70,

    [Parameter()]
    [double]$JunkFilesWarnThresholdGB = 5,

    [Parameter()]
    [string[]]$BloatwareKeywords = @(
        'toolbar',
        'trial',
        'mcafee',
        'norton',
        'candy crush',
        'wildtangent',
        'coupon',
        'search protect',
        'driver booster',
        'pc optimizer',
        'pc cleaner',
        'registry cleaner',
        'game promo',
        'browser assistant',
        'ask toolbar',
        'oem trial',
        'free vpn',
        'weather bug',
        'games for windows'
    ),

    [Parameter()]
    [switch]$SkipWindowsUpdateCheck,

    [Parameter()]
    [int]$WindowsUpdateTimeoutSeconds = 45
)

#region Helper functions

function New-CheckResult {
    <#
        Standard shape for every check's return value so the report assembly code
        can treat all nine checks uniformly.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Verdict,

        [Parameter(Mandatory)]
        [string]$Content
    )

    return [PSCustomObject]@{
        Verdict = $Verdict
        Content = $Content
    }
}

function New-SectionErrorNote {
    <#
        Builds the standard "could not retrieve this section" Markdown note, mirroring
        the style used in Invoke-SystemAudit.ps1.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Reason
    )

    return "_Unable to complete this check: $Reason (may require Administrator, or the relevant service/module is unavailable)_"
}

function ConvertTo-MarkdownTable {
    <#
        Converts an array of objects into a Markdown table using the given column
        headers (property names, in order). If the input array is empty, returns a
        note saying no data was found instead of an empty table.
    #>
    param(
        [Parameter()]
        [object[]]$InputObject,

        [Parameter(Mandatory)]
        [string[]]$Columns,

        [Parameter()]
        [string]$EmptyMessage = '_No data found._'
    )

    if (-not $InputObject -or $InputObject.Count -eq 0) {
        return $EmptyMessage
    }

    $sb = New-Object System.Text.StringBuilder

    $headerLine = '| ' + ($Columns -join ' | ') + ' |'
    $separatorLine = '| ' + (($Columns | ForEach-Object { '---' }) -join ' | ') + ' |'
    [void]$sb.AppendLine($headerLine)
    [void]$sb.AppendLine($separatorLine)

    foreach ($row in $InputObject) {
        $cells = foreach ($col in $Columns) {
            $value = $row.$col
            if ($null -eq $value -or ($value -is [string] -and $value -eq '')) {
                'N/A'
            }
            else {
                # Escape pipe characters so they don't break the table structure.
                ($value.ToString() -replace '\|', '\|')
            }
        }
        [void]$sb.AppendLine('| ' + ($cells -join ' | ') + ' |')
    }

    return $sb.ToString().TrimEnd()
}

function Get-InstalledApplicationsList {
    <#
        Shared helper: enumerates installed applications from the Windows uninstall
        registry keys, the same approach used by Invoke-SystemAudit.ps1's
        Get-InstalledApplicationsSection. Read-only registry queries only.
    #>
    [CmdletBinding()]
    param()

    $uninstallPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $entries = foreach ($path in $uninstallPaths) {
        try {
            Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        }
        catch {
            # Skip a registry hive path that isn't readable; other paths may still
            # succeed.
        }
    }

    $apps = $entries |
        Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne '' } |
        Select-Object -Property `
            @{ Name = 'DisplayName'; Expression = { $_.DisplayName } },
            @{ Name = 'DisplayVersion'; Expression = { $_.DisplayVersion } },
            @{ Name = 'Publisher'; Expression = { $_.Publisher } },
            @{ Name = 'InstallDate'; Expression = { $_.InstallDate } } |
        Sort-Object -Property DisplayName, DisplayVersion, Publisher -Unique |
        Sort-Object -Property DisplayName

    return $apps
}

#endregion Helper functions

#region Check functions

function Test-StorageCompliance {
    <#
        Checklist item 1: Storage. Checks free space % on every fixed volume,
        physical disk health, and takes one rough write-speed sample per fixed
        volume. Read-only: any temp file written for the speed sample is deleted
        immediately after the measurement.
    #>
    [CmdletBinding()]
    param(
        [int]$FreeThresholdPercent
    )

    try {
        $volumeRows = @()
        $anyBelowThreshold = $false

        $volumes = $null
        try {
            $volumes = Get-Volume -ErrorAction Stop | Where-Object { $_.DriveType -eq 'Fixed' -and $_.DriveLetter }
        }
        catch {
            $volumes = $null
        }

        if (-not $volumes) {
            # Fallback to Win32_LogicalDisk if Get-Volume isn't available.
            $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop
            $volumes = foreach ($ld in $logicalDisks) {
                [PSCustomObject]@{
                    DriveLetter   = ($ld.DeviceID -replace ':', '')
                    Size          = $ld.Size
                    SizeRemaining = $ld.FreeSpace
                }
            }
        }

        foreach ($v in $volumes) {
            $totalGb = [math]::Round($v.Size / 1GB, 2)
            $freeGb = [math]::Round($v.SizeRemaining / 1GB, 2)
            $percentFree = if ($v.Size -gt 0) { [math]::Round(($v.SizeRemaining / $v.Size) * 100, 1) } else { 0 }

            if ($percentFree -lt $FreeThresholdPercent) {
                $anyBelowThreshold = $true
            }

            # Rough write-speed sample: write ~50MB to a temp file on this volume,
            # time it, then delete the temp file immediately. Informational only.
            $speedResult = 'N/A'
            $driveRoot = "$($v.DriveLetter):\"
            if (Test-Path -Path $driveRoot) {
                $probeFile = Join-Path -Path $driveRoot -ChildPath ("itc_speedtest_{0}.tmp" -f ([guid]::NewGuid().ToString('N')))
                try {
                    $bytes = New-Object byte[] (50MB)
                    (New-Object System.Random).NextBytes($bytes)
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    [System.IO.File]::WriteAllBytes($probeFile, $bytes)
                    $sw.Stop()
                    $seconds = [math]::Max($sw.Elapsed.TotalSeconds, 0.001)
                    $mbps = [math]::Round(50 / $seconds, 1)
                    $speedResult = "~$mbps MB/s (approx.)"
                }
                catch {
                    $speedResult = "N/A (could not sample: $($_.Exception.Message))"
                }
                finally {
                    if (Test-Path -Path $probeFile) {
                        Remove-Item -Path $probeFile -Force -ErrorAction SilentlyContinue
                    }
                }
            }

            $volumeRows += [PSCustomObject]@{
                DriveLetter      = $v.DriveLetter
                'TotalGB'        = $totalGb
                'FreeGB'         = $freeGb
                'PercentFree'    = "$percentFree%"
                'WriteSpeed(approx)' = $speedResult
            }
        }

        $volumeTable = ConvertTo-MarkdownTable -InputObject $volumeRows -Columns @('DriveLetter', 'TotalGB', 'FreeGB', 'PercentFree', 'WriteSpeed(approx)') -EmptyMessage '_No fixed volumes found._'

        # Physical disk health via Get-PhysicalDisk, with graceful fallback.
        $diskHealthTable = '_Get-PhysicalDisk is not available on this system; physical disk health could not be determined this way._'
        $anyUnhealthy = $false
        try {
            if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
                $physicalDisks = Get-PhysicalDisk -ErrorAction Stop

                $diskRows = foreach ($d in $physicalDisks) {
                    $isHealthy = ($d.HealthStatus -in @('Healthy', 'OK'))
                    if (-not $isHealthy) { $anyUnhealthy = $true }

                    [PSCustomObject]@{
                        FriendlyName      = $d.FriendlyName
                        HealthStatus      = $d.HealthStatus
                        OperationalStatus = ($d.OperationalStatus -join ', ')
                        MediaType         = $d.MediaType
                    }
                }

                $diskHealthTable = ConvertTo-MarkdownTable -InputObject $diskRows -Columns @('FriendlyName', 'HealthStatus', 'OperationalStatus', 'MediaType') -EmptyMessage '_No physical disk data found._'
            }
        }
        catch {
            $diskHealthTable = New-SectionErrorNote -Reason $_.Exception.Message
        }

        $verdict = if ($anyBelowThreshold -or $anyUnhealthy) { 'FAIL' } else { 'PASS' }

        $notes = @()
        $notes += "Threshold: fail if any fixed volume has less than $FreeThresholdPercent% free space, or any physical disk HealthStatus is not Healthy/OK."
        $notes += ''
        $notes += '**Fixed Volumes:**'
        $notes += ''
        $notes += $volumeTable
        $notes += ''
        $notes += '**Physical Disk Health (Get-PhysicalDisk):**'
        $notes += ''
        $notes += $diskHealthTable
        $notes += ''
        $notes += '_Write-speed figures are a rough, uncertified, single-sample indicator only (writing ~50MB once per volume) -- they are not a certified storage benchmark. The probe file is deleted immediately after each measurement._'

        return New-CheckResult -Verdict $verdict -Content ($notes -join "`n")
    }
    catch {
        return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
    }
}

function Test-UsernameCompliance {
    <#
        Checklist item 2: Username follows the IT-defined naming convention. If no
        pattern was supplied via -ExpectedUsernamePattern, this check is SKIPPED --
        the script never guesses a company naming policy.
    #>
    [CmdletBinding()]
    param(
        [string]$Pattern
    )

    try {
        $currentUsername = $env:USERNAME

        if ([string]::IsNullOrWhiteSpace($Pattern)) {
            $content = @(
                "- **Current Username (INFO):** $currentUsername"
                ''
                '_No company naming-convention pattern was provided via -ExpectedUsernamePattern. This check is SKIPPED -- re-run with -ExpectedUsernamePattern once IT provides the naming rule. No pattern was guessed._'
            ) -join "`n"

            return New-CheckResult -Verdict 'SKIPPED' -Content $content
        }

        $isMatch = $currentUsername -match $Pattern
        $verdict = if ($isMatch) { 'PASS' } else { 'FAIL' }

        $content = @(
            "- **Current Username (INFO):** $currentUsername"
            "- **Expected Pattern:** ``$Pattern``"
            "- **Matches Pattern:** $isMatch"
        ) -join "`n"

        return New-CheckResult -Verdict $verdict -Content $content
    }
    catch {
        return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
    }
}

function Test-NetworkCompliance {
    <#
        Checklist item 3: Network belongs to the company. Reports domain join
        status, active adapters, current Wi-Fi SSID (if any), and default gateway.
        Compares against -ExpectedDomain / -ExpectedSSIDs when supplied.
    #>
    [CmdletBinding()]
    param(
        [string]$Domain,
        [string[]]$SSIDs
    )

    try {
        $infoLines = @()
        $domainVerdict = $null
        $ssidVerdict = $null

        # Domain join status.
        $isDomainJoined = $false
        $currentDomain = $null
        try {
            $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $isDomainJoined = [bool]$cs.PartOfDomain
            $currentDomain = if ($isDomainJoined) { $cs.Domain } else { $cs.Workgroup }
            $infoLines += "- **Domain Joined (INFO):** $isDomainJoined"
            $infoLines += "- **Domain/Workgroup Name (INFO):** $currentDomain"
        }
        catch {
            $infoLines += New-SectionErrorNote -Reason $_.Exception.Message
        }

        # Active network adapters.
        $adapterTable = '_No active network adapters found._'
        try {
            if (Get-Command -Name Get-NetAdapter -ErrorAction SilentlyContinue) {
                $adapters = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' }
                $rows = foreach ($a in $adapters) {
                    [PSCustomObject]@{
                        Name       = $a.Name
                        Status     = $a.Status
                        LinkSpeed  = $a.LinkSpeed
                        MacAddress = $a.MacAddress
                    }
                }
                $adapterTable = ConvertTo-MarkdownTable -InputObject $rows -Columns @('Name', 'Status', 'LinkSpeed', 'MacAddress') -EmptyMessage '_No active network adapters found._'
            }
        }
        catch {
            $adapterTable = New-SectionErrorNote -Reason $_.Exception.Message
        }

        # Current Wi-Fi SSID via netsh wlan show interfaces.
        $currentSSID = $null
        $onWifi = $false
        try {
            $wlanOutput = netsh wlan show interfaces 2>$null
            if ($LASTEXITCODE -eq 0 -and $wlanOutput) {
                foreach ($line in $wlanOutput) {
                    if ($line -match '^\s*SSID\s*:\s*(.+?)\s*$' -and $line -notmatch 'BSSID') {
                        $currentSSID = $Matches[1].Trim()
                        $onWifi = $true
                        break
                    }
                }
            }
        }
        catch {
            $currentSSID = $null
        }

        if ($onWifi -and $currentSSID) {
            $infoLines += "- **Current Wi-Fi SSID (INFO):** $currentSSID"
        }
        else {
            $infoLines += "- **Current Wi-Fi SSID (INFO):** _(not connected via Wi-Fi, or SSID undetectable)_"
        }

        # Default gateway.
        try {
            $gateways = Get-NetIPConfiguration -ErrorAction Stop |
                Where-Object { $_.IPv4DefaultGateway } |
                ForEach-Object { $_.IPv4DefaultGateway.NextHop }
            $gatewayText = if ($gateways) { ($gateways | Select-Object -Unique) -join ', ' } else { '_None found_' }
            $infoLines += "- **Default Gateway (INFO):** $gatewayText"
        }
        catch {
            $infoLines += "- **Default Gateway (INFO):** _Unable to determine: $($_.Exception.Message)_"
        }

        # Compare against expected domain.
        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            $domainMatches = ($isDomainJoined -and $currentDomain -and ($currentDomain -eq $Domain))
            $domainVerdict = if ($domainMatches) { 'PASS' } else { 'FAIL' }
            $infoLines += "- **Expected Domain:** $Domain"
            $infoLines += "- **Domain Check:** $domainVerdict"
        }

        # Compare against expected SSIDs.
        if ($SSIDs -and $SSIDs.Count -gt 0) {
            if ($onWifi -and $currentSSID) {
                $ssidMatches = $SSIDs -contains $currentSSID
                $ssidVerdict = if ($ssidMatches) { 'PASS' } else { 'FAIL' }
                $infoLines += "- **Expected SSID(s):** $($SSIDs -join ', ')"
                $infoLines += "- **SSID Check:** $ssidVerdict"
            }
            else {
                $ssidVerdict = 'N/A'
                $infoLines += "- **Expected SSID(s):** $($SSIDs -join ', ')"
                $infoLines += "- **SSID Check:** N/A (not currently connected via Wi-Fi)"
            }
        }

        $content = @()
        $content += ($infoLines -join "`n")
        $content += ''
        $content += '**Active Network Adapters:**'
        $content += ''
        $content += $adapterTable

        # Overall verdict for this check.
        $overallVerdict = $null
        if ($null -eq $domainVerdict -and $null -eq $ssidVerdict) {
            $overallVerdict = 'SKIPPED'
            $content += ''
            $content += '_Neither -ExpectedDomain nor -ExpectedSSIDs was supplied. This check is SKIPPED overall -- supply one or both parameters to enable a PASS/FAIL verdict. Informational network details above are still shown._'
        }
        else {
            $verdicts = @($domainVerdict, $ssidVerdict) | Where-Object { $_ -and $_ -ne 'N/A' }
            if ($verdicts -contains 'FAIL') {
                $overallVerdict = 'FAIL'
            }
            else {
                $overallVerdict = 'PASS'
            }
        }

        return New-CheckResult -Verdict $overallVerdict -Content ($content -join "`n")
    }
    catch {
        return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
    }
}

function Test-OsUpdateCompliance {
    <#
        Checklist item 4: OS is up to date. Reports current OS caption/version/build.
        Unless -Skip is set, searches for pending Windows updates inside a background
        job bounded by -TimeoutSeconds so a hung search can never hang the script.
    #>
    [CmdletBinding()]
    param(
        [switch]$Skip,
        [int]$TimeoutSeconds
    )

    try {
        $osInfoLines = @()
        try {
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
            $osInfoLines += "- **OS (INFO):** $($os.Caption)"
            $osInfoLines += "- **Version (INFO):** $($os.Version)"
            $osInfoLines += "- **Build (INFO):** $($os.BuildNumber)"
        }
        catch {
            $osInfoLines += New-SectionErrorNote -Reason $_.Exception.Message
        }

        if ($Skip) {
            $content = @()
            $content += ($osInfoLines -join "`n")
            $content += ''
            $content += '_-SkipWindowsUpdateCheck was specified. The pending-updates search was not performed._'
            return New-CheckResult -Verdict 'N/A' -Content ($content -join "`n")
        }

        $job = $null
        try {
            $job = Start-Job -ScriptBlock {
                try {
                    $updateSession = New-Object -ComObject Microsoft.Update.Session
                    $updateSearcher = $updateSession.CreateUpdateSearcher()
                    $searchResult = $updateSearcher.Search("IsInstalled=0 and IsHidden=0")

                    $titles = @()
                    foreach ($update in $searchResult.Updates) {
                        $titles += $update.Title
                    }

                    return [PSCustomObject]@{
                        Success = $true
                        Count   = $searchResult.Updates.Count
                        Titles  = $titles
                        Error   = $null
                    }
                }
                catch {
                    return [PSCustomObject]@{
                        Success = $false
                        Count   = 0
                        Titles  = @()
                        Error   = $_.Exception.Message
                    }
                }
            }

            $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds

            if (-not $completed -or $completed.State -eq 'Running') {
                $content = @()
                $content += ($osInfoLines -join "`n")
                $content += ''
                $content += "_Windows Update search timed out after $TimeoutSeconds second(s). Verdict is UNKNOWN -- the search may be slow on this network, or Windows Update may be unreachable/locked down. Consider re-running with a higher -WindowsUpdateTimeoutSeconds, or use -SkipWindowsUpdateCheck on known-offline machines._"
                return New-CheckResult -Verdict 'UNKNOWN' -Content ($content -join "`n")
            }

            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue

            if (-not $result -or -not $result.Success) {
                $reason = if ($result -and $result.Error) { $result.Error } else { 'Windows Update search returned no result.' }
                $content = @()
                $content += ($osInfoLines -join "`n")
                $content += ''
                $content += "_Windows Update search failed: $reason. Verdict is UNKNOWN._"
                return New-CheckResult -Verdict 'UNKNOWN' -Content ($content -join "`n")
            }

            $pendingCount = $result.Count
            $verdict = if ($pendingCount -eq 0) { 'PASS' } else { 'FAIL' }

            $content = @()
            $content += ($osInfoLines -join "`n")
            $content += "- **Pending Updates Found:** $pendingCount"
            $content += ''

            if ($pendingCount -gt 0) {
                $titleRows = $result.Titles | Select-Object -First 10 | ForEach-Object { [PSCustomObject]@{ Title = $_ } }
                $titleTable = ConvertTo-MarkdownTable -InputObject $titleRows -Columns @('Title')
                $content += "**Pending Update Titles (showing up to 10 of $pendingCount):**"
                $content += ''
                $content += $titleTable
            }
            else {
                $content += '_No pending updates found._'
            }

            return New-CheckResult -Verdict $verdict -Content ($content -join "`n")
        }
        finally {
            if ($job) {
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
    }
}

function Test-BatteryCompliance {
    <#
        Checklist item 5: Battery life is normal. Detects whether a battery is
        present; if so, runs "powercfg /batteryreport" to a short-lived temp XML
        file, parses DesignCapacity/FullChargeCapacity, and always deletes the temp
        file afterward (try/finally), even on error.
    #>
    [CmdletBinding()]
    param(
        [int]$HealthThresholdPercent
    )

    $tempXmlPath = $null

    try {
        $batteries = $null
        try {
            $batteries = Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop
        }
        catch {
            $batteries = $null
        }

        if (-not $batteries) {
            return New-CheckResult -Verdict 'N/A' -Content '_No battery detected -- this appears to be a desktop machine (or the battery could not be enumerated). Battery health check does not apply._'
        }

        $tempXmlPath = Join-Path -Path $env:TEMP -ChildPath ("itc_battery_{0}.xml" -f ([guid]::NewGuid().ToString('N')))

        try {
            $powercfgOutput = & powercfg /batteryreport /output $tempXmlPath /xml 2>&1
            $exitCode = $LASTEXITCODE

            if ($exitCode -ne 0 -or -not (Test-Path -Path $tempXmlPath)) {
                $reason = if ($powercfgOutput) { ($powercfgOutput | Out-String).Trim() } else { "powercfg exited with code $exitCode" }
                return New-CheckResult -Verdict 'UNKNOWN' -Content "_powercfg /batteryreport failed or produced no output file: $reason. This can happen if powercfg is blocked in this environment. Verdict is UNKNOWN._"
            }

            [xml]$reportXml = Get-Content -Path $tempXmlPath -Raw -ErrorAction Stop

            # The battery report XML schema nests capacities under
            # BatteryReport/Batteries/Battery, in mWh.
            $batteryNode = $reportXml.BatteryReport.Batteries.Battery
            if ($batteryNode -is [System.Array]) {
                $batteryNode = $batteryNode[0]
            }

            if (-not $batteryNode) {
                return New-CheckResult -Verdict 'UNKNOWN' -Content '_Battery report XML did not contain the expected Battery node. Verdict is UNKNOWN.'
            }

            $designCapacity = [double]$batteryNode.DesignCapacity
            $fullChargeCapacity = [double]$batteryNode.FullChargeCapacity

            if ($designCapacity -le 0) {
                return New-CheckResult -Verdict 'UNKNOWN' -Content '_Battery report XML reported a DesignCapacity of 0 or less; health percentage cannot be computed. Verdict is UNKNOWN._'
            }

            $healthPercent = [math]::Round(($fullChargeCapacity / $designCapacity) * 100, 1)
            $verdict = if ($healthPercent -ge $HealthThresholdPercent) { 'PASS' } else { 'FAIL' }

            $content = @()
            $content += "- **Design Capacity (INFO):** $designCapacity mWh"
            $content += "- **Full Charge Capacity (INFO):** $fullChargeCapacity mWh"
            $content += "- **Battery Health:** $healthPercent%"
            $content += "- **Threshold:** $HealthThresholdPercent%"

            return New-CheckResult -Verdict $verdict -Content ($content -join "`n")
        }
        catch {
            return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
        }
    }
    catch {
        return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
    }
    finally {
        if ($tempXmlPath -and (Test-Path -Path $tempXmlPath)) {
            Remove-Item -Path $tempXmlPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-SoftwareLicensingIndicator {
    <#
        Checklist item 6: Software is genuinely licensed. This script CANNOT
        verify third-party license authenticity -- that is not generically
        possible from a script. It only reports Windows OS activation status (and
        Office activation status if detectable) as an INFO-only partial indicator.
        This check never produces PASS/FAIL.
    #>
    [CmdletBinding()]
    param()

    try {
        $content = @()
        $content += '_This check is informational only (verdict is always INFO, never PASS/FAIL). It reports only Windows OS activation status and, if detectable, Microsoft Office activation status. This is a PARTIAL / INDIRECT indicator -- a full audit of every installed application''s license authenticity cannot be automated by this or any generic script and requires manual/procurement review against purchase records and vendor license terms._'
        $content += ''

        # Windows OS activation via SoftwareLicensingProduct.
        $windowsTable = '_Unable to determine Windows activation status._'
        try {
            $winProducts = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction Stop |
                Where-Object { $_.PartialProductKey -and $_.Name -match 'Windows' }

            $rows = foreach ($p in $winProducts) {
                $statusText = switch ($p.LicenseStatus) {
                    0 { 'Unlicensed' }
                    1 { 'Licensed' }
                    2 { 'OOBGrace' }
                    3 { 'OOTGrace' }
                    4 { 'NonGenuineGrace' }
                    5 { 'Notification' }
                    6 { 'ExtendedGrace' }
                    default { "Unknown ($($p.LicenseStatus))" }
                }
                [PSCustomObject]@{
                    Name          = $p.Name
                    LicenseStatus = $statusText
                }
            }

            $windowsTable = ConvertTo-MarkdownTable -InputObject $rows -Columns @('Name', 'LicenseStatus') -EmptyMessage '_No Windows licensing product entries found._'
        }
        catch {
            $windowsTable = New-SectionErrorNote -Reason $_.Exception.Message
        }

        $content += '**Windows OS Activation Status:**'
        $content += ''
        $content += $windowsTable
        $content += ''

        # Office activation status, best-effort, if detectable.
        $officeTable = '_Microsoft Office activation status was not detectable on this machine (Office may not be installed, or its licensing product entries are not exposed via SoftwareLicensingProduct)._'
        try {
            $officeProducts = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction Stop |
                Where-Object { $_.PartialProductKey -and $_.Name -match 'Office|Microsoft 365' }

            if ($officeProducts) {
                $rows = foreach ($p in $officeProducts) {
                    $statusText = switch ($p.LicenseStatus) {
                        0 { 'Unlicensed' }
                        1 { 'Licensed' }
                        2 { 'OOBGrace' }
                        3 { 'OOTGrace' }
                        4 { 'NonGenuineGrace' }
                        5 { 'Notification' }
                        6 { 'ExtendedGrace' }
                        default { "Unknown ($($p.LicenseStatus))" }
                    }
                    [PSCustomObject]@{
                        Name          = $p.Name
                        LicenseStatus = $statusText
                    }
                }
                $officeTable = ConvertTo-MarkdownTable -InputObject $rows -Columns @('Name', 'LicenseStatus')
            }
        }
        catch {
            $officeTable = New-SectionErrorNote -Reason $_.Exception.Message
        }

        $content += '**Microsoft Office Activation Status (best-effort):**'
        $content += ''
        $content += $officeTable

        return New-CheckResult -Verdict 'INFO' -Content ($content -join "`n")
    }
    catch {
        return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
    }
}

function Test-BloatwareDetection {
    <#
        Checklist item 7: Unused / bloatware programs. Enumerates installed
        applications (read-only registry query, same approach as
        Invoke-SystemAudit.ps1) and flags any whose DisplayName contains one of the
        -Keywords substrings, case-insensitively. DOES NOT UNINSTALL ANYTHING --
        detection/reporting only.
    #>
    [CmdletBinding()]
    param(
        [string[]]$Keywords
    )

    try {
        $apps = Get-InstalledApplicationsList

        $matched = @()
        if ($Keywords -and $Keywords.Count -gt 0) {
            foreach ($app in $apps) {
                foreach ($keyword in $Keywords) {
                    if ([string]::IsNullOrWhiteSpace($keyword)) { continue }
                    if ($app.DisplayName -and ($app.DisplayName.ToLowerInvariant().Contains($keyword.ToLowerInvariant()))) {
                        $matched += [PSCustomObject]@{
                            DisplayName    = $app.DisplayName
                            DisplayVersion = $app.DisplayVersion
                            Publisher      = $app.Publisher
                            MatchedKeyword = $keyword
                        }
                        break
                    }
                }
            }
        }

        $matchedCount = $matched.Count
        $table = ConvertTo-MarkdownTable -InputObject $matched -Columns @('DisplayName', 'DisplayVersion', 'Publisher', 'MatchedKeyword') -EmptyMessage '_No installed applications matched any bloatware keyword._'

        $content = @()
        $content += "**Candidate bloatware/low-value programs found:** $matchedCount (out of $($apps.Count) installed applications scanned)"
        $content += ''
        $content += "**Keywords used:** $($Keywords -join ', ')"
        $content += ''
        $content += $table
        $content += ''
        $content += '_This is a heuristic keyword match only, based on -BloatwareKeywords. It may include false positives and can miss unlisted bloatware. NO PROGRAM IS REMOVED BY THIS SCRIPT -- this list is for a human (IT) to review before deciding whether to uninstall anything._'

        return New-CheckResult -Verdict 'INFO' -Content ($content -join "`n")
    }
    catch {
        return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
    }
}

function Get-FolderSizeBytes {
    <#
        Read-only helper: recursively sums file sizes under a path without
        deleting or modifying anything. Returns 0 if the path doesn't exist or
        can't be read.
    #>
    [CmdletBinding()]
    param(
        [string]$Path
    )

    if (-not $Path -or -not (Test-Path -Path $Path)) {
        return 0
    }

    try {
        $sum = (Get-ChildItem -Path $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if ($null -eq $sum) { return 0 }
        return $sum
    }
    catch {
        return 0
    }
}

function Test-JunkFilesSizing {
    <#
        Checklist item 8: Junk files. Computes total size (does NOT delete
        anything) of common reclaimable locations: current user's %TEMP%,
        C:\Windows\Temp, the Recycle Bin, and the Windows Update download cache.
    #>
    [CmdletBinding()]
    param(
        [double]$WarnThresholdGB
    )

    try {
        $rows = @()

        $userTempBytes = Get-FolderSizeBytes -Path $env:TEMP
        $rows += [PSCustomObject]@{ Location = "User TEMP ($env:TEMP)"; SizeGB = [math]::Round($userTempBytes / 1GB, 3) }

        $windowsTempBytes = Get-FolderSizeBytes -Path 'C:\Windows\Temp'
        $rows += [PSCustomObject]@{ Location = 'C:\Windows\Temp'; SizeGB = [math]::Round($windowsTempBytes / 1GB, 3) }

        $swDistBytes = 0
        $swDistPath = 'C:\Windows\SoftwareDistribution\Download'
        if (Test-Path -Path $swDistPath) {
            $swDistBytes = Get-FolderSizeBytes -Path $swDistPath
            $rows += [PSCustomObject]@{ Location = $swDistPath; SizeGB = [math]::Round($swDistBytes / 1GB, 3) }
        }
        else {
            $rows += [PSCustomObject]@{ Location = $swDistPath; SizeGB = 'N/A (not present)' }
        }

        $recycleBinBytes = 0
        try {
            $shell = New-Object -ComObject Shell.Application
            $recycleBin = $shell.Namespace(10)
            if ($recycleBin) {
                # GetDetailsOf returns a formatted/localized size string which is
                # unreliable to parse; ExtendedProperty('Size') gives the raw byte
                # count directly and is the more robust property to sum.
                foreach ($item in $recycleBin.Items()) {
                    try {
                        $recycleBinBytes += [int64]$item.ExtendedProperty('Size')
                    }
                    catch {
                        # If ExtendedProperty isn't available, this item's size is skipped.
                    }
                }
            }
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell)
        }
        catch {
            # Recycle Bin sizing best-effort only; leave as 0 with a note below.
        }

        $rows += [PSCustomObject]@{ Location = 'Recycle Bin (all drives)'; SizeGB = [math]::Round($recycleBinBytes / 1GB, 3) }

        $totalBytes = $userTempBytes + $windowsTempBytes + $swDistBytes + $recycleBinBytes
        $totalGB = [math]::Round($totalBytes / 1GB, 3)

        $table = ConvertTo-MarkdownTable -InputObject $rows -Columns @('Location', 'SizeGB')

        $verdict = if ($totalGB -gt $WarnThresholdGB) { 'WARN' } else { 'INFO' }

        $content = @()
        $content += $table
        $content += ''
        $content += "**Total reclaimable (approx.):** $totalGB GB (warn threshold: $WarnThresholdGB GB)"
        $content += ''
        $content += "_NOTHING WAS DELETED. This is a sizing report only, for IT to review and decide on cleanup action. Recycle Bin size is best-effort via the Shell.Application COM object and may be approximate._"

        return New-CheckResult -Verdict $verdict -Content ($content -join "`n")
    }
    catch {
        return New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
    }
}

function Test-AntivirusCompliance {
    <#
        Checklist item 9: Antivirus enabled. Prefers Get-MpComputerStatus (Windows
        Defender). Also checks root/SecurityCenter2 AntiVirusProduct to catch a
        third-party AV that may have replaced/disabled Defender.
    #>
    [CmdletBinding()]
    param()

    $defenderOk = $false
    $defenderQueried = $false
    $thirdPartyOk = $false
    $thirdPartyQueried = $false
    $content = @()

    try {
        $defenderStatus = Get-MpComputerStatus -ErrorAction Stop
        $defenderQueried = $true

        $sigAgeDays = $null
        try { $sigAgeDays = $defenderStatus.AntivirusSignatureAge } catch { $sigAgeDays = $null }

        $defenderOk = [bool]$defenderStatus.RealTimeProtectionEnabled

        $content += '**Windows Defender (Get-MpComputerStatus):**'
        $content += ''
        $content += "- AntivirusEnabled: $($defenderStatus.AntivirusEnabled)"
        $content += "- RealTimeProtectionEnabled: $($defenderStatus.RealTimeProtectionEnabled)"
        $content += "- AntispywareEnabled: $($defenderStatus.AntispywareEnabled)"
        $content += "- Signature Age (days): $sigAgeDays"
        $content += ''
    }
    catch {
        $content += '**Windows Defender (Get-MpComputerStatus):**'
        $content += ''
        $content += (New-SectionErrorNote -Reason $_.Exception.Message)
        $content += ''
    }

    try {
        $avProducts = Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop
        $thirdPartyQueried = $true

        if ($avProducts) {
            $rows = foreach ($p in $avProducts) {
                # productState is a bitmask; a commonly used heuristic is that the
                # middle byte indicates enabled status (0x10 = enabled). This is a
                # best-effort, undocumented convention used widely by AV inventory
                # tools, not an official Microsoft API contract.
                $stateHex = '{0:X6}' -f [int]$p.productState
                $enabledHex = $stateHex.Substring(2, 2)
                $isEnabled = ($enabledHex -in @('10', '11'))
                if ($isEnabled) { $thirdPartyOk = $true }

                [PSCustomObject]@{
                    DisplayName  = $p.DisplayName
                    ProductState = $stateHex
                    LikelyEnabled = $isEnabled
                }
            }

            $content += '**Third-Party AV (root/SecurityCenter2 AntiVirusProduct):**'
            $content += ''
            $content += (ConvertTo-MarkdownTable -InputObject $rows -Columns @('DisplayName', 'ProductState', 'LikelyEnabled'))
            $content += ''
            $content += '_"LikelyEnabled" is derived from the undocumented productState bitmask convention used by Security Center inventory tools -- treat it as a best-effort indicator, not a guarantee._'
        }
        else {
            $content += '**Third-Party AV (root/SecurityCenter2 AntiVirusProduct):**'
            $content += ''
            $content += '_No third-party AV products registered with Security Center._'
        }
    }
    catch {
        $content += '**Third-Party AV (root/SecurityCenter2 AntiVirusProduct):**'
        $content += ''
        $content += (New-SectionErrorNote -Reason $_.Exception.Message)
        $content += '_This namespace is commonly unavailable on Windows Server SKUs._'
    }

    if (-not $defenderQueried -and -not $thirdPartyQueried) {
        return New-CheckResult -Verdict 'UNKNOWN' -Content ($content -join "`n")
    }

    $verdict = if ($defenderOk -or $thirdPartyOk) { 'PASS' } else { 'FAIL' }

    return New-CheckResult -Verdict $verdict -Content ($content -join "`n")
}

#endregion Check functions

#region Main

function Invoke-ITComplianceCheck {
    [CmdletBinding()]
    param(
        [string]$OutputPath,
        [string]$ExpectedUsernamePattern,
        [string]$ExpectedDomain,
        [string[]]$ExpectedSSIDs,
        [int]$StorageFreeThresholdPercent,
        [int]$BatteryHealthThresholdPercent,
        [double]$JunkFilesWarnThresholdGB,
        [string[]]$BloatwareKeywords,
        [switch]$SkipWindowsUpdateCheck,
        [int]$WindowsUpdateTimeoutSeconds
    )

    # Resolve / validate the output path up front.
    try {
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }
        $resolvedOutputPath = (Resolve-Path -Path $OutputPath).ProviderPath
    }
    catch {
        Write-Error "Could not resolve or create output path '$OutputPath': $($_.Exception.Message)"
        return
    }

    Write-Host "Starting IT compliance check..." -ForegroundColor Cyan

    # Ordered list of the 9 checklist items: key, Thai label, anchor/title, and a
    # scriptblock that runs the check and returns a New-CheckResult object.
    $checklist = [ordered]@{
        storage = @{
            Number    = 1
            ThaiLabel = 'Storage ใช้งานได้ปกติ/ไม่เต็ม'
            Title     = 'Storage'
            Anchor    = 'storage'
            Progress  = 'Checking storage (free space, disk health, write-speed sample)...'
            Run       = { Test-StorageCompliance -FreeThresholdPercent $StorageFreeThresholdPercent }
        }
        username = @{
            Number    = 2
            ThaiLabel = 'Username ตามที่ IT กำหนด'
            Title     = 'Username'
            Anchor    = 'username'
            Progress  = 'Checking username against naming convention...'
            Run       = { Test-UsernameCompliance -Pattern $ExpectedUsernamePattern }
        }
        network = @{
            Number    = 3
            ThaiLabel = 'Network ของบริษัท'
            Title     = 'Network'
            Anchor    = 'network'
            Progress  = 'Checking network (domain join, Wi-Fi SSID, adapters)...'
            Run       = { Test-NetworkCompliance -Domain $ExpectedDomain -SSIDs $ExpectedSSIDs }
        }
        osupdate = @{
            Number    = 4
            ThaiLabel = 'OS เวอร์ชั่นล่าสุด'
            Title     = 'OS Up To Date'
            Anchor    = 'os-up-to-date'
            Progress  = 'Checking OS update status (this may take a moment)...'
            Run       = { Test-OsUpdateCompliance -Skip:$SkipWindowsUpdateCheck -TimeoutSeconds $WindowsUpdateTimeoutSeconds }
        }
        battery = @{
            Number    = 5
            ThaiLabel = 'Battery Life ปกติ'
            Title     = 'Battery Life'
            Anchor    = 'battery-life'
            Progress  = 'Checking battery health...'
            Run       = { Test-BatteryCompliance -HealthThresholdPercent $BatteryHealthThresholdPercent }
        }
        licensing = @{
            Number    = 6
            ThaiLabel = 'Software ลิขสิทธิ์แท้ (ตรวจได้บางส่วน)'
            Title     = 'Software Licensing (Partial Indicator)'
            Anchor    = 'software-licensing-partial-indicator'
            Progress  = 'Checking Windows/Office activation status (partial indicator only)...'
            Run       = { Test-SoftwareLicensingIndicator }
        }
        bloatware = @{
            Number    = 7
            ThaiLabel = 'โปรแกรมที่ไม่ได้ใช้งาน/Debloat (ตรวจจับเท่านั้น)'
            Title     = 'Unused / Bloatware Programs (Detection Only)'
            Anchor    = 'unused--bloatware-programs-detection-only'
            Progress  = 'Scanning installed applications for bloatware keywords (detection only, nothing removed)...'
            Run       = { Test-BloatwareDetection -Keywords $BloatwareKeywords }
        }
        junkfiles = @{
            Number    = 8
            ThaiLabel = 'ไฟล์ขยะ (ตรวจจับเท่านั้น)'
            Title     = 'Junk Files (Detection Only)'
            Anchor    = 'junk-files-detection-only'
            Progress  = 'Sizing reclaimable junk file locations (detection only, nothing deleted)...'
            Run       = { Test-JunkFilesSizing -WarnThresholdGB $JunkFilesWarnThresholdGB }
        }
        antivirus = @{
            Number    = 9
            ThaiLabel = 'Antivirus เปิดใช้งาน'
            Title     = 'Antivirus Enabled'
            Anchor    = 'antivirus-enabled'
            Progress  = 'Checking antivirus / real-time protection status...'
            Run       = { Test-AntivirusCompliance }
        }
    }

    $results = [ordered]@{}

    foreach ($key in $checklist.Keys) {
        $item = $checklist[$key]
        Write-Host "  [$($item.Number)/9] $($item.Progress)" -ForegroundColor Yellow

        try {
            $result = & $item.Run
        }
        catch {
            # Belt-and-suspenders: even if a check function somehow throws instead
            # of catching internally, never let it abort the rest of the script.
            $result = New-CheckResult -Verdict 'UNKNOWN' -Content (New-SectionErrorNote -Reason $_.Exception.Message)
        }

        $results[$key] = @{
            Number    = $item.Number
            ThaiLabel = $item.ThaiLabel
            Title     = $item.Title
            Anchor    = $item.Anchor
            Verdict   = $result.Verdict
            Content   = $result.Content
        }
    }

    Write-Host "  Assembling Markdown report..." -ForegroundColor Yellow

    # Overall verdict: FAIL if any FAIL, else WARN if any WARN, else PASS.
    $allVerdicts = $results.Values | ForEach-Object { $_.Verdict }
    $overall = if ($allVerdicts -contains 'FAIL') {
        'FAIL'
    }
    elseif ($allVerdicts -contains 'WARN') {
        'WARN'
    }
    else {
        'PASS'
    }

    $reportBuilder = New-Object System.Text.StringBuilder

    [void]$reportBuilder.AppendLine('# IT Compliance Report')
    [void]$reportBuilder.AppendLine()
    [void]$reportBuilder.AppendLine("Generated by Test-ITCompliance.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$reportBuilder.AppendLine()
    [void]$reportBuilder.AppendLine("- **Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$reportBuilder.AppendLine("- **Computer Name:** $env:COMPUTERNAME")
    [void]$reportBuilder.AppendLine("- **Current User:** $env:USERNAME")
    [void]$reportBuilder.AppendLine()
    [void]$reportBuilder.AppendLine('> **This is a CHECK/AUDIT report only.** This script never deletes, uninstalls, moves, or modifies any file, program, or setting on this machine. Items 7 and 8 below only detect and report candidates/sizes for the IT team to review -- no removal action is taken by this script.')
    [void]$reportBuilder.AppendLine()

    [void]$reportBuilder.AppendLine('## Summary')
    [void]$reportBuilder.AppendLine()
    [void]$reportBuilder.AppendLine('| # | Requirement (Thai) | Verdict |')
    [void]$reportBuilder.AppendLine('| --- | --- | --- |')
    foreach ($key in $results.Keys) {
        $r = $results[$key]
        [void]$reportBuilder.AppendLine("| $($r.Number) | $($r.ThaiLabel) | $($r.Verdict) |")
    }
    [void]$reportBuilder.AppendLine()

    [void]$reportBuilder.AppendLine('## Table of Contents')
    [void]$reportBuilder.AppendLine()
    foreach ($key in $results.Keys) {
        $r = $results[$key]
        [void]$reportBuilder.AppendLine("- [$($r.Number). $($r.Title)](#$($r.Anchor))")
    }
    [void]$reportBuilder.AppendLine()

    foreach ($key in $results.Keys) {
        $r = $results[$key]
        [void]$reportBuilder.AppendLine("## $($r.Number). $($r.Title)")
        [void]$reportBuilder.AppendLine()
        [void]$reportBuilder.AppendLine("**Verdict:** $($r.Verdict)")
        [void]$reportBuilder.AppendLine()
        [void]$reportBuilder.AppendLine($r.Content)
        [void]$reportBuilder.AppendLine()
    }

    [void]$reportBuilder.AppendLine('## Overall')
    [void]$reportBuilder.AppendLine()
    [void]$reportBuilder.AppendLine("**Overall:** $overall")
    [void]$reportBuilder.AppendLine()
    [void]$reportBuilder.AppendLine('_Overall is FAIL if any item is FAIL; otherwise WARN if any item is WARN; otherwise PASS._')

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $hostName = $env:COMPUTERNAME
    $fileName = "ITComplianceReport_${hostName}_$timestamp.md"
    $fullPath = Join-Path -Path $resolvedOutputPath -ChildPath $fileName

    try {
        Set-Content -Path $fullPath -Value $reportBuilder.ToString() -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to save report to '$fullPath': $($_.Exception.Message)"
        return [PSCustomObject]@{ Saved = $false; Path = $null; Overall = $overall }
    }

    Write-Host ""
    Write-Host "IT compliance check complete. Report saved to:" -ForegroundColor Green
    Write-Host $fullPath -ForegroundColor Green
    Write-Host ""
    Write-Host "Overall verdict: $overall" -ForegroundColor $(if ($overall -eq 'FAIL') { 'Red' } elseif ($overall -eq 'WARN') { 'Yellow' } else { 'Green' })

    return [PSCustomObject]@{ Saved = $true; Path = $fullPath; Overall = $overall }
}

$runResult = Invoke-ITComplianceCheck `
    -OutputPath $OutputPath `
    -ExpectedUsernamePattern $ExpectedUsernamePattern `
    -ExpectedDomain $ExpectedDomain `
    -ExpectedSSIDs $ExpectedSSIDs `
    -StorageFreeThresholdPercent $StorageFreeThresholdPercent `
    -BatteryHealthThresholdPercent $BatteryHealthThresholdPercent `
    -JunkFilesWarnThresholdGB $JunkFilesWarnThresholdGB `
    -BloatwareKeywords $BloatwareKeywords `
    -SkipWindowsUpdateCheck:$SkipWindowsUpdateCheck `
    -WindowsUpdateTimeoutSeconds $WindowsUpdateTimeoutSeconds

# Exit code reflects overall verdict: 1 = FAIL, 0 = otherwise. This runs only
# after the report has already been saved above, so the report is never lost
# even if the exit code itself signals failure.
if ($runResult -and $runResult.Saved -and $runResult.Overall -eq 'FAIL') {
    exit 1
}
elseif ($runResult -and $runResult.Saved) {
    exit 0
}
else {
    # Report could not be saved at all -- surface as a failure exit code too.
    exit 1
}

#endregion Main
