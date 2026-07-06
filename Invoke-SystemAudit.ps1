<#
.SYNOPSIS
    Generates a Markdown system audit report of the local Windows machine.

.DESCRIPTION
    Invoke-SystemAudit collects a snapshot of the local computer's configuration and
    state -- OS info, running processes, installed applications, CPU, memory, disks,
    network adapters, and general hardware/system info -- and writes it all out as a
    single, human-readable Markdown report.

    Every section is gathered independently and wrapped in its own try/catch, so a
    failure in one section (for example, a permissions issue) will not prevent the
    rest of the report from being produced. Sections that could not be completed are
    clearly marked in the output along with the reason.

    The script only reads information that is already local to this machine. It does
    not transmit any data over the network and does not require Administrator
    privileges to run its core functionality. A few data points (e.g. some BIOS or
    hardware details) may be incomplete when run as a standard user; when that
    happens the affected section notes that elevation may be required instead of
    failing.

.PARAMETER OutputPath
    Directory in which the generated Markdown report will be saved. Defaults to the
    current user's Documents folder.

.PARAMETER TopProcessCount
    Number of top processes (by working set memory) to include in the Running
    Processes section. Default is 0, which means "show all processes".

.EXAMPLE
    .\Invoke-SystemAudit.ps1

    Runs the audit with all defaults and saves the report to the current user's
    Documents folder.

.EXAMPLE
    .\Invoke-SystemAudit.ps1 -OutputPath 'C:\Reports' -TopProcessCount 25

    Runs the audit, saving the report to C:\Reports, and lists only the top 25
    processes by memory usage.

.NOTES
    Target: Windows PowerShell 5.1+ / PowerShell 7+
    Does not require Administrator for its core function.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = [Environment]::GetFolderPath('MyDocuments'),

    [Parameter()]
    [int]$TopProcessCount = 0
)

#region Helper functions

function New-SectionErrorNote {
    <#
        Builds the standard "could not retrieve this section" Markdown note.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Reason
    )

    return "_Unable to retrieve this section: $Reason (may require Administrator)_"
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
            if ($null -eq $value -or $value -eq '') {
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

#endregion Helper functions

#region Section gathering functions

function Get-ReportHeaderSection {
    [CmdletBinding()]
    param()

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

        $lastBoot = $os.LastBootUpTime
        $uptime = (Get-Date) - $lastBoot
        $uptimeString = '{0}d {1}h {2}m {3}s' -f $uptime.Days, $uptime.Hours, $uptime.Minutes, $uptime.Seconds

        $architecture = $os.OSArchitecture

        $lines = @()
        $lines += "- **Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $lines += "- **Computer Name:** $env:COMPUTERNAME"
        $lines += "- **Current User:** $env:USERNAME"
        $lines += "- **OS:** $($os.Caption) (Version $($os.Version), Build $($os.BuildNumber))"
        $lines += "- **Architecture:** $architecture"
        $lines += "- **Last Boot:** $($lastBoot.ToString('yyyy-MM-dd HH:mm:ss'))"
        $lines += "- **Uptime:** $uptimeString"

        return ($lines -join "`n")
    }
    catch {
        return New-SectionErrorNote -Reason $_.Exception.Message
    }
}

function Get-ProcessesSection {
    [CmdletBinding()]
    param(
        [int]$TopCount = 0
    )

    try {
        $allProcesses = Get-Process -ErrorAction Stop
        $totalCount = $allProcesses.Count

        $sorted = $allProcesses | Sort-Object -Property WorkingSet64 -Descending

        if ($TopCount -gt 0) {
            $sorted = $sorted | Select-Object -First $TopCount
        }

        $rows = foreach ($p in $sorted) {
            $cpu = $null
            try { $cpu = $p.CPU } catch { $cpu = $null }

            $startTime = $null
            try { $startTime = $p.StartTime } catch { $startTime = $null }

            [PSCustomObject]@{
                Id             = $p.Id
                ProcessName    = $p.ProcessName
                'CPU(s)'       = if ($null -ne $cpu) { [math]::Round($cpu, 2) } else { 'N/A' }
                'WorkingSetMB' = [math]::Round($p.WorkingSet64 / 1MB, 2)
                StartTime      = if ($startTime) { $startTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
            }
        }

        $table = ConvertTo-MarkdownTable -InputObject $rows -Columns @('Id', 'ProcessName', 'CPU(s)', 'WorkingSetMB', 'StartTime')

        $countLine = if ($TopCount -gt 0 -and $TopCount -lt $totalCount) {
            "**Total processes running:** $totalCount (showing top $TopCount by memory usage)"
        }
        else {
            "**Total processes running:** $totalCount"
        }

        return "$countLine`n`n$table"
    }
    catch {
        return New-SectionErrorNote -Reason $_.Exception.Message
    }
}

function Get-InstalledApplicationsSection {
    [CmdletBinding()]
    param()

    try {
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
                # Skip a registry hive path that isn't readable; other paths may
                # still succeed.
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

        $totalCount = $apps.Count

        $table = ConvertTo-MarkdownTable -InputObject $apps -Columns @('DisplayName', 'DisplayVersion', 'Publisher', 'InstallDate')

        return "**Total installed applications found:** $totalCount`n`n$table"
    }
    catch {
        return New-SectionErrorNote -Reason $_.Exception.Message
    }
}

function Get-CpuSection {
    [CmdletBinding()]
    param()

    try {
        $cpus = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop

        $rows = foreach ($cpu in $cpus) {
            [PSCustomObject]@{
                Name                        = $cpu.Name
                NumberOfCores               = $cpu.NumberOfCores
                NumberOfLogicalProcessors   = $cpu.NumberOfLogicalProcessors
                'MaxClockSpeed(MHz)'        = $cpu.MaxClockSpeed
                Manufacturer                = $cpu.Manufacturer
            }
        }

        return ConvertTo-MarkdownTable -InputObject $rows -Columns @('Name', 'NumberOfCores', 'NumberOfLogicalProcessors', 'MaxClockSpeed(MHz)', 'Manufacturer')
    }
    catch {
        return New-SectionErrorNote -Reason $_.Exception.Message
    }
}

function Get-MemorySection {
    [CmdletBinding()]
    param()

    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $totalGb = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)

        $summaryLine = "**Total Physical Memory:** $totalGb GB"

        $moduleTable = '_No per-module memory data found._'
        try {
            $modules = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction Stop

            $rows = foreach ($m in $modules) {
                [PSCustomObject]@{
                    'CapacityGB'  = [math]::Round($m.Capacity / 1GB, 2)
                    'Speed(MHz)'  = $m.Speed
                    Manufacturer  = $m.Manufacturer
                    PartNumber    = if ($m.PartNumber) { $m.PartNumber.Trim() } else { $null }
                }
            }

            $moduleTable = ConvertTo-MarkdownTable -InputObject $rows -Columns @('CapacityGB', 'Speed(MHz)', 'Manufacturer', 'PartNumber')
        }
        catch {
            $moduleTable = New-SectionErrorNote -Reason $_.Exception.Message
        }

        return "$summaryLine`n`n**Memory Modules:**`n`n$moduleTable"
    }
    catch {
        return New-SectionErrorNote -Reason $_.Exception.Message
    }
}

function Get-DisksSection {
    [CmdletBinding()]
    param()

    $physicalTable = '_No physical disk data found._'
    try {
        $physicalDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop

        $rows = foreach ($d in $physicalDisks) {
            [PSCustomObject]@{
                Model          = $d.Model
                'SizeGB'       = if ($d.Size) { [math]::Round($d.Size / 1GB, 2) } else { $null }
                InterfaceType  = $d.InterfaceType
                MediaType      = $d.MediaType
            }
        }

        $physicalTable = ConvertTo-MarkdownTable -InputObject $rows -Columns @('Model', 'SizeGB', 'InterfaceType', 'MediaType')
    }
    catch {
        $physicalTable = New-SectionErrorNote -Reason $_.Exception.Message
    }

    $logicalTable = '_No logical volume data found._'
    try {
        $volumes = Get-Volume -ErrorAction Stop | Where-Object { $_.DriveLetter }

        $rows = foreach ($v in $volumes) {
            $totalGb = [math]::Round($v.Size / 1GB, 2)
            $freeGb = [math]::Round($v.SizeRemaining / 1GB, 2)
            $usedPercent = if ($v.Size -gt 0) {
                [math]::Round((($v.Size - $v.SizeRemaining) / $v.Size) * 100, 1)
            }
            else {
                0
            }

            [PSCustomObject]@{
                DriveLetter  = $v.DriveLetter
                'TotalGB'    = $totalGb
                'FreeGB'     = $freeGb
                'PercentUsed' = "$usedPercent%"
            }
        }

        $logicalTable = ConvertTo-MarkdownTable -InputObject $rows -Columns @('DriveLetter', 'TotalGB', 'FreeGB', 'PercentUsed')
    }
    catch {
        # Fallback to Win32_LogicalDisk if Get-Volume isn't available (older systems
        # or restricted environments).
        try {
            $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop

            $rows = foreach ($ld in $logicalDisks) {
                $totalGb = [math]::Round($ld.Size / 1GB, 2)
                $freeGb = [math]::Round($ld.FreeSpace / 1GB, 2)
                $usedPercent = if ($ld.Size -gt 0) {
                    [math]::Round((($ld.Size - $ld.FreeSpace) / $ld.Size) * 100, 1)
                }
                else {
                    0
                }

                [PSCustomObject]@{
                    DriveLetter   = $ld.DeviceID
                    'TotalGB'     = $totalGb
                    'FreeGB'      = $freeGb
                    'PercentUsed' = "$usedPercent%"
                }
            }

            $logicalTable = ConvertTo-MarkdownTable -InputObject $rows -Columns @('DriveLetter', 'TotalGB', 'FreeGB', 'PercentUsed')
        }
        catch {
            $logicalTable = New-SectionErrorNote -Reason $_.Exception.Message
        }
    }

    return "**Physical Disks:**`n`n$physicalTable`n`n**Logical Volumes:**`n`n$logicalTable"
}

function Get-NetworkAdaptersSection {
    [CmdletBinding()]
    param()

    try {
        $rows = @()

        if (Get-Command -Name Get-NetAdapter -ErrorAction SilentlyContinue) {
            $adapters = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.MacAddress -and $_.MacAddress.Trim() -ne '' }

            $rows = foreach ($a in $adapters) {
                [PSCustomObject]@{
                    Name        = $a.Name
                    MacAddress  = $a.MacAddress
                    'LinkSpeed' = $a.LinkSpeed
                    Status      = $a.Status
                }
            }
        }
        else {
            # Fallback for older systems without the NetAdapter cmdlets.
            $configs = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ErrorAction Stop |
                Where-Object { $_.MACAddress -and $_.MACAddress.Trim() -ne '' }

            $rows = foreach ($c in $configs) {
                [PSCustomObject]@{
                    Name        = $c.Description
                    MacAddress  = $c.MACAddress
                    'LinkSpeed' = 'N/A'
                    Status      = if ($c.IPEnabled) { 'Up' } else { 'Down' }
                }
            }
        }

        return ConvertTo-MarkdownTable -InputObject $rows -Columns @('Name', 'MacAddress', 'LinkSpeed', 'Status') -EmptyMessage '_No network adapters with a MAC address were found._'
    }
    catch {
        return New-SectionErrorNote -Reason $_.Exception.Message
    }
}

function Get-SystemSummarySection {
    [CmdletBinding()]
    param()

    $lines = @()

    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $lines += "- **Manufacturer:** $($cs.Manufacturer)"
        $lines += "- **Model:** $($cs.Model)"
    }
    catch {
        $lines += New-SectionErrorNote -Reason $_.Exception.Message
    }

    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        $serial = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { '' }
        if ([string]::IsNullOrWhiteSpace($serial)) {
            $lines += "- **BIOS Serial Number:** _(blank - common on virtual machines)_"
        }
        else {
            $lines += "- **BIOS Serial Number:** $serial"
        }
    }
    catch {
        $lines += New-SectionErrorNote -Reason $_.Exception.Message
    }

    return ($lines -join "`n")
}

#endregion Section gathering functions

#region Main

function Invoke-SystemAudit {
    [CmdletBinding()]
    param(
        [string]$OutputPath,
        [int]$TopProcessCount
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

    Write-Host "Starting system audit..." -ForegroundColor Cyan

    $sections = [ordered]@{}

    Write-Host "  Gathering report header / OS info..." -ForegroundColor Yellow
    $sections['header'] = @{
        Title   = 'Report Header'
        Anchor  = 'report-header'
        Content = Get-ReportHeaderSection
    }

    Write-Host "  Gathering running processes..." -ForegroundColor Yellow
    $sections['processes'] = @{
        Title   = 'Running Processes'
        Anchor  = 'running-processes'
        Content = Get-ProcessesSection -TopCount $TopProcessCount
    }

    Write-Host "  Gathering installed applications (this may take a moment)..." -ForegroundColor Yellow
    $sections['apps'] = @{
        Title   = 'Installed Applications'
        Anchor  = 'installed-applications'
        Content = Get-InstalledApplicationsSection
    }

    Write-Host "  Gathering CPU info..." -ForegroundColor Yellow
    $sections['cpu'] = @{
        Title   = 'CPU'
        Anchor  = 'cpu'
        Content = Get-CpuSection
    }

    Write-Host "  Gathering memory (RAM) info..." -ForegroundColor Yellow
    $sections['memory'] = @{
        Title   = 'Memory (RAM)'
        Anchor  = 'memory-ram'
        Content = Get-MemorySection
    }

    Write-Host "  Gathering disk info..." -ForegroundColor Yellow
    $sections['disks'] = @{
        Title   = 'Disks'
        Anchor  = 'disks'
        Content = Get-DisksSection
    }

    Write-Host "  Gathering network adapter info..." -ForegroundColor Yellow
    $sections['network'] = @{
        Title   = 'Network Adapters / MAC Addresses'
        Anchor  = 'network-adapters--mac-addresses'
        Content = Get-NetworkAdaptersSection
    }

    Write-Host "  Gathering system/hardware summary..." -ForegroundColor Yellow
    $sections['summary'] = @{
        Title   = 'System / Hardware Summary'
        Anchor  = 'system--hardware-summary'
        Content = Get-SystemSummarySection
    }

    Write-Host "  Assembling Markdown report..." -ForegroundColor Yellow

    $reportBuilder = New-Object System.Text.StringBuilder

    [void]$reportBuilder.AppendLine('# System Audit Report')
    [void]$reportBuilder.AppendLine()
    [void]$reportBuilder.AppendLine("Generated by Invoke-SystemAudit.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$reportBuilder.AppendLine()
    [void]$reportBuilder.AppendLine('## Table of Contents')
    [void]$reportBuilder.AppendLine()

    foreach ($key in $sections.Keys) {
        $section = $sections[$key]
        [void]$reportBuilder.AppendLine("- [$($section.Title)](#$($section.Anchor))")
    }

    [void]$reportBuilder.AppendLine()

    foreach ($key in $sections.Keys) {
        $section = $sections[$key]
        [void]$reportBuilder.AppendLine("## $($section.Title)")
        [void]$reportBuilder.AppendLine()
        [void]$reportBuilder.AppendLine($section.Content)
        [void]$reportBuilder.AppendLine()
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $hostName = $env:COMPUTERNAME
    $fileName = "SystemAuditReport_${hostName}_$timestamp.md"
    $fullPath = Join-Path -Path $resolvedOutputPath -ChildPath $fileName

    try {
        Set-Content -Path $fullPath -Value $reportBuilder.ToString() -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to save report to '$fullPath': $($_.Exception.Message)"
        return
    }

    Write-Host ""
    Write-Host "System audit complete. Report saved to:" -ForegroundColor Green
    Write-Host $fullPath -ForegroundColor Green
}

Invoke-SystemAudit -OutputPath $OutputPath -TopProcessCount $TopProcessCount

#endregion Main
