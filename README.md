# SystemAuditReport

A single, self-contained Windows PowerShell script that snapshots a machine's
running processes, installed applications, and hardware, then writes it all
out as one Markdown report.

Repo: https://github.com/PECTECTH/PECScript

## Quick Start — run with a single command (no clone required)

Open **PowerShell** and run:

```powershell
irm https://raw.githubusercontent.com/PECTECTH/PECScript/main/Invoke-SystemAudit.ps1 | iex
```

- `irm` (`Invoke-RestMethod`) downloads the script's raw text from GitHub.
- `iex` (`Invoke-Expression`) runs it in the current session with all
  defaults — the report is saved straight to your **Documents** folder.
- If the repo's default branch is not `main` (e.g. `master`), swap it into
  the URL above.

**Security note:** this one-liner downloads and executes code directly from
GitHub. That's convenient, but you should only run it against a URL/repo you
trust — the same caution as any `irm | iex` or `curl | bash` install
command. You can always inspect the script first by opening the raw URL in
a browser, or fetch it without running it:

```powershell
irm https://raw.githubusercontent.com/PECTECTH/PECScript/main/Invoke-SystemAudit.ps1 -OutFile Invoke-SystemAudit.ps1
```

### Run with custom parameters via the one-liner

`iex` on its own doesn't accept `-Parameter` values, so build a scriptblock
first if you need `-OutputPath` or `-TopProcessCount`:

```powershell
$script = irm https://raw.githubusercontent.com/PECTECTH/PECScript/main/Invoke-SystemAudit.ps1
& ([scriptblock]::Create($script)) -OutputPath 'C:\Reports' -TopProcessCount 25
```

## What it collects

- **Report Header** — generated timestamp, computer name, current user, OS
  caption/version/build, architecture, last boot time, and uptime.
- **Running Processes** — every running process (PID, name, CPU seconds,
  working-set memory in MB, start time), sorted by memory usage, with a
  total process count.
- **Installed Applications** — every application registered in the Windows
  uninstall registry keys (name, version, publisher, install date), with a
  total count. Uses the registry directly rather than `Win32_Product`, which
  is slow and can trigger MSI self-repair.
- **CPU** — name, physical cores, logical processors, max clock speed,
  manufacturer.
- **Memory (RAM)** — total physical memory plus a per-module breakdown
  (capacity, speed, manufacturer, part number).
- **Disks** — physical disks (model, size, interface/media type) and logical
  volumes (drive letter, total size, free space, % used).
- **Network Adapters / MAC Addresses** — adapter name, MAC address, link
  speed, and status for every adapter that has a MAC address.
- **System / Hardware Summary** — manufacturer, model, and BIOS serial
  number.

Every section is gathered independently inside its own `try/catch`. If one
section fails (for example, a permissions issue), the rest of the report is
still produced — the failed section is clearly marked with the reason
instead of crashing the whole script.

## Requirements

- Windows 10/11 (or Windows Server) with Windows PowerShell 5.1+, or
  PowerShell 7+.
- No Administrator privileges required for the core functionality. A few
  data points (e.g. some BIOS/hardware details) may be incomplete when run
  as a standard user; the affected line will note that elevation may be
  required instead of failing.

## Usage (after cloning the repo locally)

Run with all defaults — saves the report to your Documents folder:

```powershell
.\Invoke-SystemAudit.ps1
```

Save to a specific folder, and only list the top 25 processes by memory
usage:

```powershell
.\Invoke-SystemAudit.ps1 -OutputPath 'C:\Reports' -TopProcessCount 25
```

If script execution is blocked by your local execution policy, run it with:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-SystemAudit.ps1
```

## Output

By default, the report is written to the current user's real **Documents**
folder (`[Environment]::GetFolderPath('MyDocuments')`) as:

```
SystemAuditReport_<HOSTNAME>_<yyyyMMdd_HHmmss>.md
```

The hostname is included so reports from multiple machines can sit in the
same Documents/folder (or be collected centrally) without overwriting each
other or needing to be renamed by hand.

The full path of the saved report is printed to the console when the script
finishes.

## Sample output (structure only, no real data)

```
# System Audit Report

## Table of Contents
- [Report Header](#report-header)
- [Running Processes](#running-processes)
- [Installed Applications](#installed-applications)
- [CPU](#cpu)
- [Memory (RAM)](#memory-ram)
- [Disks](#disks)
- [Network Adapters / MAC Addresses](#network-adapters--mac-addresses)
- [System / Hardware Summary](#system--hardware-summary)

## Running Processes
| Id | ProcessName | CPU(s) | WorkingSetMB | StartTime |
| --- | --- | --- | --- | --- |

## Installed Applications
| DisplayName | DisplayVersion | Publisher | InstallDate |
| --- | --- | --- | --- |

## CPU
| Name | NumberOfCores | NumberOfLogicalProcessors | MaxClockSpeed(MHz) | Manufacturer |
| --- | --- | --- | --- | --- |
```

## Privacy

This script only reads information that is already local to the machine it
runs on. It does not transmit any data over the network, call any external
service, or upload anything anywhere. The generated report itself (which
does contain real, machine-specific data such as installed software and MAC
addresses) is excluded from version control by `.gitignore` so it is never
accidentally committed.
