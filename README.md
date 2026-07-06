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

## IT Compliance Checklist (Test-ITCompliance.ps1)

A second, separate script — `Test-ITCompliance.ps1` — checks the machine
against a company IT compliance checklist (9 numbered items) and writes its
own Markdown report. It is a **CHECK / AUDIT script only**.

> **This script never deletes or uninstalls anything.** It only reads,
> detects, and measures. Items **7** (unused/bloatware programs) and **8**
> (junk files) only **detect and report candidates/sizes** — no file,
> program, or setting is ever removed, moved, or changed by this script.
> Acting on those lists (or not) is entirely the IT team's decision.
>
> Item **6** (software licensing) is only a **partial / indirect
> indicator** (Windows/Office activation status) — it is not, and cannot
> be, a full audit of every installed application's license authenticity.

### What it checks

Each item maps directly to the company's numbered IT checklist and produces
exactly one primary verdict plus supporting details/tables:

1. **Storage ใช้งานได้ปกติ/ไม่เต็ม** — free space % on every fixed volume,
   physical disk health (`Get-PhysicalDisk`, with graceful fallback), and one
   rough, approximate write-speed sample per volume (a ~50MB temp file is
   written, timed, then deleted immediately). FAIL if any volume is below
   `-StorageFreeThresholdPercent` free, or any disk is not Healthy/OK.
2. **Username ตามที่ IT กำหนด** — compares `$env:USERNAME` against
   `-ExpectedUsernamePattern` if supplied (PASS/FAIL); otherwise **SKIPPED**
   with a note to re-run once IT provides the naming rule. The current
   username is always shown as INFO.
3. **Network ของบริษัท** — domain join status/name, active adapters, current
   Wi-Fi SSID (via `netsh wlan show interfaces`), and default gateway.
   Compares against `-ExpectedDomain` and/or `-ExpectedSSIDs` if supplied
   (PASS/FAIL, N/A for SSID if wired-only); **SKIPPED** overall (informational
   details still shown) if neither parameter is supplied.
4. **OS เวอร์ชั่นล่าสุด** — current OS caption/version/build, plus a pending
   Windows Update count (up to 10 titles listed) via the
   `Microsoft.Update.Session` COM object. The search always runs inside a
   background job bounded by `-WindowsUpdateTimeoutSeconds` so a slow/hung
   search can never hang the script (verdict UNKNOWN on timeout/error). "N/A"
   if `-SkipWindowsUpdateCheck` is used.
5. **Battery Life ปกติ** — detects a battery via `Win32_Battery`; "N/A" on
   desktops. If present, runs `powercfg /batteryreport` to a short-lived
   temp XML file, computes health % (`FullChargeCapacity / DesignCapacity`),
   and always deletes the temp file afterward (even on error). PASS/FAIL
   against `-BatteryHealthThresholdPercent`.
6. **Software ลิขสิทธิ์แท้ (ตรวจได้บางส่วน)** — Windows OS activation status
   (`SoftwareLicensingProduct`) and, if detectable, Microsoft Office
   activation status. Always **INFO** — never PASS/FAIL — because verifying
   third-party license authenticity cannot be automated generically.
7. **โปรแกรมที่ไม่ได้ใช้งาน/Debloat (ตรวจจับเท่านั้น)** — reuses the same
   uninstall-registry enumeration as `Invoke-SystemAudit.ps1`'s installed
   applications section, then flags DisplayNames containing any
   `-BloatwareKeywords` substring (case-insensitive). Always **INFO**, with a
   count and table of matches. Detection only — nothing is uninstalled.
8. **ไฟล์ขยะ (ตรวจจับเท่านั้น)** — total size of common reclaimable locations
   (`%TEMP%`, `C:\Windows\Temp`, the Recycle Bin, and
   `C:\Windows\SoftwareDistribution\Download`). **WARN** if the total exceeds
   `-JunkFilesWarnThresholdGB`, otherwise **INFO**. Sizing only — nothing is
   deleted.
9. **Antivirus เปิดใช้งาน** — prefers `Get-MpComputerStatus` (Windows
   Defender real-time protection, signature age) and also checks
   `root/SecurityCenter2 AntiVirusProduct` to catch a third-party AV that may
   have replaced Defender. PASS if either reports active protection; FAIL if
   neither does; UNKNOWN if both queries error (common on Server SKUs).

Every check is wrapped in its own function with a `try/catch`, mirroring the
defensive style in `Invoke-SystemAudit.ps1` — a failure in one check reports
verdict **UNKNOWN** with the reason instead of aborting the rest of the
script.

### Verdict scale

`PASS`, `FAIL`, `WARN`, `INFO`, `SKIPPED`, `UNKNOWN`, `N/A`.

The **Overall** verdict at the end of the report is `FAIL` if any item is
`FAIL`; otherwise `WARN` if any item is `WARN`; otherwise `PASS`. The
process exit code is `1` if overall is `FAIL`, else `0` — set only *after*
the report file has been saved, so the report is never lost.

### Parameters

| Parameter | Default | Purpose |
| --- | --- | --- |
| `-OutputPath` | Current user's Documents folder | Where the Markdown report is saved. |
| `-ExpectedUsernamePattern` | _(none)_ | Regex for the company's Windows username naming convention. Omit to get a SKIPPED verdict — never guessed. |
| `-ExpectedDomain` | _(none)_ | AD domain name the machine should be joined to. |
| `-ExpectedSSIDs` | _(none)_ | One or more allowed corporate Wi-Fi SSID(s). |
| `-StorageFreeThresholdPercent` | `15` | Fail if any fixed volume has less than this % free. |
| `-BatteryHealthThresholdPercent` | `70` | Fail if battery health % is below this. |
| `-JunkFilesWarnThresholdGB` | `5` | Warn (not fail) if reclaimable junk exceeds this many GB. |
| `-BloatwareKeywords` | Generic starter list (trial security suites, game promos, OEM trialware, toolbars, etc.) | Substrings matched case-insensitively against installed app DisplayNames. Override to tune to your company's own policy. |
| `-SkipWindowsUpdateCheck` | off | Skips the OS-update-freshness check entirely (useful on offline/locked-down machines); verdict becomes N/A. |
| `-WindowsUpdateTimeoutSeconds` | `45` | Max seconds to wait for the background Windows Update search before reporting UNKNOWN. |

### Quick Start — run with a single command (no clone required)

Run with all defaults (username and network checks will be SKIPPED until
company policy values are supplied):

```powershell
irm https://raw.githubusercontent.com/PECTECTH/PECScript/main/Test-ITCompliance.ps1 | iex
```

To pass company-policy parameters (`-ExpectedUsernamePattern`,
`-ExpectedDomain`, `-ExpectedSSIDs`, etc.), build a scriptblock first, the
same way as with `Invoke-SystemAudit.ps1`:

```powershell
$script = irm https://raw.githubusercontent.com/PECTECTH/PECScript/main/Test-ITCompliance.ps1
& ([scriptblock]::Create($script)) -ExpectedUsernamePattern '^[a-z]+\.[a-z]+$' -ExpectedDomain 'CORP' -ExpectedSSIDs 'CorpWiFi','CorpWiFi-5G'
```

Same security note as above: only run `irm | iex` one-liners against a
URL/repo you trust.

### Usage (after cloning the repo locally)

```powershell
.\Test-ITCompliance.ps1
```

```powershell
.\Test-ITCompliance.ps1 -ExpectedUsernamePattern '^[a-z]+\.[a-z]+$' -ExpectedDomain 'CORP' -ExpectedSSIDs 'CorpWiFi' -OutputPath 'C:\Reports'
```

If script execution is blocked by your local execution policy, run it with:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Test-ITCompliance.ps1
```

### Output

Saved to `-OutputPath` (default: Documents) as:

```
ITComplianceReport_<HOSTNAME>_<yyyyMMdd_HHmmss>.md
```

UTF-8 encoded, same hostname+timestamp convention as
`SystemAuditReport_*.md`, and also excluded from version control by
`.gitignore` since it contains real machine-specific data. The full path is
printed to the console when the script finishes, along with the overall
verdict.
