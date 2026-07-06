# System Audit Report

A single, self-contained Windows PowerShell script that collects a snapshot of your
local machine's configuration and writes it out as a readable Markdown report.

It is designed to run out of the box with **no Administrator privileges required**
for its core functionality. If a particular section would benefit from elevated
permissions to return complete data, that section fails gracefully on its own and
notes that Administrator rights may be needed — the rest of the report is still
produced normally.

## What it does

Running the script produces a single Markdown file containing the following
sections, in order:

1. **Report Header** — generated timestamp, computer name, current user, OS
   caption/version/build, architecture, and system uptime.
2. **Running Processes** — every running process (or the top N by memory usage, if
   requested), with process ID, name, CPU time, working set memory (MB), and start
   time. Sorted by memory usage, descending.
3. **Installed Applications** — programs listed in the Windows uninstall registry
   keys (64-bit, 32-bit/WOW6432Node, and per-user hives), with name, version,
   publisher, and install date. Deduplicated and sorted alphabetically. This does
   **not** use `Win32_Product`/`Get-WmiObject`, which is known to be slow and can
   trigger unwanted MSI self-repair operations.
4. **CPU** — processor name, physical core count, logical processor count, max clock
   speed, and manufacturer.
5. **Memory (RAM)** — total installed physical memory plus a per-module breakdown
   (capacity, speed, manufacturer, part number) where available.
6. **Disks** — physical disks (model, size, interface/media type) and logical
   volumes (drive letter, total size, free space, percent used).
7. **Network Adapters / MAC Addresses** — adapter name, MAC address, link speed, and
   status, limited to adapters that actually report a MAC address.
8. **System / Hardware Summary** — system manufacturer and model, plus the BIOS
   serial number (which may legitimately be blank on virtual machines).

The report opens with a table of contents that links to each section.

## Requirements

- Windows PowerShell 5.1 (built into Windows 10/11) — also runs under PowerShell 7+.
- No Administrator privileges required for normal use. A few data points in
  specific sections may be incomplete without elevation; those sections will note
  it inline rather than failing the whole script.

## How to run it

From the `SystemAuditReport` folder, run the script with no arguments to use all
defaults — the report is saved to your Documents folder and every process/app is
listed:

```powershell
.\SystemAuditReport\Invoke-SystemAudit.ps1
```

To customize where the report is saved and/or limit the process list to the top N
processes by memory usage:

```powershell
.\SystemAuditReport\Invoke-SystemAudit.ps1 -OutputPath 'C:\Reports' -TopProcessCount 25
```

Parameters:

| Parameter          | Type   | Default                                   | Description                                                     |
| ------------------ | ------ | ------------------------------------------ | ----------------------------------------------------------------- |
| `-OutputPath`       | string | Current user's Documents folder            | Directory where the generated Markdown report is saved.          |
| `-TopProcessCount`  | int    | `0` (show all processes)                   | If greater than 0, limits the Running Processes section to the top N processes by memory usage. |

## Where the report lands

By default, the report is saved to your Documents folder as:

```
SystemAuditReport_<yyyyMMdd_HHmmss>.md
```

For example: `SystemAuditReport_20260706_084014.md`. The full path is printed to
the console in green when the script finishes.

## Privacy / network note

This script only reads information that is already local to your machine (running
processes, registry uninstall entries, CIM/WMI hardware classes, network adapter
configuration). **It does not send any data anywhere over the network.** The
generated report is a plain local Markdown file that you fully control.

Because the report contains machine-specific details (installed software, BIOS
serial number, MAC addresses, etc.), avoid committing generated report files to a
public repository — see `.gitignore` in this repo, which already excludes them.

## Sample output

The tables below show only the column headers produced by the script — no real
machine data is included here.

**Running Processes**

| Id | ProcessName | CPU(s) | WorkingSetMB | StartTime |
| --- | --- | --- | --- | --- |

**Installed Applications**

| DisplayName | DisplayVersion | Publisher | InstallDate |
| --- | --- | --- | --- |

**CPU**

| Name | NumberOfCores | NumberOfLogicalProcessors | MaxClockSpeed(MHz) | Manufacturer |
| --- | --- | --- | --- | --- |

**Memory Modules**

| CapacityGB | Speed(MHz) | Manufacturer | PartNumber |
| --- | --- | --- | --- |

**Physical Disks**

| Model | SizeGB | InterfaceType | MediaType |
| --- | --- | --- | --- |

**Logical Volumes**

| DriveLetter | TotalGB | FreeGB | PercentUsed |
| --- | --- | --- | --- |

**Network Adapters / MAC Addresses**

| Name | MacAddress | LinkSpeed | Status |
| --- | --- | --- | --- |
