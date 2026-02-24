# Office Channel Analyzer

Determines the **active update channel** for Microsoft 365 Apps for Enterprise by evaluating registry settings in Microsoft's documented priority order. Identifies configuration conflicts and provides root cause analysis for channel switching issues.

## Quick Start

### Option 1: Download the EXE (easiest)

Download `OfficeChannelAnalyzer.exe` from [Releases](../../releases/latest) and run it:

```powershell
# Live mode — queries the local machine's registry directly
OfficeChannelAnalyzer.exe

# File mode — analyze exported registry files from diagnostic collections
OfficeChannelAnalyzer.exe "C:\Logs\CustomerMachine"
```

### Option 2: Run the PowerShell script

```powershell
# Download and run (one-liner)
irm https://raw.githubusercontent.com/<owner>/OfficeChannelAnalyzer/main/OfficeChannelAnalyzer.ps1 -OutFile OfficeChannelAnalyzer.ps1; .\OfficeChannelAnalyzer.ps1

# Or clone the repo
git clone https://github.com/<owner>/OfficeChannelAnalyzer.git
cd OfficeChannelAnalyzer
.\OfficeChannelAnalyzer.ps1
```

> **Note:** Replace `<owner>` with the GitHub username or org when the repo is published.

## Modes

| Mode | When | Command |
|------|------|---------|
| **Live** | Run directly on a machine with M365 Apps installed | `.\OfficeChannelAnalyzer.ps1` (no arguments) |
| **File** | Analyze exported registry files from SaRA, OfficeDiag, etc. | `.\OfficeChannelAnalyzer.ps1 -Path "C:\Logs\Folder"` |

## What It Does

1. **Evaluates the 7-level channel priority table** — determines which registry value is actually controlling the update channel
2. **Identifies the active (winning) channel** — highlights it in the console and HTML report
3. **Detects conflicts** — finds settings that block channel switches (disabled auto-updates, GPO overrides, SCCM COM triggers, etc.)
4. **Provides root cause analysis** — actionable findings with severity levels
5. **Generates an action plan** — step-by-step remediation guidance
6. **Outputs an interactive HTML report** — `OfficePolicies.html` with expandable registry details

## Channel Priority Order

Microsoft 365 Apps evaluates these registry values **in priority order**. The **first configured value wins** and determines the active update channel.

| Priority | Management Type | Registry Value | Registry Path |
|:--------:|-----------------|----------------|---------------|
| **1st** | Cloud Update | `UpdatePath` | `HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate` |
| **2nd** | Cloud Update | `UpdateBranch` | `HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate` |
| **3rd** | Policy/GPO | `UpdatePath` | `HKLM\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate` |
| **4th** | Policy/GPO | `UpdateBranch` | `HKLM\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate` |
| **5th** | ODT | `UpdateUrl` | `HKLM\SOFTWARE\Microsoft\office\ClickToRun\Configuration` |
| **6th** | Unmanaged | `UnmanagedUpdateURL` | `HKLM\SOFTWARE\Microsoft\office\ClickToRun\Configuration` |
| **7th** | Unmanaged | `CDNBaseUrl` | `HKLM\SOFTWARE\Microsoft\office\ClickToRun\Configuration` |

> **Reference:** [Microsoft Learn — Change update channels for Microsoft 365 Apps](https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels)

## Output

### Console Report
- Machine summary (version, platform, active channel, target channel)
- Color-coded priority table showing which level is winning
- Findings with severity (ERROR / WARN / INFO)
- Numbered action plan with remediation steps

### HTML Report (`OfficePolicies.html`)
- Interactive priority table — click any row to expand all registry settings at that path
- Summary grid with key configuration values
- Findings and action plan formatted for sharing
- Reference table at the bottom

## Prerequisites

- **Windows** with PowerShell 5.1+ (built-in on Windows 10/11 and Server 2016+)
- **Live mode:** Microsoft 365 Apps (Click-to-Run) must be installed on the machine
- **File mode:** Registry export files in UTF-16LE `.txt` format (standard `reg export` output)

## Channel GUIDs Reference

| GUID | Channel Name |
|------|-------------|
| `492350f6-3a01-4f97-b9c0-c7c6ddf67d60` | Current Channel |
| `64256afe-f5d9-4f86-8936-8840a6a4f5be` | Current Channel Preview |
| `55336b82-a18d-4dd6-b5f6-9e5095c314a6` | Monthly Enterprise Channel |
| `7ffbc6bf-bc32-4f92-8982-f9dd17fd3114` | Semi-Annual Enterprise Channel |
| `b8f9b850-328d-4355-9145-c59439a0c4cf` | Semi-Annual Enterprise Channel Preview *(deprecated)* |
| `5440fd1f-7ecb-4221-8110-145efaa6372f` | Beta Channel |

## License

[MIT](LICENSE)
