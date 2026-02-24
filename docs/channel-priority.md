# Update Channel Priority — How It Works

When Microsoft 365 Apps needs to determine which update channel to use, it evaluates registry values in a strict priority order. **The first configured value it finds wins.**

This means a Cloud Update policy (Priority 1-2) will always override a GPO setting (Priority 3-4), which will always override an ODT configuration (Priority 5), and so on.

## Why This Matters

The most common reason a channel switch "doesn't work" is because a **higher-priority setting is overriding the intended channel**. For example:

- An admin sets `UpdateBranch = MonthlyEnterprise` via GPO (Priority 4)
- But Cloud Update has `UpdatePath` set at Priority 1
- The GPO setting is completely ignored because Priority 1 wins

## The Priority Table

| Priority | Source | Registry Value | Registry Path | Notes |
|:--------:|--------|----------------|---------------|-------|
| 1st | Cloud Update | `UpdatePath` | `HKLM\...\cloud\office\16.0\Common\officeupdate` | URL pointing to update source |
| 2nd | Cloud Update | `UpdateBranch` | `HKLM\...\cloud\office\16.0\Common\officeupdate` | Branch name (e.g., `MonthlyEnterprise`) |
| 3rd | Policy/GPO | `UpdatePath` | `HKLM\...\office\16.0\Common\officeupdate` | Set by Group Policy or Intune ADMX |
| 4th | Policy/GPO | `UpdateBranch` | `HKLM\...\office\16.0\Common\officeupdate` | Set by Group Policy or Intune ADMX |
| 5th | ODT | `UpdateUrl` | `HKLM\...\ClickToRun\Configuration` | Stamped by Office Deployment Tool |
| 6th | Unmanaged | `UnmanagedUpdateURL` | `HKLM\...\ClickToRun\Configuration` | Only present on unmanaged devices |
| 7th | Unmanaged | `CDNBaseUrl` | `HKLM\...\ClickToRun\Configuration` | Stamped at install time; the "default" |

## Common Gotchas

### CDNBaseUrl stuck on Semi-Annual
If Office was originally deployed with a Semi-Annual config (e.g., via SCCM with `Channel=SemiAnnual`), the `CDNBaseUrl` at Priority 7 will point to the SAC GUID. If no higher-priority setting overrides it, the device stays on Semi-Annual even though you "moved" it in the admin center.

**Fix:** Set a management policy at Priority 1-4 to explicitly direct the channel.

### Auto-updates disabled by GPO
Even if the correct channel is configured, setting `enableautomaticupdates = dword:0` via GPO blocks the Office update engine entirely. Cloud Update can override this if `ignoregpo = dword:1` is set.

### OfficeMgmtCOM = 1
This tells Office to wait for SCCM/ConfigMgr to trigger updates rather than updating autonomously. If SCCM isn't deploying the right channel's updates, the device sits idle.

## Reference

[Microsoft Learn — Change update channels for Microsoft 365 Apps](https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels)
