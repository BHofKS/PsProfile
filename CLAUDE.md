# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

Two PowerShell scripts that define and deploy a single Windows admin PowerShell profile across many Kansas State University servers. There is no build, no test suite, and no package manifest — the repo *is* the artifact.

- `windows_Microsoft.PowerShell_profile.ps1` — the master profile. Deployed to servers, and also self-updates from GitHub raw at runtime.
- `Bootstrap-Profile.ps1` — run once per server (elevated, as an `ADS\` account) to install the profile into all locations and create the machine-local override file.

The repo is edited/committed on macOS but the code only runs on Windows PowerShell 5.1+ / pwsh on domain-joined servers. You generally cannot execute or verify these scripts locally — review changes by reading, not running.

## Critical invariant: `$ProfileVersion`

`windows_Microsoft.PowerShell_profile.ps1` line 3 holds `$ProfileVersion = 'yyyymmdd##'` (exactly 10 digits, compared as a string, so `-gt` works lexicographically).

**Any change to the profile's content must bump this number in the same commit.** Deployed servers fetch `$MasterUrl` on every ADS-account shell start and only overwrite their local copy when the upstream version string is greater. A content change without a version bump never propagates. A malformed version line makes every deployed copy stop self-updating, and `Bootstrap-Profile.ps1` will refuse a download that lacks the line (it treats it as a bad/redirected fetch).

Same-day edits increment the trailing `##` (`2026081101` → `2026081102`).

`$MasterUrl` is hardcoded in both files and must stay pointed at the `main` branch raw URL — a change only takes effect on servers once it lands on `main`.

## Deployment topology

Two Windows accounts share each server, with a home-directory layout that varies by machine. Both scripts resolve paths the same way, and the logic must stay in sync between them:

- If `C:\Users\bh1.users\` exists → `$homePath` = that, `$adminPath` = `C:\Users\bh1\`
- Otherwise → `$homePath` = `C:\Users\bh1\`, `$adminPath` = `C:\Users\bh1.ads\`

Four profile copies exist, `Documents\{WindowsPowerShell,PowerShell}\Microsoft.PowerShell_profile.ps1` under each of `$adminPath` and `$homePath`. The **primary** is `$adminPath\Documents\WindowsPowerShell\...` (most sessions start as ADS).

Reconciliation on every shell start:
- ADS sessions: fetch upstream → update primary if newer → MD5-copy primary out to the other three.
- Users-side sessions: no upstream fetch; the Users-side `WindowsPowerShell` copy is the reference and only Users-side paths are written (no permission to the ADS side).

Any pre-existing symlink at a target path is deleted and replaced with a real file — the design deliberately uses copies, not links.

## Machine-local overrides

`C:\ProgramData\PowerShell\LocalProfile.ps1` (also exposed as `$env:LocalProfile`) is dot-sourced as the last statement of the profile so a server can override anything above it. Bootstrap creates it if missing and grants `Modify` to both accounts; it never overwrites an existing one. Server-specific content belongs there, not in this repo.

## Conventions when editing the profile

- Functions are short admin shortcuts, defined at top level in **alphabetical order**, each with a one-line `#` comment describing it. New shortcuts go in alphabetical position.
- MMC/exe launchers that need elevation use `Start-Process ... -Verb runas`.
- The sync block at the top ends by `Remove-Variable`-ing its scratch variables so they don't leak into the user's session. `$homePath` is intentionally left defined — many functions depend on it.
- Backups land in `$homePath\Documents\PowerShellProfileBackups\profile-<version>.ps1`, pruned to the newest 5 by name in both scripts.
- The profile runs on every shell start: failures must be non-fatal. Network calls are wrapped in try/catch with a short `-TimeoutSec`, and downloads are validated (HTML/XML sniff + `$ProfileVersion` presence) before being written.

## Testing a change

Push to `main`, then open a new PowerShell session as the ADS account on a target server — the profile self-updates and prints `Profile sync: primary updated <old> -> <new>`. To install or force a full re-lay-down on a server, run `Bootstrap-Profile.ps1` elevated (it self-elevates via RunAs if needed); `-MasterPath <file>` skips the download and installs a local file instead.
