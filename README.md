# Company Email Archive & Search Tool

## The Problem
Large organizations enforce email retention policies that automatically purge
emails after a fixed period — typically 90 days to 6 months. For engineers
working on long-running projects that span 1 to 2 years, this creates a serious
problem: critical project emails — technical decisions, design approvals, client
instructions, issue resolutions — get permanently deleted before the project
even closes.

The only official workaround was to manually do a "Save As" for each important
email in Outlook, one by one. This was:
- Time-consuming and easy to forget
- Inconsistent across team members
- Impossible to search or retrieve quickly later

There was no central, searchable record of past project communications.

## The Solution
A lightweight, fully local tool that automates the backup of Outlook emails
and provides a clean browser-based interface to search and retrieve them —
no cloud dependency, no IT involvement, no manual effort.

Built entirely using Python, HTML, and Windows Batch Scripting, the tool runs
silently in the background on a scheduled basis and keeps a structured,
searchable archive of all important emails on the local machine.

## Features
- **Automated Backup** — Python script connects to Outlook and archives emails
  to structured local files automatically, no manual "Save As" needed
- **Browser-Based Search Interface** — Clean HTML viewer served via a local
  HTTP server; search and filter emails by keyword, sender, subject, or date
- **Scheduled Daily Runs** — Windows Task Scheduler integration runs the backup
  every weekday at a user-defined time
- **Battery & AC Support** — Scheduler configured to run regardless of whether
  the machine is on battery or plugged in
- **Fallback Port Handling** — Automatically switches to an alternate port if
  the default is already in use
- **Zero Cloud Dependency** — Everything stays local; no data leaves the machine

## How to Use

### First-Time Setup
1. Run `SetBackupSchedule.bat` as Administrator
2. Enter your preferred daily backup time when prompted (e.g. `10:30 AM`)
3. The tool registers itself as a Windows Scheduled Task and runs automatically
   every weekday from that point on

### Viewing Archived Emails
1. Run `Company_Email_Archive.bat`
2. The local server starts and the search interface opens in your default browser
3. Search, filter, and browse your archived emails
4. Close the window when done — the server shuts down automatically

## Files
- `outlook_backup.py` — Core Python script that connects to Outlook via COM
  interface and exports emails to structured local files
- `search.html` — Browser-based search and viewer interface
- `Company_Email_Archive.bat` — Launches the local HTTP server and opens
  the viewer in the browser
- `SetBackupSchedule.bat` — Registers the backup script as a recurring
  Windows Scheduled Task with user-defined timing

## Tech Stack
- **Python** — Outlook COM automation, file handling, local HTTP server
- **HTML / JavaScript** — Search interface and email viewer
- **Windows Batch Scripting** — Launcher and scheduler setup
- **Windows Task Scheduler API** — Automated daily execution via PowerShell
  and schtasks

## Why This Matters
On projects lasting over a year, having a reliable, searchable record of past
emails is not just convenient — it is essential. Design decisions, scope changes,
approvals, and technical clarifications are all communicated over email. Without
an archive, teams lose institutional memory and spend hours trying to reconstruct
context. This tool solves that problem quietly, automatically, and without
requiring any changes to how the team works.
