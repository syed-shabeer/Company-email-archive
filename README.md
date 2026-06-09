# Stellantis Email Archive & Search Tool

An internal productivity tool built to solve a real problem at Stellantis India —
searching and retrieving important project emails was slow and manual. This tool
automates the entire process of backing up Outlook emails and makes them instantly
searchable through a browser-based interface.

## The Problem
Project-critical emails were buried in Outlook folders with no quick way to search
or retrieve them across dates, senders, or keywords. Manual searching was
time-consuming and error-prone.

## The Solution
A lightweight local tool that:
- Automatically backs up Outlook emails to structured local files using Python
- Serves a clean browser-based search interface via a local HTTP server
- Lets users filter and retrieve emails by keyword, sender, date, and more
- Runs on a scheduled basis every weekday via Windows Task Scheduler
- Works on both battery and AC power with no manual intervention needed

## How to Use
1. Run `Stellantis_Email_Archive.bat` to launch the search viewer in your browser
2. Run `SetBackupSchedule.bat` as Administrator to set up the daily auto-backup
3. Enter your preferred backup time when prompted (e.g. 10:30 AM)
4. The tool will back up and archive emails automatically every weekday

## Files
- `outlook_backup.py` — Core Python script that connects to Outlook and backs up emails
- `search.html` — Browser-based search and viewer interface
- `Stellantis_Email_Archive.bat` — Launches the local server and opens the viewer
- `SetBackupSchedule.bat` — Registers the backup as a scheduled Windows Task

## Built With
Python · HTML · Windows Batch Scripting · Windows Task Scheduler API
