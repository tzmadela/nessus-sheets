# Nessus to Google Sheets Automation

A Node.js tool that pulls vulnerability scan data from a Nessus server and automatically populates a Google Sheet with organized, filterable results.

## What This Does

This tool connects to your Nessus vulnerability scanner via its REST API, retrieves scan results including all hosts and their vulnerabilities, and pushes everything into a Google Sheet for easy viewing, filtering, and sharing.

### The Problem

- Nessus is behind a VPN — not everyone has access
- Exporting reports manually is time-consuming with hundreds of hosts
- Teams need a simple spreadsheet they can filter and search
- Data needs to be refreshed automatically on a schedule

### The Solution
Your Machine (on VPN)
|
|-- 1. Connects to Nessus API (private network)
| Retrieves all hosts and vulnerabilities
|
|-- 2. Pushes data to Google Sheets API (internet)
Creates organized, filterable tabs

text

Your machine acts as a bridge between the private Nessus server and Google Sheets.

## Features

- **Automated data pull** from Nessus REST API
- **Historical scan support** — pull any past scan run
- **Auto-scheduling** via cron (Tuesday 9am, Friday 8am)
- **Progress tracking** with ETA for large scans
- **Auto-save** every 100 hosts to prevent data loss
- **Smart sheet sizing** — dynamic row allocation with growth buffer
- **Auto-clear** — clears old data before each run (preserves Run Log)
- **Filtered sheets** — Google Sheets filters auto-applied
- **Run logging** — tracks every execution with timestamps
- **CVSS-aligned risk scoring** — hosts ranked by weighted vulnerability severity
- **Info exclusion** — informational alerts excluded to reduce noise
- **Secure credentials** — all secrets in .env file, never in code

## Google Sheet Output

| Tab | Description |
|-----|-------------|
| **Dashboard** | Severity summary, scan info, top 30 riskiest hosts |
| **Hosts** | All scanned hosts with severity counts and risk scores |
| **All Vulnerabilities** | Every vulnerability (excl. info) — filterable by hostname |
| **Critical and High** | Only Critical and High severity for quick action |
| **Run Log** | History of every script execution (preserved across runs) |

### Finding a Specific Host's Vulnerabilities

1. Go to "All Vulnerabilities" tab
2. Click the filter dropdown on "Hostname" column
3. Select the host you want
4. Only that host's vulnerabilities are shown

## Risk Score

Each host gets a risk score using CVSS-aligned severity weights:
Risk Score = (Critical x 10) + (High x 5) + (Medium x 2) + (Low x 1)

text

| Severity | Weight | Rationale |
|----------|--------|-----------|
| Critical | x10 | CVSS 9.0-10.0 — immediate action required |
| High | x5 | CVSS 7.0-8.9 — serious risk |
| Medium | x2 | CVSS 4.0-6.9 — moderate risk |
| Low | x1 | CVSS 0.1-3.9 — minor risk |

Hosts are sorted by risk score so the most vulnerable appear first.

## Prerequisites

- **Node.js** v18 or higher — [Download](https://nodejs.org)
- **Nessus Professional/Expert** with API access
- **Google Cloud** account with Sheets API enabled
- **Network access** to Nessus server (e.g., VPN)

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/nessus-sheets.git
cd nessus-sheets
2. Install Dependencies
bash
npm install
3. Configure Environment
bash
cp .env.example .env
nano .env
Fill in your values:

text
NESSUS_URL=https://your-nessus-server:8834
NESSUS_ACCESS_KEY=your-access-key
NESSUS_SECRET_KEY=your-secret-key
SPREADSHEET_ID=your-spreadsheet-id
GOOGLE_CREDENTIALS_PATH=./credentials.json
Lock the file:

bash
chmod 600 .env
4. Google Sheets API Setup
Create a Service Account
Go to Google Cloud Console
Create or select a project
Enable Google Sheets API (APIs & Services > Enable APIs)
Go to IAM & Admin > Service Accounts
Click + Create Service Account
Name it, click through to create
Download Key
Click on the service account
Go to Keys tab
Click Add Key > Create new key > JSON
Save as credentials.json in project folder
bash
chmod 600 credentials.json
Share Your Google Sheet
Open credentials.json, find client_email
Open your Google Sheet
Click Share, paste the email, set to Editor
5. Get Your Spreadsheet ID
From the Google Sheet URL:

text
https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit
Usage
Interactive Mode
bash
node nessus-to-sheets.js
text
Options:
  1. Test Connection
  2. List Scans
  3. Pull Hosts & Vulns (Latest)
  4. Pull Hosts & Vulns (Historical)
  5. Exit
Auto-Run Mode
bash
node nessus-to-sheets.js --auto 5 39 42
Runs without interaction — perfect for cron jobs.

Finding Scan IDs
Run interactive mode, choose option 2 to list all scans.

Scheduling
Using cron (macOS/Linux)
bash
crontab -e
cron
# Tuesday at 9:00 AM
0 9 * * 2 /bin/bash /path/to/nessus-sheets/auto-run.sh

# Friday at 8:00 AM
0 8 * * 5 /bin/bash /path/to/nessus-sheets/auto-run.sh
Nessus API Endpoints Used
Endpoint	Purpose
GET /server/status	Test connectivity
GET /scans	List all scans
GET /scans/{id}	Get hosts from a scan
GET /scans/{id}?history_id={hid}	Get historical scan data
GET /scans/{id}/hosts/{host_id}	Get vulnerabilities for a host
What Gets Excluded
Info-level vulnerabilities (severity 0) are excluded from sheets
Info COUNT is still shown on the Dashboard for reference
This typically reduces data by 60-70%
Data Flow
text
1. Script starts
2. Clears old data sheets (keeps Run Log)
3. Connects to Nessus API
4. Pulls all hosts from specified scan
5. For each host, pulls all vulnerabilities
6. Skips info-level findings
7. Writes Hosts sheet (sorted by risk score)
8. Writes All Vulnerabilities sheet (sorted by severity)
9. Writes Critical and High sheet
10. Creates Dashboard with summary
11. Logs the run to Run Log
12. Done
Project Structure
text
nessus-sheets/
├── nessus-to-sheets.js    # Main script
├── auto-run.sh            # Scheduled run script
├── package.json           # Dependencies
├── .env                   # Secrets (NOT in git)
├── .env.example           # Template for .env
├── credentials.json       # Google API key (NOT in git)
├── .gitignore             # Protects secrets
├── README.md              # This file
└── logs/                  # Auto-run logs (NOT in git)
Security
Credentials stored in .env and credentials.json
Both excluded from git via .gitignore
File permissions set to 600 (owner only)
No secrets in any code file
Google Sheets access via service account with Editor permissions only
Troubleshooting
Problem	Solution
Cannot connect to Nessus	Check VPN connection
ENOTFOUND	DNS issue — verify hostname
401 Unauthorized	Check API keys in .env
Google API error	Enable Sheets API in Cloud Console
Permission denied on sheet	Share sheet with service account email
Cell limit exceeded	Old sheets too large — delete all tabs and re-run
Percentages over 100%	Update to latest script version
Future Plans
Migration to Microsoft Excel on SharePoint
Azure Function deployment for company-independent execution
Microsoft Graph API integration
License
MIT License
