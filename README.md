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

Your machine acts as a bridge between the private Nessus server and Google Sheets. It connects to the Nessus API on the private network, then pushes data to Google Sheets over the internet.

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

    git clone https://github.com/tzmadela/nessus-sheets.git
    cd nessus-sheets

### 2. Install Dependencies

    npm install

### 3. Configure Environment

    cp .env.example .env
    nano .env

Fill in your values:

    NESSUS_URL=https://your-nessus-server
    NESSUS_ACCESS_KEY=your-access-key
    NESSUS_SECRET_KEY=your-secret-key
    SPREADSHEET_ID=your-spreadsheet-id
    GOOGLE_CREDENTIALS_PATH=./credentials.json

Lock the file:

    chmod 600 .env

### 4. Google Sheets API Setup

**Create a Service Account:**

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create or select a project
3. Enable **Google Sheets API** (APIs & Services > Enable APIs)
4. Go to **IAM & Admin > Service Accounts**
5. Click **+ Create Service Account**
6. Name it, click through to create

**Download Key:**

1. Click on the service account
2. Go to **Keys** tab
3. Click **Add Key > Create new key > JSON**
4. Save as credentials.json in project folder

**Share Your Google Sheet:**

1. Open credentials.json, find client_email
2. Open your Google Sheet
3. Click **Share**, paste the email, set to **Editor**

### 5. Get Your Spreadsheet ID

From the Google Sheet URL:

    https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit

## Usage

### Interactive Mode

    node nessus-to-sheets.js

Options:
1. Test Connection
2. List Scans
3. Pull Hosts and Vulns (Latest)
4. Pull Hosts and Vulns (Historical)
5. Exit

### Auto-Run Mode

    node nessus-to-sheets.js --auto 5 39 42

Runs without interaction — perfect for cron jobs.

### Finding Scan IDs

Run interactive mode, choose option 2 to list all scans.

## Scheduling (cron)

    crontab -e

Add these lines:

    # Tuesday at 9:00 AM
    0 9 * * 2 /bin/bash /path/to/nessus-sheets/auto-run.sh

    # Friday at 8:00 AM
    0 8 * * 5 /bin/bash /path/to/nessus-sheets/auto-run.sh

## Nessus API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| GET /server/status | Test connectivity |
| GET /scans | List all scans |
| GET /scans/{id} | Get hosts from a scan |
| GET /scans/{id}?history_id={hid} | Get historical scan data |
| GET /scans/{id}/hosts/{host_id} | Get vulnerabilities for a host |

## What Gets Excluded

- **Info-level vulnerabilities** (severity 0) are excluded from sheets
- Info COUNT is still shown on the Dashboard for reference
- This typically reduces data by 60-70%

## Security

- Credentials stored in .env and credentials.json
- Both excluded from git via .gitignore
- File permissions set to 600 (owner only)
- No secrets in any code file
- Google Sheets access via service account with Editor permissions only

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Cannot connect to Nessus | Check VPN connection |
| ENOTFOUND | DNS issue — verify hostname |
| 401 Unauthorized | Check API keys in .env |
| Google API error | Enable Sheets API in Cloud Console |
| Permission denied on sheet | Share sheet with service account email |
| Cell limit exceeded | Delete all tabs and re-run |
| Percentages over 100% | Update to latest script version |

## Future Plans

- Migration to Microsoft Excel on SharePoint
- Azure Function deployment for company-independent execution
- Microsoft Graph API integration

## License

MIT License
