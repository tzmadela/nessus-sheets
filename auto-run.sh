#!/bin/bash
# Nessus Auto-Run Script
# Runs specified scans and pushes to Google Sheets
# Edit the scan IDs below to match your scans

LOG_FILE="$HOME/nessus-sheets/logs/auto-run.log"
mkdir -p "$HOME/nessus-sheets/logs"

echo "==============================" >> "$LOG_FILE"
echo "Auto-run started: $(date)" >> "$LOG_FILE"
echo "==============================" >> "$LOG_FILE"

cd "$HOME/nessus-sheets"

# EDIT THESE SCAN IDs
node nessus-to-sheets.js --auto 5 39 42 >> "$LOG_FILE" 2>&1

echo "Auto-run finished: $(date)" >> "$LOG_FILE"
echo "" >> "$LOG_FILE"
