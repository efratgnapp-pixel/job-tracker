#!/bin/bash
# Loads the Job Tracker service for this session (also auto-starts on every login).
# Safe to run even after a reboot — launchd handles the auto-start itself.

PLIST="$HOME/Library/LaunchAgents/com.efrat.jobtracker.plist"

if launchctl list | grep -q "com.efrat.jobtracker"; then
    echo "Job Tracker is already running."
else
    launchctl load "$PLIST"
    echo "✓ Job Tracker started."
fi

echo "  → http://localhost:3001/job-tracker.html"
echo "  → Logs: ~/job-tracker/server.log"
