#!/bin/bash
# Stops the Job Tracker for this session.
# The service will start again automatically on your next login.
# To prevent auto-start permanently, delete the plist from ~/Library/LaunchAgents/.

PLIST="$HOME/Library/LaunchAgents/com.efrat.jobtracker.plist"

if launchctl list | grep -q "com.efrat.jobtracker"; then
    launchctl unload "$PLIST"
    echo "✓ Job Tracker stopped."
else
    echo "Job Tracker is not currently running."
fi
