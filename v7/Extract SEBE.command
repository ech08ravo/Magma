#!/bin/bash
# Double-click in Finder: syncs CALM-dashboard-* folders and makes SEBE files for any new one.
cd "$(dirname "$0")"
python3 update-dashboard-data.py
echo; read -n 1 -s -r -p "Press any key to close"
