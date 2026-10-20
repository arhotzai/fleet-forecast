#!/bin/bash
cd "$(dirname "$0")"
# Load local secrets (API key) if present
[ -f .env.local ] && set -a && . ./.env.local && set +a
echo "Starting Fleet ERP…  (leave this window open; close it to stop the app)"
exec streamlit run app.py --server.port 8501 --browser.gatherUsageStats false
