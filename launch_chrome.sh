#!/bin/bash
# ============================================================
# Launch Chrome with Remote Debugging for Excel Online MCP
# ============================================================
# Run this BEFORE starting the MCP server.
# This opens Chrome with CDP enabled so the MCP can connect.
# ============================================================

PORT=9222

# Detect OS and Chrome path
if [[ "$OSTYPE" == "darwin"* ]]; then
    CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
elif [[ "$OSTYPE" == "msys" || "$OSTYPE" == "cygwin" || "$OSTYPE" == "win32" ]]; then
    CHROME="C:/Program Files/Google/Chrome/Application/chrome.exe"
else
    CHROME="google-chrome"
fi

echo "=== Excel Online MCP — Chrome Launcher ==="
echo "Starting Chrome with --remote-debugging-port=$PORT"
echo "Chrome path: $CHROME"
echo ""
echo "IMPORTANT: Log into SharePoint/Excel Online in this Chrome window."
echo "Then start the MCP server in another terminal."
echo "============================================"

"$CHROME" \
    --remote-debugging-port=$PORT \
    --user-data-dir="$HOME/.chrome" \
    --no-first-run \
    --no-default-browser-check \
    "$@"
