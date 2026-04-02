@echo off
title Excel Online MCP Launcher
color 0A

echo.
echo  ================================================
echo   Excel Online MCP Launcher  v1.2026.4
echo  ================================================
echo.

:: Launch Chrome with remote debugging
echo [1/2] Launching Chrome on port 9222...
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%APPDATA%\chrome-mcp"
echo        Done. Open your Excel Online workbook in the Chrome window.
echo.

:: Start the Python server (stays open, shows live logs)
echo [2/2] Starting MCP server on http://127.0.0.1:5111
echo        Open dashboard.html in any browser to use the dashboard.
echo.
echo  ------------------------------------------------
echo.

cd /d "%~dp0"
C:\Python313\python.exe server.py

echo.
echo  Server stopped.
pause
