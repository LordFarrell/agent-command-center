# ============================================================
# Launch Chrome with Remote Debugging for Excel Online MCP
# ============================================================
# Run this BEFORE starting the MCP server (Windows PowerShell).
# ============================================================

$Port = 9222
$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$UserDataDir = "$env:USERPROFILE\.chrome"

Write-Host "=== Excel Online MCP - Chrome Launcher ===" -ForegroundColor Cyan
Write-Host "Starting Chrome with --remote-debugging-port=$Port"
Write-Host "Chrome path: $ChromePath"
Write-Host ""
Write-Host "IMPORTANT: Log into SharePoint/Excel Online in this Chrome window." -ForegroundColor Yellow
Write-Host "Then start the MCP server in another terminal."
Write-Host "============================================"

& $ChromePath `
    --remote-debugging-port=$Port `
    --user-data-dir=$UserDataDir `
    --no-first-run `
    --no-default-browser-check
