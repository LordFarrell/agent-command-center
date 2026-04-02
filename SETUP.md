# Excel Online MCP + Dashboard — Setup Guide

## What This Is

Two interfaces for the same automation engine:

1. **MCP Server** (`excel_online_mcp.py`) — For Claude/LLM tool integration via MCP protocol
2. **Dashboard** (`dashboard.html` + `server.py`) — Standalone web UI with one-click actions, live terminal, and screenshot preview

Both use Playwright CDP to control your existing Chrome session. No Microsoft API keys, no OAuth, no login flow. Your browser IS the auth.

The core workflow: open the ChatGPT Excel add-in → set mode (Fast/Standard/Heavy) → inject a prompt → let ChatGPT act on the spreadsheet.

---

## Prerequisites

- Python 3.10+
- Google Chrome
- ChatGPT for Excel add-in installed in your Excel Online instance

---

## Install (one-time)

```bash
pip install mcp pydantic playwright httpx
playwright install chromium
```

---

## Launch Sequence

### Step 1: Start Chrome with CDP

**macOS/Linux:**
```bash
chmod +x launch_chrome.sh
./launch_chrome.sh
```

**Windows PowerShell:**
```powershell
.\launch_chrome.ps1
```

**Or manually:**
```bash
chrome --remote-debugging-port=9222 --user-data-dir="$HOME/.chrome"
```

### Step 2: Log into SharePoint

In the Chrome window that opens, navigate to your Excel Online workbook and make sure you're logged in.

### Step 3: Start the MCP Server

```bash
python excel_online_mcp.py
```

Or configure it in Claude Desktop — see `claude_desktop_config.json`.

---

## Claude Desktop Configuration

Copy this into your Claude Desktop MCP config (`~/Library/Application Support/Claude/claude_desktop_config.json` on macOS):

```json
{
  "mcpServers": {
    "excel_online_mcp": {
      "command": "python",
      "args": ["/full/path/to/excel_online_mcp.py"],
      "env": {
        "CDP_ENDPOINT": "http://localhost:9222"
      }
    }
  }
}
```

---

## Available Tools (15 total)

### Navigation & Sheets
| Tool | Description |
|------|-------------|
| `excel_navigate` | Open an Excel Online workbook URL |
| `excel_get_sheet_info` | List all sheet tabs, identify active sheet |
| `excel_switch_sheet` | Switch to a different sheet tab |

### Cell Operations
| Tool | Description |
|------|-------------|
| `excel_read_cells` | Read value/formula from a cell or range |
| `excel_write_cell` | Write a value or formula to a cell |
| `excel_batch_write` | Write multiple cells in one call |
| `excel_get_active_cell` | Get current cell reference and formula bar content |

### ChatGPT Add-in
| Tool | Description |
|------|-------------|
| `excel_chatgpt_open` | Open the ChatGPT add-in from the Home ribbon |
| `excel_chatgpt_prompt` | Inject a prompt and optionally wait for response |
| `excel_chatgpt_read_response` | Read the latest ChatGPT response |

### Formatting & Interaction
| Tool | Description |
|------|-------------|
| `excel_click_ribbon` | Click any ribbon button by label |
| `excel_keyboard` | Send keyboard shortcuts (Ctrl+C, etc.) |
| `excel_format_cells` | Apply bold/italic/underline/font size |
| `excel_find_replace` | Find and optionally replace text |

### Utility
| Tool | Description |
|------|-------------|
| `excel_screenshot` | Capture screenshot of current state |
| `excel_run_js` | Execute arbitrary JS in the page (escape hatch) |
| `excel_version` | Server version and connection diagnostics |

---

## Example Workflows

### Read data then ask ChatGPT to analyze it
```
1. excel_navigate → open the workbook
2. excel_read_cells → read A1:F20
3. excel_chatgpt_open → open the add-in
4. excel_chatgpt_prompt → "Summarize the trends in the selected data"
5. excel_chatgpt_read_response → get the analysis
```

### Have ChatGPT build a formula
```
1. excel_chatgpt_open → open the add-in
2. excel_chatgpt_prompt → "Write a VLOOKUP formula that..."
3. excel_chatgpt_read_response → get the formula
4. excel_write_cell → paste it into the target cell
```

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `CDP_ENDPOINT` | `http://localhost:9222` | Chrome DevTools Protocol endpoint |
| `EXCEL_URL` | *(empty)* | Default workbook URL (optional) |

---

## Troubleshooting

**"Browser not connected"** → Chrome isn't running with `--remote-debugging-port=9222`, or the port is different.

**"No browser pages found"** → Chrome is running but no tabs are open. Open Excel Online first.

**"ChatGPT add-in button not found"** → The add-in isn't installed, or it's on a different ribbon tab. Check that it's visible on the Home tab.

**"Could not find prompt input"** → The ChatGPT panel may not be fully loaded. Run `excel_chatgpt_open` first and wait a few seconds.

**Clipboard/range reading issues** → Some SharePoint tenants restrict clipboard access. Individual cell reads always work; range reads may fall back to cell-by-cell.
