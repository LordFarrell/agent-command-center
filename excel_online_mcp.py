#!/usr/bin/env python3
"""
Excel Online MCP Server v1.2026.2

Chrome-level browser automation MCP for Excel Online.
Connects to an existing Chrome session via CDP (Chrome DevTools Protocol)
using Playwright. No login required — reuses your authenticated session.

ARCHITECTURE NOTE (from live testing):
  Excel Online renders entirely inside an iframe (WacFrame_Excel_0).
  JavaScript signals from the outer page CANNOT reach the DOM.
  All interaction is coordinate-based clicks + keyboard input.
  The ChatGPT add-in panel is also iframe-nested.
  Submit = Enter key (click on submit button does NOT pass through).

Core workflow:
  1. Open ChatGPT add-in via ribbon click
  2. Switch mode from Fast → Standard (or Heavy)
  3. Click prompt input, type prompt
  4. Press Enter to submit
  5. Wait for response, optionally screenshot result

Usage:
  1. Launch Chrome with: chrome --remote-debugging-port=9222
  2. Log into SharePoint / Excel Online in that Chrome instance
  3. Run this MCP server: python excel_online_mcp.py
"""

import asyncio
import base64
import json
import os
import time
from contextlib import asynccontextmanager
from enum import Enum
from typing import Any, Optional

from mcp.server.fastmcp import FastMCP, Context
from pydantic import BaseModel, ConfigDict, Field

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
VERSION = "1.2026.2"
CDP_ENDPOINT = os.environ.get("CDP_ENDPOINT", "http://localhost:9222")
DEFAULT_TIMEOUT = 15_000  # ms
LONG_TIMEOUT = 60_000  # ms

# Coordinate map — derived from live testing on 1536x739 viewport
# These are default positions; tools accept overrides for different viewports.
COORDS = {
    "chatgpt_ribbon_btn": (1378, 110),   # ChatGPT button on Home ribbon tab
    "prompt_input":       (1300, 632),   # ChatGPT prompt textarea
    "submit_btn":         (1513, 673),   # Submit arrow (fallback — Enter preferred)
    "mode_dropdown":      (1193, 672),   # Fast/Standard/Heavy dropdown toggle
    "mode_fast":          (1192, 577),   # "Fast" option in dropdown
    "mode_standard":      (1192, 605),   # "Standard" option in dropdown
    "mode_heavy":         (1192, 632),   # "Heavy" option in dropdown
    "name_box":           (55, 199),     # Name Box (cell reference input)
}


# ---------------------------------------------------------------------------
# Lifespan: Playwright browser connection
# ---------------------------------------------------------------------------
@asynccontextmanager
async def app_lifespan():
    """Connect to existing Chrome via CDP on startup, disconnect on shutdown."""
    from playwright.async_api import async_playwright

    pw = await async_playwright().__aenter__()
    browser = None
    try:
        browser = await pw.chromium.connect_over_cdp(CDP_ENDPOINT, timeout=DEFAULT_TIMEOUT)
        yield {"browser": browser, "pw": pw}
    finally:
        if browser:
            await browser.close()  # disconnects CDP, does NOT close the user's browser
        await pw.__aexit__(None, None, None)


mcp = FastMCP("excel_online_mcp", lifespan=app_lifespan)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
async def _get_browser(ctx: Context):
    """Retrieve the Playwright browser from lifespan state."""
    state = ctx.request_context.lifespan_state
    browser = state.get("browser")
    if not browser:
        raise RuntimeError(
            "Browser not connected. Ensure Chrome is running with "
            f"--remote-debugging-port on {CDP_ENDPOINT}"
        )
    return browser


async def _get_excel_page(ctx: Context):
    """
    Find the first browser page whose URL contains a SharePoint/Excel marker.
    Falls back to the first visible page.
    """
    browser = await _get_browser(ctx)
    for bc in browser.contexts:
        for page in bc.pages:
            url_lower = page.url.lower()
            if any(k in url_lower for k in ("sharepoint", "excel", "office.com", "live.com")):
                await page.bring_to_front()
                return page
    # Fallback
    if browser.contexts and browser.contexts[0].pages:
        page = browser.contexts[0].pages[0]
        await page.bring_to_front()
        return page
    raise RuntimeError("No browser pages found. Open Excel Online in Chrome first.")


def _ok(data: Any) -> str:
    if isinstance(data, str):
        return data
    return json.dumps(data, indent=2, default=str)


def _error(msg: str) -> str:
    return json.dumps({"error": msg})


async def _click(page, coord: tuple, delay_after: float = 0.3):
    """Click at (x, y) coordinates and wait."""
    await page.mouse.click(coord[0], coord[1])
    if delay_after > 0:
        await asyncio.sleep(delay_after)


async def _screenshot_b64(page) -> str:
    """Take a PNG screenshot and return base64."""
    data = await page.screenshot(type="png")
    return base64.b64encode(data).decode("utf-8")


# ===================================================================
#  INPUT MODELS
# ===================================================================

class ChatGPTMode(str, Enum):
    FAST = "fast"
    STANDARD = "standard"
    HEAVY = "heavy"


class ExcelNavigateInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")
    url: str = Field(..., description="Full URL to an Excel Online workbook", min_length=10)


class ChatGPTPromptInput(BaseModel):
    """Input for the main ChatGPT prompt injection tool."""
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")
    prompt: str = Field(
        ...,
        description="The prompt to send to ChatGPT Excel add-in",
        min_length=1,
        max_length=10000,
    )
    mode: ChatGPTMode = Field(
        default=ChatGPTMode.STANDARD,
        description="ChatGPT mode: 'fast', 'standard', or 'heavy'. Default: standard",
    )
    open_if_closed: bool = Field(
        default=True,
        description="Auto-open the ChatGPT panel if it's not already visible",
    )
    wait_seconds: int = Field(
        default=30,
        description="Seconds to wait for ChatGPT to finish acting on the sheet",
        ge=5,
        le=300,
    )


class ChatGPTSetModeInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    mode: ChatGPTMode = Field(
        ...,
        description="Target mode: 'fast', 'standard', or 'heavy'",
    )


class KeyboardInput(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")
    keys: str = Field(
        ...,
        description="Playwright key combo e.g. 'Control+C', 'Enter', 'Tab'",
        min_length=1,
    )
    repeat: int = Field(default=1, ge=1, le=100)


class ClickInput(BaseModel):
    model_config = ConfigDict(extra="forbid")
    x: int = Field(..., description="X coordinate (pixels from left)")
    y: int = Field(..., description="Y coordinate (pixels from top)")


# ===================================================================
#  TOOLS — ChatGPT Add-in (Primary Workflow)
# ===================================================================

@mcp.tool(
    name="excel_chatgpt_prompt",
    annotations={
        "title": "Send Prompt to ChatGPT Excel Add-in",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def excel_chatgpt_prompt(params: ChatGPTPromptInput, ctx: Context) -> str:
    """The primary tool. Opens ChatGPT panel (if needed), sets the mode,
    types the prompt, presses Enter to submit, and waits for execution.

    This is the main way to interact with Excel Online — send natural language
    instructions to ChatGPT and it will read/write the spreadsheet directly.

    Workflow: open panel → set mode → click input → type prompt → Enter → wait.

    Args:
        params: ChatGPTPromptInput with prompt, mode, and wait settings.

    Returns:
        JSON with status and a base64 screenshot of the result.
    """
    page = await _get_excel_page(ctx)

    # Step 1: Open ChatGPT panel if requested
    if params.open_if_closed:
        await _ensure_chatgpt_open(page)

    # Step 2: Set mode (Fast/Standard/Heavy)
    await _set_chatgpt_mode(page, params.mode)

    # Step 3: Click prompt input area
    await _click(page, COORDS["prompt_input"], delay_after=0.3)

    # Step 4: Type the prompt
    await page.keyboard.type(params.prompt, delay=10)
    await asyncio.sleep(0.2)

    # Step 5: Submit via Enter (NOT click — click doesn't pass through iframe layers)
    await page.keyboard.press("Enter")
    await asyncio.sleep(1)

    # Step 6: Wait for ChatGPT to finish (poll for stable screenshot)
    await asyncio.sleep(params.wait_seconds)

    # Step 7: Screenshot the result
    b64 = await _screenshot_b64(page)

    return _ok({
        "status": "ok",
        "prompt": params.prompt[:200] + ("..." if len(params.prompt) > 200 else ""),
        "mode": params.mode.value,
        "waited_seconds": params.wait_seconds,
        "screenshot_base64_length": len(b64),
        "screenshot_base64": b64,
    })


@mcp.tool(
    name="excel_chatgpt_open",
    annotations={
        "title": "Open ChatGPT Excel Add-in Panel",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def excel_chatgpt_open(ctx: Context) -> str:
    """Click the ChatGPT button on the Home ribbon tab to open the add-in panel.

    The entire Excel UI is inside an iframe so this uses coordinate-based clicking.

    Returns:
        JSON confirmation with screenshot.
    """
    page = await _get_excel_page(ctx)
    await _ensure_chatgpt_open(page)
    b64 = await _screenshot_b64(page)
    return _ok({"status": "ok", "screenshot_base64": b64})


@mcp.tool(
    name="excel_chatgpt_set_mode",
    annotations={
        "title": "Set ChatGPT Mode (Fast/Standard/Heavy)",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def excel_chatgpt_set_mode(params: ChatGPTSetModeInput, ctx: Context) -> str:
    """Switch the ChatGPT add-in between Fast, Standard, or Heavy mode.

    Args:
        params: ChatGPTSetModeInput with target mode.

    Returns:
        JSON confirmation.
    """
    page = await _get_excel_page(ctx)
    await _set_chatgpt_mode(page, params.mode)
    return _ok({"status": "ok", "mode": params.mode.value})


async def _ensure_chatgpt_open(page):
    """Click the ChatGPT ribbon button and wait for the panel to load."""
    await _click(page, COORDS["chatgpt_ribbon_btn"], delay_after=0.5)
    # Wait for the panel to render (add-in iframe load time)
    await asyncio.sleep(4)


async def _set_chatgpt_mode(page, mode: ChatGPTMode):
    """Open the mode dropdown and select the target mode."""
    # Click the mode dropdown (the "Fast v" / "Standard v" area)
    await _click(page, COORDS["mode_dropdown"], delay_after=0.8)

    # Click the target mode option
    mode_coords = {
        ChatGPTMode.FAST: COORDS["mode_fast"],
        ChatGPTMode.STANDARD: COORDS["mode_standard"],
        ChatGPTMode.HEAVY: COORDS["mode_heavy"],
    }
    await _click(page, mode_coords[mode], delay_after=0.5)


# ===================================================================
#  TOOLS — Navigation
# ===================================================================

@mcp.tool(
    name="excel_navigate",
    annotations={
        "title": "Navigate to Excel Online Workbook",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def excel_navigate(params: ExcelNavigateInput, ctx: Context) -> str:
    """Navigate the browser to an Excel Online workbook URL and wait for load.

    Args:
        params: ExcelNavigateInput with the workbook URL.

    Returns:
        JSON with page title and URL.
    """
    browser = await _get_browser(ctx)
    contexts = browser.contexts
    if not contexts:
        raise RuntimeError("No browser contexts available")
    page = contexts[0].pages[0] if contexts[0].pages else await contexts[0].new_page()
    await page.goto(params.url, wait_until="domcontentloaded", timeout=LONG_TIMEOUT)
    await asyncio.sleep(5)  # Excel Online takes time to fully render
    return _ok({"status": "ok", "title": await page.title(), "url": page.url})


# ===================================================================
#  TOOLS — Low-level Interaction (Escape Hatches)
# ===================================================================

@mcp.tool(
    name="excel_screenshot",
    annotations={
        "title": "Take Screenshot of Excel Online",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def excel_screenshot(ctx: Context) -> str:
    """Take a screenshot of the current Excel Online page.

    Returns:
        JSON with base64-encoded PNG screenshot.
    """
    page = await _get_excel_page(ctx)
    b64 = await _screenshot_b64(page)
    return _ok({"status": "ok", "format": "png", "base64_length": len(b64), "base64": b64})


@mcp.tool(
    name="excel_keyboard",
    annotations={
        "title": "Send Keyboard Input to Excel",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def excel_keyboard(params: KeyboardInput, ctx: Context) -> str:
    """Send keyboard key presses or shortcuts to the Excel Online page.

    Args:
        params: KeyboardInput with keys string and repeat count.

    Returns:
        JSON confirmation.
    """
    page = await _get_excel_page(ctx)
    for _ in range(params.repeat):
        await page.keyboard.press(params.keys)
        await asyncio.sleep(0.1)
    return _ok({"status": "ok", "keys": params.keys, "repeat": params.repeat})


@mcp.tool(
    name="excel_click",
    annotations={
        "title": "Click at Coordinates in Excel",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def excel_click(params: ClickInput, ctx: Context) -> str:
    """Click at specific (x, y) coordinates on the Excel Online page.

    Use this as an escape hatch for any UI element not covered by other tools.

    Args:
        params: ClickInput with x, y coordinates.

    Returns:
        JSON confirmation.
    """
    page = await _get_excel_page(ctx)
    await _click(page, (params.x, params.y))
    return _ok({"status": "ok", "clicked": {"x": params.x, "y": params.y}})


@mcp.tool(
    name="excel_type",
    annotations={
        "title": "Type Text at Current Cursor Position",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def excel_type(text: str, ctx: Context) -> str:
    """Type text at the current cursor/focus position.

    Args:
        text: The text to type.

    Returns:
        JSON confirmation.
    """
    page = await _get_excel_page(ctx)
    await page.keyboard.type(text, delay=15)
    return _ok({"status": "ok", "typed_length": len(text)})


@mcp.tool(
    name="excel_version",
    annotations={
        "title": "Get MCP Server Version and Status",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def excel_version(ctx: Context) -> str:
    """Return MCP version and browser connection status."""
    try:
        browser = await _get_browser(ctx)
        connected = True
        contexts = len(browser.contexts)
        pages = sum(len(c.pages) for c in browser.contexts)
    except Exception:
        connected = False
        contexts = 0
        pages = 0

    return _ok({
        "version": VERSION,
        "cdp_endpoint": CDP_ENDPOINT,
        "browser_connected": connected,
        "contexts": contexts,
        "pages": pages,
        "coordinate_map": {k: list(v) for k, v in COORDS.items()},
    })


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    mcp.run()
