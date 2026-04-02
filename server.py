#!/usr/bin/env python3
"""
Excel Online Automation Server v1.2026.4

Flask API bridging the HTML dashboard to Chrome via Playwright CDP.
ChatGPT panel opened via frame_locator (text search) — NOT pixel coords.
Resolution-independent: works on any screen size.

Launch Chrome:  chrome --remote-debugging-port=9222 --user-data-dir=%APPDATA%/chrome
Run server:     python server.py
Open:           dashboard.html
"""

import asyncio
import base64
import collections
import json
import os
import time
import threading

from flask import Flask, jsonify, request
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

CDP_ENDPOINT = os.environ.get("CDP_ENDPOINT", "http://localhost:9222")
VERSION = "1.2026.4"

# ---------------------------------------------------------------------------
# Log ring buffer (500 entries) — polled by dashboard every 3s
# ---------------------------------------------------------------------------
_logs = collections.deque(maxlen=500)
_log_counter = 0


def _log(msg, level="info"):
    global _log_counter
    _log_counter += 1
    _logs.append({"id": _log_counter, "ts": time.time(), "level": level, "msg": msg})
    print(f"[{level.upper()}] {msg}", flush=True)

# ---------------------------------------------------------------------------
# Playwright singleton — single background event loop
# ---------------------------------------------------------------------------
_loop = None
_browser = None
_pw = None


def _get_loop():
    global _loop
    if _loop is None or _loop.is_closed():
        _loop = asyncio.new_event_loop()
        threading.Thread(target=_loop.run_forever, daemon=True).start()
    return _loop


def _run(coro):
    """Submit a coroutine to the background loop and block until done."""
    return asyncio.run_coroutine_threadsafe(coro, _get_loop()).result(timeout=180)


async def _connect():
    global _browser, _pw
    if _browser and _browser.is_connected():
        return _browser
    from playwright.async_api import async_playwright
    if _pw is None:
        _pw = await async_playwright().__aenter__()
    _browser = await _pw.chromium.connect_over_cdp(CDP_ENDPOINT, timeout=15000)
    _log(f"Connected to Chrome at {CDP_ENDPOINT}", "success")
    return _browser


async def _page():
    """Return the first Excel Online page, bring it to front."""
    browser = await _connect()
    for bc in browser.contexts:
        for p in bc.pages:
            if any(k in p.url.lower() for k in ("sharepoint", "excel", "office.com", "live.com")):
                await p.bring_to_front()
                return p
    # fallback — return first available page
    if browser.contexts and browser.contexts[0].pages:
        return browser.contexts[0].pages[0]
    raise RuntimeError("No Excel Online page found. Open the workbook in Chrome first.")


async def _any_visible_page():
    """Return the best candidate for navigation.

    Priority: existing Excel/SharePoint tab → any real http/https tab → new tab.
    Skips extension background pages and other non-visible Chrome internals
    (chrome-extension://, devtools://, about:blank, etc.) which are always
    first in the pages list and cause goto() to navigate invisibly.
    """
    browser = await _connect()
    # 1. Prefer an existing Excel/SharePoint tab
    for bc in browser.contexts:
        for p in bc.pages:
            if any(k in p.url.lower() for k in ("sharepoint", "excel", "office.com", "live.com")):
                return p
    # 2. Any real http/https page (skips chrome-extension://, about:blank, devtools://)
    for bc in browser.contexts:
        for p in bc.pages:
            if p.url.startswith("http://") or p.url.startswith("https://"):
                return p
    # 3. Open a fresh tab as last resort
    if browser.contexts:
        return await browser.contexts[0].new_page()
    raise RuntimeError("No browser context available — is Chrome running with --remote-debugging-port=9222?")

# ---------------------------------------------------------------------------
# ChatGPT panel helpers — targets real DOM structure from DevTools inspection
# ---------------------------------------------------------------------------

async def _dismiss_fluent_overlay(page):
    """Disable pointer-events on fluent-default-layer-host overlays in all frames.

    Excel Online renders zoom-warning callouts inside the WacFrame via Fluent UI's
    layer portal (fluent-default-layer-host). These sit on top of ribbon buttons and
    intercept all click events. Removing their pointer-events lets clicks through.
    """
    for frame in page.frames:
        try:
            await frame.evaluate(
                "() => {"
                "  const h = document.getElementById('fluent-default-layer-host');"
                "  if (h) h.style.pointerEvents = 'none';"
                "}"
            )
        except Exception:
            continue


async def _find_and_click_chatgpt(page):
    """Click the ChatGPT ribbon button, clearing Fluent overlay first."""
    await _dismiss_fluent_overlay(page)

    # Strategy 1: aria-label/title on the actual button element — most specific
    try:
        btn = page.frame_locator('#WacFrame_Excel_0').locator(
            '[aria-label*="ChatGPT"], [title*="ChatGPT"]'
        ).first
        await btn.click(force=True, timeout=4000)
        _log("ChatGPT button clicked (WacFrame aria-label, force)", "success")
        return True
    except Exception as e:
        _log(f"ChatGPT S1 failed: {e}", "warn")

    # Strategy 2: text content span inside WacFrame
    try:
        btn = page.frame_locator('#WacFrame_Excel_0').get_by_text('ChatGPT', exact=True).first
        await btn.click(force=True, timeout=4000)
        _log("ChatGPT button clicked (WacFrame text, force)", "success")
        return True
    except Exception as e:
        _log(f"ChatGPT S2 failed: {e}", "warn")

    # Strategy 3: scan every frame
    for frame in page.frames:
        try:
            btn = frame.locator('[aria-label*="ChatGPT"], [title*="ChatGPT"]').first
            await btn.click(force=True, timeout=2000)
            _log(f"ChatGPT button clicked (frame scan: {frame.name or frame.url[:50]})", "success")
            return True
        except Exception:
            continue

    _log("ChatGPT button NOT found in any frame", "error")
    return False


async def _set_mode_smart(page, mode):
    """Set Fast/Standard/Heavy via the reasoning-effort-select button (id="reasoning-effort-select")."""
    label = mode.capitalize()

    # Strategy 1: target the known button id / aria-label in every frame
    for frame in page.frames:
        try:
            btn = frame.locator(
                '[id="reasoning-effort-select"], [aria-label*="Thinking effort"]'
            ).first
            if await btn.count() == 0:
                continue
            await btn.click(force=True, timeout=3000)
            await asyncio.sleep(0.4)
            option = frame.get_by_text(label, exact=True).first
            await option.click(force=True, timeout=3000)
            _log(f"Mode set to {label} (reasoning-effort-select)", "success")
            return True
        except Exception:
            continue

    # Strategy 2: any aria-haspopup="menu" dropdown (generic fallback)
    for frame in page.frames:
        try:
            toggle = frame.locator('[aria-haspopup="menu"], [aria-haspopup="listbox"]').first
            if await toggle.count() == 0:
                continue
            await toggle.click(force=True, timeout=2000)
            await asyncio.sleep(0.4)
            await frame.get_by_text(label, exact=True).first.click(force=True, timeout=2000)
            _log(f"Mode set to {label} (dropdown fallback)", "success")
            return True
        except Exception:
            continue

    _log(f"Could not set mode to {label} — continuing anyway", "warn")
    return False


async def _type_prompt_in_frame(page, prompt):
    """Find the ProseMirror editor in any frame, type the prompt, return the frame."""
    # Strategy 1: look for .ProseMirror (the visible rich-text editor) in all frames
    for frame in page.frames:
        try:
            editor = frame.locator('.ProseMirror').first
            if not await editor.is_visible(timeout=800):
                continue
            await editor.click(timeout=3000)
            await editor.fill("")
            await editor.type(prompt, delay=8)
            _log(f"Prompt typed into ProseMirror ({len(prompt)} chars)", "success")
            return frame
        except Exception:
            continue

    # Strategy 2: visible contenteditable with role=textbox (not aria-hidden)
    for frame in page.frames:
        try:
            editor = frame.locator(
                '[contenteditable="true"][role="textbox"]:not([aria-hidden="true"])'
            ).first
            if not await editor.is_visible(timeout=800):
                continue
            await editor.click(timeout=3000)
            await editor.fill("")
            await editor.type(prompt, delay=8)
            _log(f"Prompt typed into contenteditable ({len(prompt)} chars)", "success")
            return frame
        except Exception:
            continue

    # Strategy 3: keyboard fallback — types at whatever has focus
    _log("Prompt input not found — typing at current focus (fallback)", "warn")
    await page.keyboard.type(prompt, delay=10)
    return None


async def _submit_prompt(page, chatgpt_frame=None):
    """Click the Send message button inside the ChatGPT panel frame."""
    frames = [chatgpt_frame] if chatgpt_frame else page.frames

    # Strategy 1: [aria-label="Send message"] button
    for frame in frames:
        if frame is None:
            continue
        try:
            send = frame.locator('[aria-label="Send message"]').first
            if await send.count() == 0:
                continue
            await send.click(force=True, timeout=3000)
            _log("Prompt submitted (Send button)", "success")
            return
        except Exception:
            continue

    # Strategy 2: Ctrl+Enter (most rich-text editors submit with this)
    try:
        await page.keyboard.press("Control+Return")
        _log("Prompt submitted (Ctrl+Enter)", "info")
        return
    except Exception:
        pass

    # Strategy 3: plain Enter
    await page.keyboard.press("Return")
    _log("Prompt submitted (Enter key)", "info")

# ---------------------------------------------------------------------------
# Core prompt pipeline (shared by /api/chatgpt/prompt and /api/preset/*)
# ---------------------------------------------------------------------------

async def _run_prompt_pipeline(prompt, mode="standard", wait_secs=30):
    """Open panel → set mode → type prompt → submit → wait → screenshot."""
    page = await _page()

    _log("Opening ChatGPT panel...", "info")
    ok = await _find_and_click_chatgpt(page)
    if not ok:
        raise RuntimeError("Could not open ChatGPT panel — button not found")
    await asyncio.sleep(3)

    _log(f"Setting mode: {mode}", "info")
    await _set_mode_smart(page, mode)
    await asyncio.sleep(0.5)

    _log(f"Typing prompt ({len(prompt)} chars)...", "info")
    chatgpt_frame = await _type_prompt_in_frame(page, prompt)
    await asyncio.sleep(0.3)

    _log("Submitting...", "info")
    await _submit_prompt(page, chatgpt_frame)
    await asyncio.sleep(1)

    _log(f"Waiting {wait_secs}s for ChatGPT to respond...", "info")
    await asyncio.sleep(wait_secs)

    img = await page.screenshot(type="png")
    b64 = base64.b64encode(img).decode()
    _log("Done. Screenshot captured.", "success")
    return b64

# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------

@app.route("/api/status")
def api_status():
    try:
        browser = _run(_connect())
        pages = sum(len(c.pages) for c in browser.contexts)
        return jsonify({"connected": True, "pages": pages, "version": VERSION})
    except Exception as e:
        return jsonify({"connected": False, "error": str(e), "version": VERSION})


@app.route("/api/logs")
def api_logs():
    since = int(request.args.get("since", 0))
    entries = [e for e in list(_logs) if e["id"] > since]
    return jsonify({"logs": entries, "latest_id": _log_counter})


@app.route("/api/screenshot")
def api_screenshot():
    try:
        async def _do():
            page = await _page()
            img = await page.screenshot(type="png")
            return base64.b64encode(img).decode()
        b64 = _run(_do())
        _log("Screenshot captured", "info")
        return jsonify({"ok": True, "image": b64})
    except Exception as e:
        _log(f"Screenshot error: {e}", "error")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/chatgpt/open", methods=["POST"])
def api_chatgpt_open():
    try:
        async def _do():
            page = await _page()
            ok = await _find_and_click_chatgpt(page)
            if ok:
                await asyncio.sleep(4)
            return ok
        ok = _run(_do())
        return jsonify({"ok": ok, "error": None if ok else "Button not found"})
    except Exception as e:
        _log(f"Open panel error: {e}", "error")
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/chatgpt/prompt", methods=["POST"])
def api_chatgpt_prompt():
    data = request.json or {}
    prompt = (data.get("prompt") or "").strip()
    mode = (data.get("mode") or "standard").lower()
    wait_secs = int(data.get("wait") or 30)
    if not prompt:
        return jsonify({"ok": False, "error": "prompt required"}), 400
    try:
        b64 = _run(_run_prompt_pipeline(prompt, mode, wait_secs))
        return jsonify({"ok": True, "image": b64})
    except Exception as e:
        _log(f"Prompt error: {e}", "error")
        return jsonify({"ok": False, "error": str(e)}), 500


# ---------------------------------------------------------------------------
# Preset actions
# ---------------------------------------------------------------------------

PRESETS = {
    "format-table": (
        "Format the selected data range as a professional table. "
        "Apply alternating row colours, bold the header row, auto-fit all column widths, "
        "and add a clean border."
    ),
    "add-totals": (
        "Add a Totals row at the bottom of the data. "
        "SUM all numeric columns, label the first cell 'TOTAL', and bold the entire row."
    ),
    "create-chart": (
        "Create the most appropriate chart (bar, line, or pie) from the selected data. "
        "Give it a descriptive title and place it below the table."
    ),
    "clean-data": (
        "Clean and standardise the selected data range: remove leading/trailing spaces, "
        "fix capitalisation, standardise date formats, and flag or remove duplicate rows."
    ),
    "add-formulas": (
        "Analyse the data and add useful formulas — totals, averages, percentages, "
        "or lookups wherever appropriate. Add a brief comment explaining each formula."
    ),
    "conditional-format": (
        "Apply conditional formatting to the data: colour scale for numeric columns, "
        "top 10% highlighted green, bottom 10% red, blanks and errors flagged yellow."
    ),
}


@app.route("/api/preset/<preset_id>", methods=["POST"])
def api_preset(preset_id):
    if preset_id not in PRESETS:
        return jsonify({"ok": False, "error": f"Unknown preset '{preset_id}'"}), 400
    data = request.json or {}
    mode = (data.get("mode") or "standard").lower()
    wait_secs = int(data.get("wait") or 45)
    prompt = PRESETS[preset_id]
    _log(f"Preset: {preset_id}", "info")
    try:
        b64 = _run(_run_prompt_pipeline(prompt, mode, wait_secs))
        return jsonify({"ok": True, "image": b64})
    except Exception as e:
        _log(f"Preset error: {e}", "error")
        return jsonify({"ok": False, "error": str(e)}), 500


# ---------------------------------------------------------------------------
# Workflow persistence (save/load named workflows as JSON files)
# ---------------------------------------------------------------------------

WORKFLOWS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "workflows")
os.makedirs(WORKFLOWS_DIR, exist_ok=True)

# Active runs: { run_id: { status, logs, workflow_name } }
_runs = {}
_run_counter = 0


@app.route("/api/workflows", methods=["GET"])
def api_list_workflows():
    try:
        files = [f[:-5] for f in os.listdir(WORKFLOWS_DIR) if f.endswith(".json")]
        return jsonify({"ok": True, "workflows": sorted(files)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/workflows/<name>", methods=["GET"])
def api_load_workflow(name):
    path = os.path.join(WORKFLOWS_DIR, f"{name}.json")
    if not os.path.exists(path):
        return jsonify({"ok": False, "error": "Not found"}), 404
    with open(path) as f:
        data = json.load(f)
    return jsonify({"ok": True, "workflow": data})


@app.route("/api/workflows/<name>", methods=["POST"])
def api_save_workflow(name):
    data = request.json or {}
    path = os.path.join(WORKFLOWS_DIR, f"{name}.json")
    with open(path, "w") as f:
        json.dump(data, f, indent=2)
    _log(f"Workflow saved: {name}", "success")
    return jsonify({"ok": True})


@app.route("/api/workflows/<name>", methods=["DELETE"])
def api_delete_workflow(name):
    path = os.path.join(WORKFLOWS_DIR, f"{name}.json")
    if os.path.exists(path):
        os.remove(path)
    return jsonify({"ok": True})


@app.route("/api/runs", methods=["GET"])
def api_list_runs():
    return jsonify({"runs": [
        {"id": k, "status": v["status"], "name": v["name"], "started": v["started"]}
        for k, v in _runs.items()
    ]})


@app.route("/api/runs/<run_id>", methods=["GET"])
def api_run_status(run_id):
    run = _runs.get(run_id)
    if not run:
        return jsonify({"ok": False, "error": "Run not found"}), 404
    return jsonify({"ok": True, **run})


@app.route("/api/runs/<run_id>/cancel", methods=["POST"])
def api_cancel_run(run_id):
    if run_id in _runs:
        _runs[run_id]["status"] = "cancelled"
        _log(f"Run {run_id} cancelled", "warn")
    return jsonify({"ok": True})


@app.route("/api/run", methods=["POST"])
def api_run_workflow():
    global _run_counter
    data = request.json or {}
    workflow = data.get("workflow", {})
    _run_counter += 1
    run_id = f"run_{_run_counter:04d}"
    run_name = workflow.get("name", "UNNAMED_OP")
    _runs[run_id] = {
        "id": run_id, "status": "running", "name": run_name,
        "started": time.time(), "logs": [], "current_node": None
    }

    def _add_run_log(msg, level="info"):
        _runs[run_id]["logs"].append({"ts": time.time(), "level": level, "msg": msg})
        _log(f"[{run_id}] {msg}", level)

    async def _execute():
        try:
            nodes = workflow.get("nodes", {})
            connections = workflow.get("connections", [])
            _add_run_log(f"Deploying operation: {run_name}", "info")

            # Build execution order (topological from START node)
            start_nodes = [n for n in nodes.values() if n.get("type") == "start"]
            if not start_nodes:
                raise RuntimeError("No START block found — add a START block to your workflow")

            # Simple linear execution: follow connections from start
            order = []
            current_id = start_nodes[0]["id"]
            visited = set()
            while current_id and current_id not in visited:
                visited.add(current_id)
                order.append(current_id)
                next_conn = next((c for c in connections if c["from"] == current_id), None)
                current_id = next_conn["to"] if next_conn else None

            _add_run_log(f"Execution chain: {len(order)} blocks", "info")

            async def _exec_node(ntype, cfg):
                """Execute a single node. loop is handled by run_chain, not here."""
                if ntype == "start":
                    _add_run_log("Operation initiated", "success")

                elif ntype == "navigate_url":
                    url = cfg.get("url", "").strip()
                    if not url:
                        _add_run_log("Navigate URL: no URL configured", "warn")
                    else:
                        try:
                            page = await _any_visible_page()
                            _add_run_log(f"Navigate URL: using tab at {page.url[:60] or 'blank'}", "info")
                            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
                            await page.bring_to_front()
                            _add_run_log(f"Navigated to: {url}", "success")
                        except Exception as e:
                            _add_run_log(f"Navigate URL error: {e}", "error")

                elif ntype == "chatgpt_prompt":
                    prompt = cfg.get("prompt", "").strip()
                    mode = cfg.get("mode", "standard")
                    wait = int(cfg.get("wait", 30))
                    if not prompt:
                        _add_run_log("ChatGPT Prompt: no prompt configured — skipping", "warn")
                    else:
                        b64 = await _run_prompt_pipeline(prompt, mode, wait)
                        _runs[run_id]["last_screenshot"] = b64
                        _add_run_log("ChatGPT prompt executed", "success")

                elif ntype == "open_panel":
                    page = await _page()
                    ok = await _find_and_click_chatgpt(page)
                    if ok:
                        await asyncio.sleep(3)
                        _add_run_log("ChatGPT panel opened", "success")
                    else:
                        _add_run_log("Could not open panel", "error")

                elif ntype == "screenshot":
                    page = await _page()
                    img = await page.screenshot(type="png")
                    _runs[run_id]["last_screenshot"] = base64.b64encode(img).decode()
                    _add_run_log("Screenshot captured", "success")

                elif ntype == "wait":
                    secs = int(cfg.get("seconds", 5))
                    _add_run_log(f"Waiting {secs}s...", "info")
                    await asyncio.sleep(secs)

                elif ntype == "read_file":
                    path = cfg.get("path", "").strip()
                    if not path:
                        _add_run_log("Read File: no path configured", "warn")
                    else:
                        try:
                            with open(path, "r", encoding="utf-8", errors="replace") as fh:
                                content = fh.read()
                            _runs[run_id]["file_content"] = content
                            _add_run_log(
                                f"Read file: {os.path.basename(path)} ({len(content)} chars)",
                                "success"
                            )
                        except Exception as e:
                            _add_run_log(f"Read File error: {e}", "error")

                elif ntype == "upload_file":
                    path = cfg.get("path", "").strip()
                    if not path:
                        _add_run_log("Upload File: no path configured", "warn")
                    elif not os.path.exists(path):
                        _add_run_log(f"Upload File: file not found — {path}", "error")
                    else:
                        try:
                            page = await _page()
                            uploaded = False
                            # Strategy 1: file input inside WacFrame
                            try:
                                fi = page.frame_locator('#WacFrame_Excel_0').locator(
                                    'input[type="file"]'
                                ).first
                                await fi.set_input_files(path, timeout=5000)
                                _add_run_log(
                                    f"File uploaded (WacFrame): {os.path.basename(path)}",
                                    "success"
                                )
                                uploaded = True
                            except Exception:
                                pass
                            # Strategy 2: file input on main page
                            if not uploaded:
                                fi = page.locator('input[type="file"]').first
                                await fi.set_input_files(path, timeout=5000)
                                _add_run_log(
                                    f"File uploaded: {os.path.basename(path)}", "success"
                                )
                                uploaded = True
                            if not uploaded:
                                _add_run_log("Upload File: no file input found in page", "error")
                        except Exception as e:
                            _add_run_log(f"Upload File error: {e}", "error")

                elif ntype == "log_message":
                    msg = cfg.get("msg", "").strip()
                    _add_run_log(f"[LOG] {msg}" if msg else "[LOG] (no message configured)", "info")

                elif ntype == "save_result":
                    filename = (cfg.get("filename") or "result").strip()
                    results_dir = os.path.join(
                        os.path.dirname(os.path.abspath(__file__)), "results"
                    )
                    os.makedirs(results_dir, exist_ok=True)
                    screenshot_b64 = _runs[run_id].get("last_screenshot")
                    file_content   = _runs[run_id].get("file_content")
                    if screenshot_b64:
                        if not any(filename.lower().endswith(e) for e in (".png", ".jpg", ".jpeg")):
                            out_path = os.path.join(results_dir, filename + ".png")
                        else:
                            out_path = os.path.join(results_dir, filename)
                        with open(out_path, "wb") as fh:
                            fh.write(base64.b64decode(screenshot_b64))
                        _add_run_log(f"Screenshot saved → {out_path}", "success")
                    elif file_content is not None:
                        out_path = os.path.join(results_dir, filename)
                        with open(out_path, "w", encoding="utf-8") as fh:
                            fh.write(file_content)
                        _add_run_log(f"File content saved → {out_path}", "success")
                    else:
                        _add_run_log(
                            "Save Result: nothing to save — run Screenshot or Read File first",
                            "warn"
                        )

                elif ntype == "mcp_tool":
                    tool = cfg.get("tool", "").strip()
                    _add_run_log(
                        f"MCP tool: {tool or '(none)'} — stub, wire up your MCP here", "warn"
                    )

                elif ntype == "condition":
                    expr = cfg.get("expr", "").strip()
                    _add_run_log(
                        f"Condition '{expr}': branching not supported in linear executor"
                        " — continuing on first connection",
                        "warn"
                    )

                else:
                    _add_run_log(f"Unknown block type: {ntype!r} — skipping", "warn")

            async def run_chain(chain):
                for i, node_id in enumerate(chain):
                    if _runs[run_id]["status"] == "cancelled":
                        break
                    node = nodes.get(node_id)
                    if not node:
                        continue
                    ntype = node.get("type")
                    cfg   = node.get("config", {})
                    _runs[run_id]["current_node"] = node_id
                    _add_run_log(f"Executing: {ntype}", "info")

                    if ntype == "loop":
                        count = int(cfg.get("count", 3))
                        remaining = chain[i + 1:]
                        if not remaining:
                            _add_run_log("Loop: no blocks after loop node — nothing to repeat", "warn")
                        else:
                            _add_run_log(f"Loop: {count} × {len(remaining)} block(s)", "info")
                            for iteration in range(count):
                                if _runs[run_id]["status"] == "cancelled":
                                    break
                                _add_run_log(f"Loop iteration {iteration + 1}/{count}", "info")
                                await run_chain(remaining)
                        return  # loop node consumed the rest of the chain
                    else:
                        await _exec_node(ntype, cfg)

            await run_chain(order)

            _runs[run_id]["status"] = "done"
            _add_run_log("Operation complete", "success")

        except Exception as e:
            _runs[run_id]["status"] = "error"
            _add_run_log(f"Error: {e}", "error")

    fut = asyncio.run_coroutine_threadsafe(_execute(), _get_loop())
    threading.Thread(target=lambda: fut.result(), daemon=True).start()
    return jsonify({"ok": True, "run_id": run_id})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    _log(f"Excel Online Automation Server {VERSION} starting...", "info")
    _log(f"CDP endpoint: {CDP_ENDPOINT}", "info")
    _log("Dashboard: open dashboard.html in your browser", "info")
    _log("Connecting to Chrome in background...", "info")

    # Pre-connect on startup so first request is instant
    try:
        _run(_connect())
    except Exception as e:
        _log(f"Chrome not reachable at startup: {e} — will retry on first request", "warn")

    app.run(host="0.0.0.0", port=5111, debug=False, threaded=True)
