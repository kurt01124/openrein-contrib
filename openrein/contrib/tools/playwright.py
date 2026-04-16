"""Playwright tools — headful Chromium browser automation.

첫 번째 툴 호출 시 Chromium 브라우저가 자동으로 실행된다.
headless 여부는 PW_HEADLESS 환경변수로 제어 (기본: false, 창 표시).

의존성 설치:
    pip install playwright
    playwright install chromium
"""
from __future__ import annotations

import json
import os
import threading
import time
from pathlib import Path

import openrein

# ---------------------------------------------------------------------------
# 설정 상수
# ---------------------------------------------------------------------------

_SCREENSHOT_DIR = Path(os.environ.get(
    "PW_SCREENSHOT_DIR",
    str(Path.home() / ".openrein" / "screenshots"),
))
_SNAPSHOT_MAX_CHARS = int(os.environ.get("PW_SNAPSHOT_MAX_CHARS", "8000"))

# ---------------------------------------------------------------------------
# 브라우저 상태 (모듈 수준 싱글톤)
# ---------------------------------------------------------------------------

_playwright = None
_browser = None
_context = None
_pages: list = []
_current_idx: int = 0
_lock = threading.Lock()


def _ensure_browser() -> None:
    """브라우저가 없으면 Chromium을 시작한다."""
    global _playwright, _browser, _context, _pages, _current_idx
    with _lock:
        if _browser is not None:
            return
        try:
            from playwright.sync_api import sync_playwright  # noqa: PLC0415
        except ImportError:
            raise RuntimeError(
                "playwright is not installed.\n"
                "Run: pip install playwright && playwright install chromium"
            )
        _SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
        headless = os.environ.get("PW_HEADLESS", "false").lower() in ("1", "true")
        _playwright = sync_playwright().start()
        _browser = _playwright.chromium.launch(headless=headless)
        _context = _browser.new_context(
            viewport={"width": 1280, "height": 800},
            locale=os.environ.get("PW_LOCALE", "ko-KR"),
        )
        page = _context.new_page()
        _pages.clear()
        _pages.append(page)
        _current_idx = 0


def _page():
    """현재 활성 페이지 반환."""
    _ensure_browser()
    if not _pages:
        raise RuntimeError("No open pages. Use pw_new_page() to open one.")
    return _pages[_current_idx]


def _page_info(page=None) -> dict:
    """URL + 타이틀 딕셔너리 반환."""
    p = page or _page()
    try:
        return {"url": p.url, "title": p.title()}
    except Exception:
        return {"url": getattr(p, "url", ""), "title": ""}


# ---------------------------------------------------------------------------
# 툴 함수 구현
# ---------------------------------------------------------------------------

def _navigate(url: str, wait_until: str = "load") -> dict:
    page = _page()
    try:
        page.goto(url, wait_until=wait_until, timeout=30_000)
        return _page_info(page)
    except Exception as e:
        return {"error": str(e), **_page_info(page)}


def _snapshot(selector: str = ":root", max_chars: int = _SNAPSHOT_MAX_CHARS) -> str:
    page = _page()
    try:
        snap = page.locator(selector).aria_snapshot()
    except Exception:
        try:
            snap = page.evaluate("() => document.body ? document.body.innerText : ''")
        except Exception as e:
            return f"Error: {e}"
    if len(snap) > max_chars:
        snap = snap[:max_chars] + f"\n\n[...TRUNCATED — {len(snap)} chars total]"
    return snap


def _screenshot(selector: str = "", full_page: bool = False) -> dict:
    page = _page()
    ts = int(time.time() * 1000)
    path = str(_SCREENSHOT_DIR / f"pw_{ts}.png")
    try:
        if selector:
            page.locator(selector).first.screenshot(path=path)
        else:
            page.screenshot(path=path, full_page=full_page)
        return {"path": path, **_page_info(page)}
    except Exception as e:
        return {"error": str(e), **_page_info(page)}


def _get_url() -> dict:
    return _page_info()


def _click(
    selector: str,
    button: str = "left",
    double: bool = False,
    timeout_ms: int = 10_000,
) -> dict:
    page = _page()
    locator = page.locator(selector).first
    try:
        if double:
            locator.dblclick(button=button, timeout=timeout_ms)
        else:
            locator.click(button=button, timeout=timeout_ms)
        return {"ok": True, **_page_info(page)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def _type(
    selector: str,
    text: str,
    clear: bool = True,
    submit: bool = False,
    timeout_ms: int = 10_000,
) -> dict:
    page = _page()
    locator = page.locator(selector).first
    try:
        if clear:
            locator.fill(text, timeout=timeout_ms)
        else:
            locator.type(text, timeout=timeout_ms)
        if submit:
            locator.press("Enter", timeout=timeout_ms)
        return {"ok": True, **_page_info(page)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def _press(key: str) -> dict:
    page = _page()
    try:
        page.keyboard.press(key)
        return {"ok": True, "key": key}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def _select(selector: str, value: str, timeout_ms: int = 10_000) -> dict:
    page = _page()
    locator = page.locator(selector).first
    try:
        locator.select_option(value=value, timeout=timeout_ms)
        return {"ok": True}
    except Exception:
        try:
            locator.select_option(label=value, timeout=timeout_ms)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}


def _hover(selector: str, timeout_ms: int = 10_000) -> dict:
    page = _page()
    try:
        page.locator(selector).first.hover(timeout=timeout_ms)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def _wait(
    selector: str = "",
    text: str = "",
    url: str = "",
    ms: int = 0,
    state: str = "visible",
    timeout_ms: int = 30_000,
) -> dict:
    page = _page()
    try:
        if ms > 0:
            page.wait_for_timeout(ms)
        if text:
            page.get_by_text(text).first.wait_for(state="visible", timeout=timeout_ms)
        if selector:
            page.locator(selector).first.wait_for(state=state, timeout=timeout_ms)
        if url:
            page.wait_for_url(url, timeout=timeout_ms)
        return {"ok": True, **_page_info(page)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def _evaluate(script: str) -> dict:
    page = _page()
    try:
        result = page.evaluate(script)
        return {"result": result}
    except Exception as e:
        return {"error": str(e)}


def _new_page(url: str = "") -> dict:
    global _current_idx
    _ensure_browser()
    page = _context.new_page()
    _pages.append(page)
    _current_idx = len(_pages) - 1
    if url:
        try:
            page.goto(url, timeout=30_000)
        except Exception as e:
            return {"index": _current_idx, "error": str(e)}
    return {"index": _current_idx, **_page_info(page)}


def _list_pages() -> dict:
    _ensure_browser()
    result = []
    for i, p in enumerate(_pages):
        try:
            result.append({
                "index": i,
                "url": p.url,
                "title": p.title(),
                "active": i == _current_idx,
            })
        except Exception:
            result.append({"index": i, "url": "?", "title": "?", "active": i == _current_idx})
    return {"pages": result, "current": _current_idx}


def _switch_page(index: int) -> dict:
    global _current_idx
    _ensure_browser()
    if index < 0 or index >= len(_pages):
        return {"ok": False, "error": f"Index {index} out of range (0–{len(_pages) - 1})"}
    _current_idx = index
    return {"ok": True, "index": _current_idx, **_page_info()}


def _close_page() -> dict:
    global _current_idx
    _ensure_browser()
    if not _pages:
        return {"ok": False, "error": "No pages open"}
    page = _pages.pop(_current_idx)
    try:
        page.close()
    except Exception:
        pass
    if _pages:
        _current_idx = min(_current_idx, len(_pages) - 1)
    return {"ok": True, "remaining": len(_pages)}


# ---------------------------------------------------------------------------
# ToolBase 구현
# ---------------------------------------------------------------------------

class _Tool(openrein.ToolBase):
    def __init__(self, name: str, desc: str, schema: dict, fn) -> None:
        self._name = name
        self._desc = desc
        self._schema = schema
        self._fn = fn

    def name(self) -> str:          return self._name
    def description(self) -> str:   return self._desc
    def input_schema(self) -> dict: return self._schema

    def call(self, input: dict) -> str:
        try:
            result = self._fn(**input)
            if isinstance(result, str):
                return result
            return json.dumps(result, ensure_ascii=False, default=str)
        except Exception as e:
            return f"Error: {e}"


# ---------------------------------------------------------------------------
# 팩토리
# ---------------------------------------------------------------------------

def create_playwright_tools() -> list[openrein.ToolBase]:
    """Playwright 브라우저 자동화 툴 리스트 반환."""
    return [
        _Tool(
            name="pw_navigate",
            desc=(
                "Navigate to a URL. Returns final URL and page title. "
                "Always call this first before reading or interacting with a page."
            ),
            schema={
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "Full URL to navigate to (https://...)"},
                    "wait_until": {
                        "type": "string",
                        "enum": ["load", "domcontentloaded", "networkidle", "commit"],
                        "description": "When to consider navigation complete (default: load)",
                    },
                },
                "required": ["url"],
            },
            fn=_navigate,
        ),
        _Tool(
            name="pw_snapshot",
            desc=(
                "Read the current page as an ARIA accessibility tree. "
                "Returns YAML-like text with element roles, names, and hierarchy. "
                "Use this to understand page structure and find selectors before clicking or typing. "
                "Call pw_snapshot before every interaction on an unfamiliar page."
            ),
            schema={
                "type": "object",
                "properties": {
                    "selector": {
                        "type": "string",
                        "description": "CSS selector to snapshot a subtree (default: :root = entire page)",
                    },
                    "max_chars": {
                        "type": "integer",
                        "description": f"Max characters to return (default: {_SNAPSHOT_MAX_CHARS})",
                    },
                },
            },
            fn=_snapshot,
        ),
        _Tool(
            name="pw_screenshot",
            desc=(
                "Take a screenshot of the current page or a specific element. "
                "Saves as PNG and returns the file path. "
                "Use for visual verification after navigation or after interactions."
            ),
            schema={
                "type": "object",
                "properties": {
                    "selector": {
                        "type": "string",
                        "description": "CSS selector for element-only screenshot (omit for full page)",
                    },
                    "full_page": {
                        "type": "boolean",
                        "description": "Capture full scrollable page height (default: false)",
                    },
                },
            },
            fn=_screenshot,
        ),
        _Tool(
            name="pw_get_url",
            desc="Get the current page URL and title.",
            schema={"type": "object", "properties": {}},
            fn=_get_url,
        ),
        _Tool(
            name="pw_click",
            desc=(
                "Click an element. Selector formats: CSS (#id, .class, button[type=submit]), "
                "text content (text=Submit), or role (role=button[name='Submit']). "
                "Call pw_snapshot first to identify the right selector."
            ),
            schema={
                "type": "object",
                "properties": {
                    "selector": {"type": "string", "description": "Element selector"},
                    "button": {
                        "type": "string",
                        "enum": ["left", "right", "middle"],
                        "description": "Mouse button (default: left)",
                    },
                    "double": {"type": "boolean", "description": "Double-click (default: false)"},
                    "timeout_ms": {"type": "integer", "description": "Timeout ms (default: 10000)"},
                },
                "required": ["selector"],
            },
            fn=_click,
        ),
        _Tool(
            name="pw_type",
            desc=(
                "Type text into an input or textarea. "
                "Clears existing content by default (clear=true). "
                "Set submit=true to press Enter after typing."
            ),
            schema={
                "type": "object",
                "properties": {
                    "selector": {"type": "string", "description": "Input element selector"},
                    "text": {"type": "string", "description": "Text to type"},
                    "clear": {
                        "type": "boolean",
                        "description": "Clear existing content first (default: true)",
                    },
                    "submit": {
                        "type": "boolean",
                        "description": "Press Enter after typing (default: false)",
                    },
                    "timeout_ms": {"type": "integer", "description": "Timeout ms (default: 10000)"},
                },
                "required": ["selector", "text"],
            },
            fn=_type,
        ),
        _Tool(
            name="pw_press",
            desc=(
                "Press a keyboard key or shortcut globally. "
                "Examples: 'Enter', 'Tab', 'Escape', 'ArrowDown', "
                "'Control+a', 'Control+c', 'Control+v', 'F5'."
            ),
            schema={
                "type": "object",
                "properties": {
                    "key": {
                        "type": "string",
                        "description": "Key name or combo (e.g. 'Enter', 'Control+a')",
                    },
                },
                "required": ["key"],
            },
            fn=_press,
        ),
        _Tool(
            name="pw_select",
            desc="Select an option from a <select> dropdown by value attribute or visible label text.",
            schema={
                "type": "object",
                "properties": {
                    "selector": {"type": "string", "description": "Select element selector"},
                    "value": {
                        "type": "string",
                        "description": "Option value attribute or visible text label",
                    },
                    "timeout_ms": {"type": "integer", "description": "Timeout ms (default: 10000)"},
                },
                "required": ["selector", "value"],
            },
            fn=_select,
        ),
        _Tool(
            name="pw_hover",
            desc=(
                "Hover the mouse over an element. "
                "Use to reveal dropdown menus, tooltips, or trigger hover states."
            ),
            schema={
                "type": "object",
                "properties": {
                    "selector": {"type": "string", "description": "Element selector"},
                    "timeout_ms": {"type": "integer", "description": "Timeout ms (default: 10000)"},
                },
                "required": ["selector"],
            },
            fn=_hover,
        ),
        _Tool(
            name="pw_wait",
            desc=(
                "Wait for a condition before proceeding. "
                "Specify one or more: element selector, visible text, URL pattern, or fixed delay. "
                "Multiple conditions are checked in sequence."
            ),
            schema={
                "type": "object",
                "properties": {
                    "selector": {"type": "string", "description": "Wait for element to reach given state"},
                    "text": {"type": "string", "description": "Wait for this text to be visible on the page"},
                    "url": {"type": "string", "description": "Wait for URL to match (glob pattern)"},
                    "ms": {"type": "integer", "description": "Fixed delay in milliseconds"},
                    "state": {
                        "type": "string",
                        "enum": ["visible", "hidden", "attached", "detached"],
                        "description": "Target state when waiting for selector (default: visible)",
                    },
                    "timeout_ms": {
                        "type": "integer",
                        "description": "Max total wait time ms (default: 30000)",
                    },
                },
            },
            fn=_wait,
        ),
        _Tool(
            name="pw_evaluate",
            desc=(
                "Execute JavaScript in the browser page context and return the result. "
                "Use for data extraction, reading DOM state, or triggering page effects. "
                "Write as a function: '() => document.title' or '() => [...document.links].map(l => l.href)'"
            ),
            schema={
                "type": "object",
                "properties": {
                    "script": {
                        "type": "string",
                        "description": "JS expression or () => ... function to evaluate in page",
                    },
                },
                "required": ["script"],
            },
            fn=_evaluate,
        ),
        _Tool(
            name="pw_new_page",
            desc="Open a new browser tab and make it the active page. Optionally navigate to a URL.",
            schema={
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "Optional URL to navigate to"},
                },
            },
            fn=_new_page,
        ),
        _Tool(
            name="pw_list_pages",
            desc="List all open browser tabs with index, URL, title, and which is active.",
            schema={"type": "object", "properties": {}},
            fn=_list_pages,
        ),
        _Tool(
            name="pw_switch_page",
            desc="Switch the active tab by its index number (from pw_list_pages).",
            schema={
                "type": "object",
                "properties": {
                    "index": {"type": "integer", "description": "Tab index from pw_list_pages"},
                },
                "required": ["index"],
            },
            fn=_switch_page,
        ),
        _Tool(
            name="pw_close_page",
            desc="Close the current browser tab. Switches to the previous tab if one exists.",
            schema={"type": "object", "properties": {}},
            fn=_close_page,
        ),
    ]
