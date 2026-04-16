# Playwright Pack

Control a real Chromium browser for web automation and scraping.
The browser launches automatically on first tool call — no setup required beyond installation.

---

## Requirements

- `playwright` Python package + Chromium browser binary

---

## Setup

```bash
pip install "openrein-contrib[playwright]"
playwright install chromium
```

---

## Usage

```python
import openrein
import openrein.contrib as contrib

engine = openrein.Engine()
contrib.add_skill(engine, 'playwright')   # loads browser workflow guide
contrib.add_tool(engine, 'playwright')    # registers 15 tools
```

---

## Tools

### Navigation & Reading

| Tool | Description |
|------|-------------|
| `pw_navigate(url)` | Go to a URL — call first before anything else |
| `pw_snapshot(selector?)` | Read page as ARIA accessibility tree — call before every interaction |
| `pw_screenshot(selector?, full_page?)` | Take a screenshot, saved as PNG, returns file path |
| `pw_get_url()` | Get current URL and page title |

### Interaction

| Tool | Description |
|------|-------------|
| `pw_click(selector)` | Click an element |
| `pw_type(selector, text)` | Type into an input field (clears first by default) |
| `pw_press(key)` | Press a keyboard key or shortcut (Enter, Tab, Control+a, ...) |
| `pw_select(selector, value)` | Select option from a `<select>` dropdown |
| `pw_hover(selector)` | Hover to reveal menus or tooltips |

### Wait & Execute

| Tool | Description |
|------|-------------|
| `pw_wait(...)` | Wait for element, text, URL match, or fixed delay |
| `pw_evaluate(script)` | Run JavaScript in the page and return the result |

### Tab Management

| Tool | Description |
|------|-------------|
| `pw_new_page(url?)` | Open a new tab |
| `pw_list_pages()` | List all open tabs with index, URL, title |
| `pw_switch_page(index)` | Switch active tab by index |
| `pw_close_page()` | Close the current tab |

---

## Selectors

```
CSS:          #id  /  .class  /  input[name="email"]
Text:         text=로그인
Role:         role=button[name='Submit']
nth match:    .item >> nth=2
```

Prefer `text=` and `role=` — they survive CSS refactors better than class selectors.

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PW_HEADLESS` | `false` | `true` = no visible window |
| `PW_SCREENSHOT_DIR` | `~/.openrein/screenshots` | Where screenshots are saved |
| `PW_LOCALE` | `ko-KR` | Browser locale |
| `PW_SNAPSHOT_MAX_CHARS` | `8000` | Max characters from `pw_snapshot` |

---

## Smoke Test

```bash
# structure + browser launch + LLM tool use check (headless, no API needed)
python tests/smoke_playwright.py

# interactive — browser window opens, LLM performs the task
python tests/smoke_playwright.py --make "네이버에서 파이썬 튜토리얼 검색해줘"
```

---

## Workflow Example

```python
# Agent interaction pattern
pw_navigate("https://www.naver.com")
pw_snapshot()                              # read page → find search box selector
pw_type("input#query", "파이썬 튜토리얼")
pw_press("Enter")
pw_wait(selector=".search_result")         # wait for results
pw_snapshot()                              # read results
pw_screenshot()                            # visual confirm
```
