---
description: Automate and scrape websites using a live Chromium browser via Playwright
when_to_use: 사용자가 웹 브라우저 자동화, 스크래핑, 폼 제출, UI 테스트, 웹 탐색을 요청할 때
---

# Playwright Pack

You control a real Chromium browser. The browser opens automatically on first use.
Read the page with `pw_snapshot` before every interaction — always know what's on screen before touching it.

---

## Tools

| Tool | When to use |
|------|-------------|
| `pw_navigate(url)` | Go to a URL — **always call first** before reading or interacting |
| `pw_snapshot(selector?)` | Read the page as an ARIA tree — **always call before clicking/typing** |
| `pw_screenshot(selector?, full_page?)` | See the page visually — use to confirm results |
| `pw_get_url()` | Get current URL and title |
| `pw_click(selector)` | Click a button, link, or any element |
| `pw_type(selector, text)` | Type into an input field (clears first by default) |
| `pw_press(key)` | Press keyboard key or shortcut (Enter, Tab, Control+a, ...) |
| `pw_select(selector, value)` | Choose option from a `<select>` dropdown |
| `pw_hover(selector)` | Hover to reveal dropdown menus or tooltips |
| `pw_wait(...)` | Wait for element, text, URL, or fixed delay |
| `pw_evaluate(script)` | Run JavaScript in the page and return the result |
| `pw_new_page(url?)` | Open a new tab |
| `pw_list_pages()` | List all open tabs |
| `pw_switch_page(index)` | Switch active tab |
| `pw_close_page()` | Close current tab |

---

## Selectors

Playwright supports multiple selector formats — use whichever matches the page:

```
CSS (most common):
  #submit-btn              ← by id
  .btn-primary             ← by class
  input[name="email"]      ← by attribute
  form button[type=submit] ← combined

Text content:
  text=로그인              ← visible text (exact)
  text=/로그인/i           ← regex match

Role (most robust for buttons/inputs):
  role=button[name='Submit']
  role=textbox[name='이메일']
  role=link[name='더 보기']

nth match (when there are duplicates):
  .item >> nth=2           ← 3rd item (0-indexed)
```

**Rule: prefer `role=` and `text=` selectors.** They survive CSS/class refactors better than `#id` or `.class`.

---

## Core Principles

1. **Read before you touch.** Always call `pw_snapshot()` before `pw_click` or `pw_type` on an unfamiliar page. The snapshot shows you what's actually there.
2. **Navigate first.** Start every task with `pw_navigate(url)`. Do not assume any page is already open.
3. **Verify with screenshots.** After completing a meaningful step (login, form submit, navigation), call `pw_screenshot()` to confirm the result.
4. **Wait after page transitions.** After clicking a button that triggers navigation or loading, call `pw_wait(selector=...)` before reading the new page.
5. **Use `pw_evaluate` for extraction.** For structured data (tables, lists, prices), `pw_evaluate` with JS is faster and more reliable than parsing the snapshot.

---

## Workflow Patterns

### Basic navigation and read

```
1. pw_navigate(url)
2. pw_snapshot()              ← understand page structure
3. pw_screenshot()            ← optional visual check
```

### Filling and submitting a form

```
1. pw_navigate(url)
2. pw_snapshot()              ← find input selectors
3. pw_type(selector, text)    ← fill each field
4. pw_click(submit_selector)  ← submit
5. pw_wait(text="완료")        ← wait for confirmation
6. pw_screenshot()            ← verify result
```

### Login flow

```
1. pw_navigate("https://example.com/login")
2. pw_snapshot()
3. pw_type("input[name=email]", "user@example.com")
4. pw_type("input[name=password]", "password123")
5. pw_click("button[type=submit]")
6. pw_wait(url="**/dashboard**")   ← wait for redirect
7. pw_screenshot()
```

### Scraping data from a page

```
1. pw_navigate(url)
2. pw_wait(selector=".data-table")   ← wait for content to load
3. pw_evaluate("() => [...document.querySelectorAll('tr')].map(r => r.innerText)")
```

### Handling dropdowns / menus

```
# <select> element:
pw_select("select#country", "Korea")

# Custom dropdown (click to open, then click option):
pw_click(".dropdown-trigger")
pw_wait(selector=".dropdown-menu")
pw_click("text=Korea")
```

### Multi-tab workflow

```
1. pw_navigate("https://site-a.com")
2. pw_new_page("https://site-b.com")  ← opens tab index 1
3. pw_switch_page(0)                  ← back to site-a
4. pw_list_pages()                    ← check all tabs
```

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PW_HEADLESS` | `false` | Set to `true` to run without visible browser window |
| `PW_SCREENSHOT_DIR` | `~/.openrein/screenshots` | Where screenshots are saved |
| `PW_LOCALE` | `ko-KR` | Browser locale |
| `PW_SNAPSHOT_MAX_CHARS` | `8000` | Max chars returned by pw_snapshot |

---

## Common Mistakes to Avoid

- **Never interact before pw_snapshot.** You will click the wrong element or miss it entirely.
- **Never hardcode `.nth(0)` blindly.** Verify with snapshot that there is only one match, or use a more specific selector.
- **Don't skip pw_wait after clicks.** Buttons often trigger async loading. Wait for the expected result before reading the new state.
- **Don't use pw_type with clear=false unless you mean to append.** Default clear=true is almost always what you want.
- **For structured data extraction, prefer pw_evaluate over parsing the snapshot.** The snapshot is for navigation; JS is for data.
- **Don't navigate to relative URLs.** Always use full `https://...` URLs.
