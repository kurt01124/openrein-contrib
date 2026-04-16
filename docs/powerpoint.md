# PowerPoint Pack

Control PowerPoint through a live Office Add-in bridge.
An embedded HTTP server starts automatically on first import — no separate process needed.

---

## Requirements

- Windows with Microsoft PowerPoint installed
- Office Add-in loaded in PowerPoint (one-time setup)

---

## Setup

### 1. Install

```bash
pip install openrein openrein-contrib
```

### 2. Load the Add-in in PowerPoint

Find the manifest path:

```bash
python -c "from openrein.contrib.tools.powerpoint import _ADDIN_DIR; print(_ADDIN_DIR)"
```

Then in PowerPoint:

1. **Insert** → **Add-ins** → **My Add-ins** → `···` → **Upload Add-in**
2. Select `manifest.xml` from the path printed above
3. Click **OpenRein** in the Home ribbon to open the taskpane

The taskpane status dot turns green and shows **"openrein connected"** once the bridge is live.

> **Note:** The Add-in must be loaded every time PowerPoint opens.
> Pin it via **Insert → Add-ins → My Add-ins** so it appears in the ribbon automatically.

---

## Add-in

The Add-in is a standard Office Web Add-in (taskpane) that runs inside PowerPoint.

### Files

| File | Description |
|------|-------------|
| `addin/powerpoint/manifest.xml` | Add-in manifest — tells PowerPoint where to load the taskpane from |
| `addin/powerpoint/taskpane.html` | Taskpane UI shell |
| `addin/powerpoint/taskpane.js` | Bridge logic: polls for commands, executes via Office.js + JSZip |
| `addin/powerpoint/taskpane.css` | Taskpane styles |

### Bridge Server

When you `import openrein.contrib.tools.powerpoint`, an embedded HTTP server starts on `localhost:19876`:

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/taskpane.html` | GET | Serves the Add-in UI |
| `/api/poll` | GET | Add-in polls here for pending commands |
| `/api/result` | POST | Add-in posts command results here |
| `/api/report_context` | POST | Add-in reports current slide context |
| `/api/status` | GET | Returns `{ connected, queue }` |
| `/api/command` | POST | External HTTP trigger (smoke tests, scripts) |

### Supported Actions (taskpane.js)

| Action | Description |
|--------|-------------|
| `get_office_context` | Slide count, active slide index |
| `office_command` | add_slide, delete_slide, add_text_box, add_shape, set_background, get_layouts, apply_layout, … |
| `list_slide_shapes` | All shapes with id, type, position, size, text, font, fill |
| `edit_slide_xml` | Full slide OOXML replacement via JSZip |
| `edit_slide_chart` | Chart OOXML via JSZip (3-file pattern) |
| `verify_slides` | Layout quality check (coverage, gaps, overlaps, contrast) |
| `verify_slide_visual` | Screenshot of a specific slide (base64 PNG) |
| `set_z_order` | Shape z-order via `Office.ZOrderType` |

---

## Usage

```python
import openrein
import openrein.contrib as contrib

engine = openrein.Engine()
contrib.add_skill(engine, 'powerpoint')   # loads workflow guide
contrib.add_tool(engine, 'powerpoint')    # registers 8 tools
```

---

## Tools

| Tool | Description |
|------|-------------|
| `ppt_get_context` | Get slide count and document state — call before anything else |
| `ppt_list_shapes` | List all shapes with id, position, size, text, fill |
| `ppt_edit_xml` | Edit slide OOXML directly via JSZip — primary layout tool |
| `ppt_edit_chart` | Insert/replace charts via OOXML (3-file pattern) |
| `ppt_command` | Office.js commands: add_slide, delete_slide, set_background, etc. |
| `ppt_verify` | Check layout quality: coverage, gaps, overlaps, contrast |
| `ppt_screenshot` | Take a slide screenshot for visual verification |
| `ppt_set_z_order` | Change shape stacking order |

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PPT_PORT` | `19876` | Bridge server port |

---

## Smoke Test

```bash
# structure + server + LLM tool use check
python tests/smoke_powerpoint.py

# actually build something (requires Add-in connected)
python tests/smoke_powerpoint.py --make "파이썬 소개 3슬라이드 프레젠테이션 만들어줘"

# vision model variant (glm-4.6v-flash sees screenshots directly)
python tests/smoke_powerpoint-4.6v.py --make "..."
```

---

## How It Works

```
Python (openrein engine)
    └─► ppt_edit_xml(slide, code)
            └─► HTTP POST → localhost:19876/api/command
                    └─► PowerPoint Add-in (taskpane.js)
                            └─► JSZip → OOXML edit → insertSlidesFromBase64
```

1. The tool pushes a command to an internal queue
2. The Add-in polls `/api/poll` every ~300ms and picks up the command
3. The Add-in executes it via Office.js and posts the result to `/api/result`
4. The tool returns the result to the engine (blocking, 60s timeout)
