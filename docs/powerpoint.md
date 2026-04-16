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

```bash
# Find the manifest path
python -c "from openrein.contrib.tools.powerpoint import _ADDIN_DIR; print(_ADDIN_DIR)"
```

Then in PowerPoint:
- Insert → Add-ins → My Add-ins → `...` → **Upload Add-in**
- Select the `manifest.xml` from the path above

### 3. Open the taskpane

Click the **OpenRein** button in the PowerPoint ribbon.
The taskpane shows "openrein connected" when ready.

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
