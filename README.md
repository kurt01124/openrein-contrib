# openrein-contrib

Community Skills and Tools for [openrein](https://github.com/kurt01124/openrein).

## Packs

| Pack | Docs | Description |
|------|------|-------------|
| PowerPoint | [docs/powerpoint.md](docs/powerpoint.md) | Office Add-in bridge — create and edit slides via OOXML |
| Playwright | [docs/playwright.md](docs/playwright.md) | Chromium browser automation — navigate, scrape, interact |

## Installation

```bash
pip install openrein openrein-contrib
```

## Quick Start

```python
import openrein
import openrein.contrib as contrib

# See what's available
print(contrib.list_skills())   # ['powerpoint', ...]
print(contrib.list_tools())    # ['powerpoint', ...]

# Set up an engine with what you need
engine = openrein.Engine()
contrib.add_skill(engine, 'powerpoint')
contrib.add_tool(engine, 'powerpoint')
```

---

## PowerPoint Add-in

`openrein-contrib` includes a PowerPoint Office Add-in that bridges your Python agent
to PowerPoint via Office.js. The bridge server starts automatically when you import the package.

### Setup

1. **Import the package** — the bridge server starts on `localhost:19876`

   ```python
   import openrein.contrib  # server starts here
   ```

2. **Load the Add-in in PowerPoint**
   - Open PowerPoint
   - Insert → Add-ins → My Add-ins → `...` → Upload Add-in
   - Select: `<site-packages>/openrein_contrib-*.dist-info/../addin/powerpoint/manifest.xml`

   > Or find the path: `python -c "import openrein.contrib.tools.powerpoint as p; print(p._ADDIN_DIR)"`

3. **Open the OpenRein taskpane** in PowerPoint → status shows "openrein connected"

### PowerPoint Pack

```python
import openrein
import openrein.contrib as contrib

engine = openrein.Engine()
contrib.add_skill(engine, 'powerpoint')   # loads PowerPoint workflow guide
contrib.add_tool(engine, 'powerpoint')    # registers 8 PowerPoint tools
```

Available tools:

| Tool | Description |
|------|-------------|
| `ppt_get_context` | Get slide count and document state |
| `ppt_command` | Office.js commands (add_slide, set_background, etc.) |
| `ppt_list_shapes` | List all shapes with id, position, size, text, fill |
| `ppt_edit_xml` | Edit slide OOXML directly via JSZip (primary tool) |
| `ppt_edit_chart` | Insert/replace charts via OOXML |
| `ppt_verify` | Check layout quality (coverage, gaps, overlaps, contrast) |
| `ppt_screenshot` | Take a slide screenshot for visual verification |
| `ppt_set_z_order` | Change shape stacking order |

---

## Writing a New Tool

1. Create `openrein/contrib/tools/<name>.py`

```python
import openrein

class _Tool(openrein.ToolBase):
    def name(self) -> str:          return "my_tool"
    def description(self) -> str:   return "Does something useful."
    def input_schema(self) -> dict:
        return {
            "type": "object",
            "properties": {
                "input": {"type": "string", "description": "Input value"},
            },
            "required": ["input"],
        }
    def call(self, input: dict) -> str:
        return f"Result: {input['input']}"


def create_my_tools() -> list[openrein.ToolBase]:
    return [_Tool()]
```

2. Register in `openrein/contrib/tools/__init__.py`:

```python
from openrein.contrib.tools.my_module import create_my_tools
```

3. Register in `openrein/contrib/__init__.py` `_TOOL_FACTORIES`:

```python
_TOOL_FACTORIES = {
    "powerpoint": "openrein.contrib.tools.powerpoint.create_powerpoint_tools",
    "my_module":  "openrein.contrib.tools.my_module.create_my_tools",
}
```

---

## Writing a New Skill

Add a Markdown file to `openrein/contrib/skills/<name>.md`:

```markdown
---
description: One-line description of what this skill does
when_to_use: When should the agent use this skill
---

# Skill Title

Instructions and workflow patterns for the agent...
```

Skills are automatically discovered by `list_skills()` and loaded by `add_skill()`.

---

## License

MIT
