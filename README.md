# openrein-contrib

Community Skills and Tools for [openrein](https://github.com/kurt01124/openrein).

## Packs

| Pack | Docs | Description |
|------|------|-------------|
| PowerPoint | [docs/powerpoint.md](docs/powerpoint.md) | Office Add-in bridge — create and edit slides via OOXML |
| Playwright | [docs/playwright.md](docs/playwright.md) | Chromium browser automation — navigate, scrape, interact |

---

## Installation

```bash
pip install openrein openrein-contrib

# Playwright (optional)
pip install "openrein-contrib[playwright]"
playwright install chromium
```

---

## Quick Start

```python
import openrein
import openrein.contrib as contrib

engine = openrein.Engine()
contrib.add_skill(engine, 'playwright')
contrib.add_tool(engine, 'playwright')

# See what's available
print(contrib.list_skills())   # ['playwright', 'powerpoint']
print(contrib.list_tools())    # ['playwright', 'powerpoint']
```

---

## Writing a New Pack

**Tool** — `openrein/contrib/tools/<name>.py`:

```python
import openrein

class _Tool(openrein.ToolBase):
    def name(self) -> str:          return "my_tool"
    def description(self) -> str:   return "Does something useful."
    def input_schema(self) -> dict:
        return {"type": "object", "properties": {"input": {"type": "string"}}, "required": ["input"]}
    def call(self, input: dict) -> str:
        return f"Result: {input['input']}"

def create_my_tools() -> list[openrein.ToolBase]:
    return [_Tool()]
```

**Skill** — `openrein/contrib/skills/<name>.md`:

```markdown
---
description: One-line description
when_to_use: When should the agent use this
---

# Skill Title

Instructions and workflow patterns...
```

Register both in `openrein/contrib/tools/__init__.py` and `openrein/contrib/__init__.py`.

---

## License

MIT
