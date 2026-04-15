# openrein-contrib

Community Skills and Tools for [openrein](https://github.com/kurt01124/openrein).

## Installation

```bash
pip install openrein openrein-contrib
```

## Usage

```python
import openrein
from openrein.contrib import SKILLS_DIR
from openrein.contrib.tools import SomeTool

engine = openrein.Engine(system_prompt="You are a helpful assistant.")

# Load all contrib skills
engine.skill_add(SKILLS_DIR)

# Register a contrib tool
engine.register_tool(SomeTool())
```

## Contents

### Skills (`openrein.contrib.skills/`)

Reusable prompt recipes as Markdown files.
Place them in your engine with `engine.skill_add(SKILLS_DIR)`.

| Skill | Description |
|-------|-------------|
| *(coming soon)* | |

### Tools (`openrein.contrib.tools`)

`ToolBase` implementations ready to register with `engine.register_tool()`.

| Tool | Description | Extra dependency |
|------|-------------|-----------------|
| *(coming soon)* | | |

## Contributing

1. **Skill** — add a `.md` file under `openrein/contrib/skills/`
2. **Tool** — add a `.py` file under `openrein/contrib/tools/`, subclass `openrein.ToolBase`

## License

MIT
