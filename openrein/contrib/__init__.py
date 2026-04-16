"""
openrein.contrib — Community Skills and Tools for openrein.

Quick start:
    from openrein.contrib import list_skills, get_skill, list_tools, get_tools

    print(list_skills())          # ['powerpoint', ...]
    print(list_tools())           # ['powerpoint', ...]

    skill = get_skill('powerpoint')   # skill prompt string
    tools = get_tools('powerpoint')   # list[ToolBase]

Or use with an openrein Engine directly:
    import openrein
    from openrein.contrib import SKILLS_DIR
    from openrein.contrib.tools import create_powerpoint_tools

    engine = openrein.Engine()
    engine.skill_add(SKILLS_DIR)
    for tool in create_powerpoint_tools():
        engine.register_tool(tool)
"""

from pathlib import Path

__version__ = "0.6.2"

# Skills directory (bundled with this package)
SKILLS_DIR = str(Path(__file__).parent / "skills")


# ---------------------------------------------------------------------------
# Skills
# ---------------------------------------------------------------------------

def list_skills() -> list[str]:
    """Available skill names (without .md extension).

    Example:
        >>> list_skills()
        ['powerpoint']
    """
    return sorted(p.stem for p in Path(SKILLS_DIR).glob("*.md"))


def get_skill(name: str) -> str:
    """Return the prompt body of a skill by name.

    Frontmatter (---...---) is stripped; only the body is returned.

    Args:
        name: Skill name without .md extension (e.g. 'powerpoint')

    Returns:
        Skill prompt as a string.

    Raises:
        KeyError: If the skill is not found.

    Example:
        >>> print(get_skill('powerpoint')[:80])
    """
    path = Path(SKILLS_DIR) / f"{name}.md"
    if not path.exists():
        available = list_skills()
        raise KeyError(f"Skill '{name}' not found. Available: {available}")

    raw = path.read_text(encoding="utf-8")
    lines = raw.split("\n")
    if lines and lines[0].strip() == "---":
        end = next((i for i in range(1, len(lines)) if lines[i].strip() == "---"), -1)
        if end != -1:
            return "\n".join(lines[end + 1:]).lstrip("\n")
    return raw


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

_TOOL_FACTORIES: dict[str, str] = {
    "powerpoint": "openrein.contrib.tools.powerpoint.create_powerpoint_tools",
    "playwright": "openrein.contrib.tools.playwright.create_playwright_tools",
    # "word":       "openrein.contrib.tools.word.create_word_tools",
    # "excel":      "openrein.contrib.tools.excel.create_excel_tools",
    # "hwp2024":    "openrein.contrib.tools.hwp2024.create_hwp_tools",
}


def list_tools() -> list[str]:
    """Available tool pack names.

    Example:
        >>> list_tools()
        ['powerpoint']
    """
    return sorted(_TOOL_FACTORIES.keys())


def add_skill(engine, name: str) -> None:
    """Load a single skill by name and add it to the engine (on-demand).

    list_skills() first to see what's available.

    Example:
        skills = list_skills()   # ['powerpoint', ...]  — cheap, just filenames
        add_skill(engine, 'powerpoint')                  # loads .md now
    """
    path = Path(SKILLS_DIR) / f"{name}.md"
    if not path.exists():
        raise KeyError(f"Skill '{name}' not found. Available: {list_skills()}")
    engine.skill_add(str(path))


def add_tool(engine, name: str) -> None:
    """Instantiate a tool pack by name and register it with the engine (on-demand).

    list_tools() first to see what's available.

    Example:
        tools = list_tools()     # ['powerpoint', ...]  — cheap, just names
        add_tool(engine, 'powerpoint')                   # imports + registers now
    """
    for tool in get_tools(name):
        engine.register_tool(tool)


def get_tools(name: str) -> list:
    """Return a list of ToolBase instances for the given tool pack.

    Args:
        name: Tool pack name (e.g. 'powerpoint')

    Returns:
        list[openrein.ToolBase]

    Raises:
        KeyError: If the tool pack is not found.

    Example:
        >>> tools = get_tools('powerpoint')
        >>> engine.register_tool(t) for t in tools
    """
    if name not in _TOOL_FACTORIES:
        raise KeyError(f"Tool pack '{name}' not found. Available: {list_tools()}")

    module_path, func_name = _TOOL_FACTORIES[name].rsplit(".", 1)
    mod = __import__(module_path, fromlist=[func_name])
    return getattr(mod, func_name)()
