"""
openrein.contrib — Community Skills and Tools for openrein.

Usage:
    import openrein
    from openrein.contrib import SKILLS_DIR
    from openrein.contrib.tools import SomeTool

    engine = openrein.Engine(system_prompt="...")
    engine.skill_add(SKILLS_DIR)
    engine.register_tool(SomeTool())
"""

import os

__version__ = "0.1.0"

# Absolute path to the bundled skills directory.
SKILLS_DIR = os.path.join(os.path.dirname(__file__), "skills")
