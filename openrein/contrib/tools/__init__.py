"""
openrein.contrib.tools — Tool 팩토리 모음.

각 애플리케이션별로 독립된 파일을 유지한다:
  powerpoint.py  → create_powerpoint_tools()
  playwright.py  → create_playwright_tools()
  word.py        → create_word_tools()       (예정)
  excel.py       → create_excel_tools()      (예정)
  hwp2024.py     → create_hwp_tools()        (예정)

사용 예 (dressage agent.py):
  from openrein.contrib.tools import create_powerpoint_tools, create_playwright_tools
  Pack(name="powerpoint", tools=create_powerpoint_tools(), skills=["powerpoint"])
  Pack(name="browser",    tools=create_playwright_tools(), skills=["playwright"])
"""
from __future__ import annotations

# PowerPoint
try:
    from openrein.contrib.tools.powerpoint import create_powerpoint_tools  # noqa: F401
except ImportError:
    def create_powerpoint_tools():  # type: ignore
        return []

# Playwright
try:
    from openrein.contrib.tools.playwright import create_playwright_tools  # noqa: F401
except ImportError:
    def create_playwright_tools():  # type: ignore
        return []

# Word (예정)
# from openrein.contrib.tools.word import create_word_tools

# Excel (예정)
# from openrein.contrib.tools.excel import create_excel_tools

# 한글 2024 (예정)
# from openrein.contrib.tools.hwp2024 import create_hwp_tools
