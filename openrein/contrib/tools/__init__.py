"""
openrein.contrib.tools — Tool 레지스트리.

tags/*.json 에서 "tools": ["BrowserTool"] 처럼 이름으로 참조하면
TOOL_REGISTRY 를 통해 실제 클래스로 매핑된다.

새 Tool 추가 방법:
  1. tools/<name>.py 에 ToolBase 서브클래스 구현
  2. 아래 _register_* 함수 추가
  3. tags/<tag>.json 의 "tools" 에 이름 등록
"""

from __future__ import annotations

TOOL_REGISTRY: dict[str, type] = {}


def _try(name: str, import_path: str, class_name: str) -> None:
    """ImportError 없이 tool 을 레지스트리에 등록."""
    try:
        mod = __import__(import_path, fromlist=[class_name])
        TOOL_REGISTRY[name] = getattr(mod, class_name)
    except ImportError:
        pass


# --- 구현된 Tool 이 생기면 여기에 추가 ---

# _try("BrowserTool", "openrein.contrib.tools.browser", "BrowserTool")
# _try("GitTool",     "openrein.contrib.tools.git",     "GitTool")
# _try("SQLiteTool",  "openrein.contrib.tools.sqlite",  "SQLiteTool")
# _try("OfficeTool",  "openrein.contrib.tools.office",  "OfficeTool")
