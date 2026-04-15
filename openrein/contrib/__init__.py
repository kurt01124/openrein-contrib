"""
openrein.contrib — Community Skills and Tools for openrein.

Usage:
    from openrein.contrib import load_tag_defs, SKILLS_DIR

    # dressage 등 에이전트에서 태그 로드:
    for d in load_tag_defs():
        tags.append(Tag(**d))  # Tag는 에이전트 측 dataclass

    # 또는 직접 skill 디렉토리만 사용:
    engine.skill_add(SKILLS_DIR)
"""

import json
import os
from pathlib import Path

__version__ = "0.1.0"

# 번들 디렉토리 경로
SKILLS_DIR = str(Path(__file__).parent / "skills")
TAGS_DIR   = Path(__file__).parent / "tags"


def load_tag_defs() -> list[dict]:
    """tags/*.json 을 읽어 태그 정의 리스트를 반환한다.

    각 항목은 Tag(**d) 로 바로 사용할 수 있는 dict:
      {
        "name":        str,
        "description": str,
        "tools":       list,   # ToolBase 인스턴스 (optional_deps 미충족 시 생략)
        "skills":      list[str],
      }

    optional_deps 가 설치되지 않은 경우 해당 tool 은 조용히 건너뛴다.
    """
    from .tools import TOOL_REGISTRY

    defs = []
    for tag_file in sorted(TAGS_DIR.glob("*.json")):
        raw = json.loads(tag_file.read_text(encoding="utf-8"))

        tools = []
        for tool_name in raw.get("tools", []):
            factory = TOOL_REGISTRY.get(tool_name)
            if factory is None:
                continue          # 아직 구현 안 된 tool
            try:
                tools.append(factory())
            except ImportError:
                pass              # optional_dep 미설치

        defs.append({
            "name":        raw["name"],
            "description": raw["description"],
            "tools":       tools,
            "skills":      raw.get("skills", []),
        })

    return defs
