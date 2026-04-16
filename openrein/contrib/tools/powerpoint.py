"""PowerPoint tools — Office Add-in bridge with embedded HTTP server.

addin_server를 별도로 실행할 필요 없이 툴을 처음 호출하는 순간
백그라운드 스레드로 서버가 자동 시작된다.

서버는 openrein/contrib/addin/powerpoint/ 디렉토리의 파일을 서빙하며
taskpane.js ↔ Python 간 명령 브릿지 역할을 한다.
"""
from __future__ import annotations

import json
import os
import threading
import time
import urllib.error
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import openrein

# ---------------------------------------------------------------------------
# 경로 상수
# ---------------------------------------------------------------------------

# addin/ 은 패키지 외부 (레포 루트)에 있음 — python 패키지에 포함하지 않는 설계
# openrein/contrib/tools/powerpoint.py → ../../../../addin/powerpoint/
_ADDIN_DIR = str(Path(__file__).parent.parent.parent.parent / "addin" / "powerpoint")
_PORT = 19876  # static 파일 + API 브릿지 통합 서버

# ---------------------------------------------------------------------------
# 내장 addin 서버 (백그라운드 스레드)
# ---------------------------------------------------------------------------

_command_queue:   list[dict] = []
_queue_lock                  = threading.Lock()
_command_counter             = 0
_last_poll_time: float       = 0
_addin_context:  dict        = {}
_server_instance             = None
_server_start_lock           = threading.Lock()


def _push_command(command_data: dict) -> str:
    """명령을 큐에 넣고 애드인 결과가 올 때까지 블로킹 (최대 60초)."""
    global _command_counter
    with _queue_lock:
        _command_counter += 1
        cmd_id = _command_counter
        entry  = {"id": cmd_id, "command": command_data, "status": "pending", "result": None}
        _command_queue.append(entry)

    deadline = time.time() + 60
    while time.time() < deadline:
        with _queue_lock:
            for e in _command_queue:
                if e["id"] == cmd_id and e["status"] == "done":
                    result = e["result"]
                    _command_queue.remove(e)
                    return result
        time.sleep(0.3)

    with _queue_lock:
        _command_queue[:] = [e for e in _command_queue if e["id"] != cmd_id]
    return "Error: Add-in did not respond within 60 seconds. Is PowerPoint open with the Add-in loaded?"


def _is_connected() -> bool:
    """애드인이 최근 10초 이내에 폴링했으면 True."""
    return bool(_last_poll_time) and (time.time() - _last_poll_time) < 10


class _Handler(SimpleHTTPRequestHandler):
    """addin_server 브릿지 + 정적 파일 서빙."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=_ADDIN_DIR, **kwargs)

    def do_GET(self):
        if self.path == "/api/poll":
            self._serve_poll()
        elif self.path == "/api/status":
            self._json(200, {"connected": _is_connected(), "queue": len(_command_queue)})
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == "/api/result":
            self._handle_result()
        elif self.path == "/api/report_context":
            self._handle_context()
        elif self.path == "/api/command":
            self._handle_command()
        else:
            self._json(404, {"error": "Not found"})

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    # ------------------------------------------------------------------
    # 핸들러 구현
    # ------------------------------------------------------------------

    def _serve_poll(self):
        global _last_poll_time
        _last_poll_time = time.time()
        with _queue_lock:
            pending = [e for e in _command_queue if e["status"] == "pending"]
            if pending:
                cmd = pending[0]
                cmd["status"] = "executing"
                self._json(200, {"id": cmd["id"], "command": cmd["command"]})
                return
        self._json(200, {"id": None, "command": None})

    def _handle_result(self):
        try:
            body = self._read_body()
            with _queue_lock:
                for e in _command_queue:
                    if e["id"] == body.get("id"):
                        e["result"] = body.get("result", "")
                        e["status"] = "done"
                        break
            self._json(200, {"ok": True})
        except Exception as ex:
            self._json(400, {"error": str(ex)})

    def _handle_context(self):
        global _addin_context
        try:
            _addin_context = self._read_body()
            self._json(200, {"ok": True})
        except Exception as ex:
            self._json(400, {"error": str(ex)})

    def _handle_command(self):
        """외부 프로세스(smoke_powerpoint.py 등)에서 HTTP로 명령 전송 시 처리."""
        try:
            body   = self._read_body()
            action = body.get("action") or body.get("name", "")
            params = body.get("params", {})
            if not action:
                self._json(400, {"error": "Missing action"})
                return
            if not _is_connected():
                self._json(503, {"error": "PowerPoint Add-in not connected."})
                return
            result = _push_command({"action": action, "params": params})
            self._json(200, {"result": result})
        except Exception as ex:
            self._json(500, {"error": str(ex)})

    # ------------------------------------------------------------------
    # 유틸
    # ------------------------------------------------------------------

    def _read_body(self) -> dict:
        length = int(self.headers.get("Content-Length", 0))
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _json(self, code: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", len(body))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass  # 접속 로그 숨김


def _start_server() -> None:
    """서버 시작 (static + API 통합, 포트 19876).
    모듈 임포트 즉시 호출되므로 Python이 살아있는 한 add-in UI가 항상 로드됨."""
    global _server_instance
    with _server_start_lock:
        if _server_instance is not None:
            return
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", _PORT), _Handler)
            t   = threading.Thread(target=srv.serve_forever, daemon=True)
            t.start()
            _server_instance = srv
        except OSError:
            pass  # 이미 다른 프로세스가 포트를 점유 중


# 모듈 임포트 즉시 서버 시작
_start_server()

# ---------------------------------------------------------------------------
# 툴 공통 호출 함수
# ---------------------------------------------------------------------------

def _call(action: str, params: dict) -> str:
    if not _is_connected():
        return (
            f"Error: PowerPoint Add-in is not connected.\n"
            f"manifest: {_ADDIN_DIR}/manifest.xml"
        )
    result = _push_command({"action": action, "params": params})
    return result if isinstance(result, str) else json.dumps(result, ensure_ascii=False)


# ---------------------------------------------------------------------------
# ToolBase 구현
# ---------------------------------------------------------------------------

class _Tool(openrein.ToolBase):
    def __init__(self, action: str, tool_name: str, desc: str, schema: dict) -> None:
        self._action = action
        self._tname  = tool_name
        self._desc   = desc
        self._schema = schema

    def name(self) -> str:          return self._tname
    def description(self) -> str:   return self._desc
    def input_schema(self) -> dict: return self._schema

    def call(self, input: dict) -> str:
        return _call(self._action, input)


def create_powerpoint_tools() -> list[openrein.ToolBase]:
    """PowerPoint 팩용 ToolBase 인스턴스 리스트 반환. 첫 call() 시 서버 자동 시작."""
    return [
        _Tool(
            action="get_office_context",
            tool_name="ppt_get_context",
            desc=(
                "Get the current PowerPoint document state: slide count, active slide. "
                "Call before making changes to understand what's open."
            ),
            schema={"type": "object", "properties": {}},
        ),
        _Tool(
            action="office_command",
            tool_name="ppt_command",
            desc=(
                "Send an Office.js command to PowerPoint. "
                "Common actions: add_slide, delete_slide, add_text_box, add_shape, "
                "set_background, get_slide_count, get_layouts, apply_layout."
            ),
            schema={
                "type": "object",
                "properties": {
                    "action": {"type": "string", "description": "Action name (e.g. 'add_slide', 'set_background')"},
                    "params": {"type": "object", "description": "Action parameters"},
                },
                "required": ["action"],
            },
        ),
        _Tool(
            action="list_slide_shapes",
            tool_name="ppt_list_shapes",
            desc=(
                "List all shapes on a slide: index, id, name, type, position (pt), size (pt), text, font, fill. "
                "Always call before modifying an existing slide. Use shape id (not name) for subsequent ops."
            ),
            schema={
                "type": "object",
                "properties": {
                    "slide": {"type": "integer", "description": "1-based slide number"},
                },
                "required": ["slide"],
            },
        ),
        _Tool(
            action="edit_slide_xml",
            tool_name="ppt_edit_xml",
            desc=(
                "Edit a slide by running JavaScript code with JSZip access to the PPTX ZIP internals. "
                "Primary tool for precise OOXML layouts, backgrounds, and complex arrangements.\n\n"
                "Code receives: zip (JSZip), markDirty(), slidePath (string), slideNum (int). "
                "Modify zip files then call markDirty() to trigger reimport into PowerPoint."
            ),
            schema={
                "type": "object",
                "properties": {
                    "slide":       {"type": "integer", "description": "1-based slide number"},
                    "code":        {"type": "string",  "description": "Async JS body: receives (zip, markDirty, slidePath, slideNum)"},
                    "explanation": {"type": "string",  "description": "What this edit does"},
                },
                "required": ["slide", "code"],
            },
        ),
        _Tool(
            action="edit_slide_chart",
            tool_name="ppt_edit_chart",
            desc=(
                "Insert or replace a chart using OOXML. "
                "Required 3-file pattern: ppt/charts/chart{n}.xml + [Content_Types].xml override + rels entry. "
                "Same interface as ppt_edit_xml — code receives (zip, markDirty, slidePath, slideNum)."
            ),
            schema={
                "type": "object",
                "properties": {
                    "slide":       {"type": "integer", "description": "1-based slide number"},
                    "code":        {"type": "string",  "description": "Async JS body modifying chart XML files"},
                    "explanation": {"type": "string",  "description": "What chart is being added/replaced"},
                },
                "required": ["slide", "code"],
            },
        ),
        _Tool(
            action="verify_slides",
            tool_name="ppt_verify",
            desc=(
                "Check slide layout quality. Returns JSON per slide:\n"
                "- coverage_ratio: content area / safe area (≥0.70 good)\n"
                "- bottom_gap: pt from lowest shape to slide bottom (≤100 good)\n"
                "- overlaps: overlapping shape name pairs\n"
                "- out_of_bounds: shapes outside slide edges\n"
                "- contrast_warnings: text vs fill contrast < 4.5:1 (WCAG AA)"
            ),
            schema={
                "type": "object",
                "properties": {
                    "from_slide": {"type": "integer", "description": "Start slide (default: 1)"},
                    "to_slide":   {"type": "integer", "description": "End slide (default: last)"},
                },
            },
        ),
        _Tool(
            action="verify_slide_visual",
            tool_name="ppt_screenshot",
            desc=(
                "Take a screenshot of a slide for visual verification. "
                "Always include a 'question' describing what you just did and what to check."
            ),
            schema={
                "type": "object",
                "properties": {
                    "slide":    {"type": "integer", "description": "1-based slide number"},
                    "question": {"type": "string",  "description": "What you just edited and what to verify (e.g. '제목 텍스트와 배경색이 올바르게 적용됐는지 확인')"},
                },
                "required": ["slide", "question"],
            },
        ),
        _Tool(
            action="set_z_order",
            tool_name="ppt_set_z_order",
            desc="Change shape z-order: BringToFront, SendToBack, BringForward, or SendBackward.",
            schema={
                "type": "object",
                "properties": {
                    "slide":       {"type": "integer", "description": "1-based slide number"},
                    "shape_id":    {"type": "string",  "description": "Shape ID from ppt_list_shapes"},
                    "shape_index": {"type": "integer", "description": "1-based index (fallback)"},
                    "order": {
                        "type": "string",
                        "enum": ["BringToFront", "SendToBack", "BringForward", "SendBackward"],
                    },
                },
                "required": ["slide", "order"],
            },
        ),
    ]
