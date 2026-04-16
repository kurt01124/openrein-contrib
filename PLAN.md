# openrein-contrib 구현 계획

## 개요

`openrein-contrib`는 openrein 기반 에이전트(dressage 등)에서 사용할
Tools와 Skills를 제공하는 패키지다.

태그 정의 파일(`tags/*.json`)에 tool + skill을 선언하면
`load_tag_defs()`가 이를 읽어 에이전트에 등록할 수 있는 형태로 반환한다.

---

## 태그 목록

### `web` — 웹 브라우저 제어
- **Tools**: `BrowserTool`
- **Skills**: `web_navigation`
- **의존성**: `playwright`
- **상태**: 구현 예정
- **설명**: Playwright로 실제 브라우저를 제어. 구글 검색, 로그인,
  SPA 페이지, 폼 조작 등 WebFetch로 불가능한 작업 처리.

### `git` — Git 작업
- **Tools**: `GitTool`
- **Skills**: `commit`, `pr`, `code_review`
- **의존성**: 없음 (subprocess)
- **상태**: 구현 예정
- **설명**: git 명령 실행, 커밋 메시지 생성, PR 워크플로우.

### `data` — 데이터/DB
- **Tools**: `SQLiteTool`
- **Skills**: `sql_query`
- **의존성**: 없음 (stdlib sqlite3)
- **상태**: 구현 예정
- **설명**: SQLite DB 쿼리 및 결과 분석.

### `office` — Office 문서 제어
- **Tools**: `OfficeTool`
- **Skills**: `powerpoint`, `excel`, `word`
- **의존성**: `pywin32` (Windows), ClaudeGateway addin
- **상태**: 구현 예정
- **설명**: PowerPoint/Word/Excel을 COM 인터페이스로 직접 제어.
  참고: `ClaudeGateway/client/archive_tools.py`

---

## Skills 목록 (MD 파일)

| 파일 | 태그 | 내용 |
|------|------|------|
| `web_navigation.md` | web | BrowserTool 사용법 (navigate→snapshot→act 루프) |
| `commit.md` | git | 좋은 커밋 메시지 작성 가이드 |
| `pr.md` | git | PR 제목/본문 작성 가이드 |
| `code_review.md` | git | 코드 리뷰 체크리스트 |
| `sql_query.md` | data | SQL 쿼리 작성 패턴 |
| `powerpoint.md` | office | PPT 슬라이드 구성 가이드 |
| `excel.md` | office | Excel 수식/시트 조작 가이드 |
| `word.md` | office | Word 문서 서식 가이드 |

---

## Tools 구현 순서

1. **BrowserTool** (`tools/browser.py`) — Playwright 기반 브라우저 제어
   - `navigate(url)` — URL 이동
   - `snapshot()` — 페이지 접근성 트리 읽기 (ARIA)
   - `act(kind, ...)` — click / type / fill / press / select
   - `screenshot()` — 화면 캡처
   - `close()` — 브라우저 종료

2. **GitTool** (`tools/git.py`) — git subprocess 래퍼
   - `run(args)` — git 명령 실행
   - `status()` / `diff()` / `log()`

3. **SQLiteTool** (`tools/sqlite.py`) — sqlite3 래퍼
   - `query(db_path, sql)` — SELECT 실행
   - `execute(db_path, sql)` — INSERT/UPDATE/DELETE

4. **OfficeTool** (`tools/office.py`) — COM 제어 (Windows)
   - ClaudeGateway `archive_tools.py` 코드 이식

---

## 태그 정의 시스템

```
openrein/contrib/
├── tags/
│   ├── web.json      ← Tag 선언
│   ├── git.json
│   ├── data.json
│   └── office.json
├── tools/
│   ├── browser.py    ← BrowserTool
│   ├── git.py        ← GitTool
│   ├── sqlite.py     ← SQLiteTool
│   └── office.py     ← OfficeTool
└── skills/
    ├── web_navigation.md
    ├── commit.md
    └── ...
```

JSON 예시 (`tags/web.json`):
```json
{
  "name": "web",
  "description": "...",
  "tools": ["BrowserTool"],
  "skills": ["web_navigation"],
  "optional_deps": ["playwright"]
}
```

로드 방법 (dressage에서):
```python
from openrein.contrib import load_tag_defs
from dressage.tag import Tag

for d in load_tag_defs():
    tags.append(Tag(**d))
```
