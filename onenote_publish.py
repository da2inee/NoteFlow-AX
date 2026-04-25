"""
out.json 내용을 Microsoft 원노트에 새 페이지로 만든다 (Microsoft Graph).

사전 준비:
1) Azure Portal에서 앱 등록 → 공용 클라이언트(모바일 및 데스크톱) 허용.
2) API 사용 권한(Microsoft Graph 위임): Notes.ReadWrite, User.Read — 관리자 동의 필요 시 승인.
3) `ONENOTE_CLIENT_ID` 설정: 터미널 `export` **또는** 프로젝트 루트 `.env`에 `ONENOTE_CLIENT_ID=...` 한 줄 (`.env`는 git에 안 올라감)

원노트는 클라우드에 동기화된 노트북이 있어야 한다. 로컬만·비로그인이면 Graph로 접근할 수 없다.

선택: 특정 섹션에만 넣으려면 export ONENOTE_SECTION_ID='섹션 GUID'
"""

from __future__ import annotations

import argparse
import html
import json
import os
import secrets
import sys
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import httpx
import msal

REPO_DIR = Path(__file__).resolve().parent
_ENV_CANDIDATES = (".env", "noteflow.env")


def _apply_env_line(k: str, v: str) -> None:
    if not k:
        return
    v = (v or "").strip()
    if len(v) >= 2 and v[0] in "'\"" and v[0] == v[-1]:
        v = v[1:-1]
    if k not in os.environ:
        os.environ[k] = v
        return
    if k == "ONENOTE_CLIENT_ID" and not (os.environ.get("ONENOTE_CLIENT_ID") or "").strip():
        os.environ[k] = v


def _parse_env_file(raw: str) -> None:
    for line in raw.splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        if s.startswith("export "):
            s = s[7:].strip()
        if "=" not in s:
            continue
        k, _, v = s.partition("=")
        k = k.strip()
        _apply_env_line(k, v)


def _load_local_env() -> None:
    """요청마다 호출해도 됨(서버 가동 후 .env 생성·수정 반영)."""
    for name in _ENV_CANDIDATES:
        p = REPO_DIR / name
        if not p.is_file():
            continue
        try:
            _parse_env_file(p.read_text(encoding="utf-8-sig"))
        except (OSError, UnicodeError):
            continue

    one = REPO_DIR / "onenote_client_id.txt"
    if one.is_file():
        try:
            s = one.read_text(encoding="utf-8-sig").strip()
            if s and (
                "ONENOTE_CLIENT_ID" not in os.environ
                or not (os.environ.get("ONENOTE_CLIENT_ID") or "").strip()
            ):
                os.environ["ONENOTE_CLIENT_ID"] = s
        except OSError:
            pass


_load_local_env()


def get_client_id() -> str | None:
    _load_local_env()
    w = (os.environ.get("ONENOTE_CLIENT_ID") or "").strip()
    return w or None


def env_file_hint() -> str:
    """에러 메시지용: 어떤 경로를 봤는지."""
    _load_local_env()
    parts = [f"폴더: {REPO_DIR}"]
    for name in _ENV_CANDIDATES:
        p = REPO_DIR / name
        parts.append(f"  {p.name} → {'있음' if p.is_file() else '없음'}")
    p1 = REPO_DIR / "onenote_client_id.txt"
    parts.append(f"  {p1.name} → {'있음' if p1.is_file() else '없음'}")
    return "\n".join(parts)


_load_local_env()

AUTHORITY = os.environ.get(
    "ONENOTE_AUTHORITY", "https://login.microsoftonline.com/common"
)
SCOPES = [
    "https://graph.microsoft.com/Notes.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]
GRAPH = "https://graph.microsoft.com/v1.0"
TZ_SEOUL = ZoneInfo("Asia/Seoul")
OUT_JSON = Path(__file__).resolve().parent / "out.json"
OUT_MD = Path(__file__).resolve().parent / "out.md"
CACHE_PATH = Path(__file__).resolve().parent / ".msal_cache.json"
DEFAULT_NOTEBOOK_NAME = "Note_20260416_171212"


class OneNoteConfigError(RuntimeError):
    """ONENOTE_CLIENT_ID 등이 없을 때. 웹 서버(wsgi 아닌 베이스 HTTP)는 exit 대신 잡는다."""


def _build_app() -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
    cid = get_client_id()
    if not cid:
        raise OneNoteConfigError(
            "ONENOTE_CLIENT_ID를 설정하세요. "
            f"({REPO_DIR}/.env에 ONENOTE_CLIENT_ID=앱클라이언트ID) 또는 "
            "터미널에서 export ONENOTE_CLIENT_ID='Azure 앱 등록의 애플리케이션(클라이언트) ID'"
        )
    cache = msal.SerializableTokenCache()
    if CACHE_PATH.is_file():
        try:
            cache.deserialize(CACHE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    app = msal.PublicClientApplication(cid, authority=AUTHORITY, token_cache=cache)
    return app, cache


def _persist_cache(cache: msal.SerializableTokenCache) -> None:
    if getattr(cache, "has_state_changed", False):
        CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")


def get_token(*, interactive: bool = True) -> str | None:
    """
    interactive=False 인 경우: 캐시에 토큰이 있으면 반환, 없으면 None.
    interactive=True  인 경우: 필요하면 device code 로그인까지 수행.
    """
    try:
        app, cache = _build_app()
    except OneNoteConfigError:
        if not interactive:
            return None
        raise
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _persist_cache(cache)
            return result["access_token"]

    if not interactive:
        return None

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(flow.get("error_description", flow))
    print(flow["message"], flush=True)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", result))
    _persist_cache(cache)
    return result["access_token"]


def list_notebooks(token: str) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client() as client:
        r = client.get(f"{GRAPH}/me/onenote/notebooks", headers=headers, timeout=60)
        r.raise_for_status()
        values = list(r.json().get("value", []))
        if not values:
            raise RuntimeError("원노트 노트북이 없습니다. 원노트에서 노트북을 하나 만든 뒤 다시 시도하세요.")
        return values


def pick_notebook_id(token: str, notebook_name: str | None) -> str:
    notebooks = list_notebooks(token)
    name = (notebook_name or os.environ.get("ONENOTE_NOTEBOOK_NAME") or DEFAULT_NOTEBOOK_NAME).strip()
    for nb in notebooks:
        if str(nb.get("displayName", "")).strip() == name:
            return nb["id"]
    # 고정 이름이 없으면 첫 노트북으로 fallback
    return notebooks[0]["id"]


def list_sections(token: str, notebook_id: str) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client() as client:
        r = client.get(
            f"{GRAPH}/me/onenote/notebooks/{notebook_id}/sections",
            headers=headers,
            timeout=60,
        )
        r.raise_for_status()
        return list(r.json().get("value", []))


def create_section(token: str, notebook_id: str, name: str) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client() as client:
        r = client.post(
            f"{GRAPH}/me/onenote/notebooks/{notebook_id}/sections",
            headers=headers,
            json={"displayName": name},
            timeout=60,
        )
        if r.status_code >= 400:
            raise RuntimeError(f"{r.status_code} {r.text}")
        return r.json()


def ensure_section_id(token: str, notebook_id: str, section_name: str) -> str:
    section_name = (section_name or "").strip()
    if not section_name:
        section_name = "ax 과제"
    secs = list_sections(token, notebook_id)
    for sec in secs:
        if str(sec.get("displayName", "")).strip() == section_name:
            return sec["id"]
    created = create_section(token, notebook_id, section_name)
    return created["id"]


def json_to_html(data: dict) -> str:
    body = json.dumps(data, ensure_ascii=False, indent=2)
    safe = html.escape(body)
    title = html.escape(str(data.get("title") or "NoteFlow 정리"))
    return f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>{title}</title>
</head>
<body>
  <h1>{title}</h1>
  <pre>{safe}</pre>
</body>
</html>"""


def md_to_html(title: str, md_text: str, meta: dict | None = None) -> str:
    safe_md = html.escape(md_text or "")
    title_safe = html.escape(title or "NoteFlow 정리")
    meta_lines = []
    if meta:
        cat = meta.get("categories") or []
        sec = meta.get("recommended_section") or ""
        if cat:
            meta_lines.append(f"<div><b>카테고리</b>: {html.escape(', '.join(map(str, cat)))}</div>")
        if str(sec).strip():
            meta_lines.append(f"<div><b>탭</b>: {html.escape(str(sec))}</div>")
    meta_html = "\n".join(meta_lines)
    return f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>{title_safe}</title>
</head>
<body>
  <h1>{title_safe}</h1>
  {meta_html}
  <pre>{safe_md}</pre>
</body>
</html>"""


def create_page(token: str, section_id: str, html_doc: str) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH}/me/onenote/sections/{section_id}/pages"
    files = {"Presentation": (None, html_doc, "text/html; charset=utf-8")}
    with httpx.Client() as client:
        r = client.post(url, headers=headers, files=files, timeout=120)
        if r.status_code >= 400:
            raise RuntimeError(f"{r.status_code} {r.text}")
        return r.json()


def topic_title_marker(section: str, page_keyword: str) -> str:
    """페이지 제목·탐색에 쓰는 토픽 식별 문자열(중복이면 이 문자열이 제목에 포함됨)."""
    return f"[{section.strip()}][{str(page_keyword).strip()}]"


def list_section_pages(token: str, section_id: str) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    out: list[dict] = []
    next_url: str | None = f"{GRAPH}/me/onenote/sections/{section_id}/pages"
    first_params: dict = {"$select": "id,title,lastModifiedDateTime", "$top": 100}
    with httpx.Client() as client:
        while next_url:
            r = client.get(
                next_url,
                headers=headers,
                params=first_params if out == [] else None,
                timeout=60,
            )
            if r.status_code >= 400:
                raise RuntimeError(f"{r.status_code} {r.text}")
            j = r.json()
            out.extend(j.get("value", []))
            next_url = j.get("@odata.nextLink")
    return out


def find_page_id_for_topic_marker(
    token: str, section_id: str, marker: str) -> str | None:
    """제목에 marker(예: [cos 개발][캠페인 미션])이 들어간 페이지 1개. 복수면 최근 수정이 우선."""
    if not marker or not str(marker).strip():
        return None
    m = str(marker).strip()
    pages = [p for p in list_section_pages(token, section_id) if m in (p.get("title") or "")]
    if not pages:
        return None
    pages.sort(key=lambda p: p.get("lastModifiedDateTime") or "", reverse=True)
    return pages[0].get("id")


def get_onenote_client_url_for_page(token: str, page_id: str) -> str | None:
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client() as client:
        r = client.get(
            f"{GRAPH}/me/onenote/pages/{page_id}?$select=links",
            headers=headers,
            timeout=60,
        )
        if r.status_code >= 400:
            return None
        j = r.json()
        return ((j.get("links") or {}).get("oneNoteClientUrl") or {}).get("href")


def today_ymd_seoul() -> str:
    return datetime.now(tz=TZ_SEOUL).date().isoformat()


def now_hhmm_seoul() -> str:
    """서울 기준 저장 시각 (같은 날 이어 붙일 때 항목별로 표시)."""
    return datetime.now(tz=TZ_SEOUL).strftime("%H:%M")


def get_page_content_html(token: str, page_id: str) -> str:
    """Graph: GET /me/onenote/pages/{id}/content (HTML)."""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH}/me/onenote/pages/{page_id}/content"
    with httpx.Client() as client:
        r = client.get(url, headers=headers, timeout=120)
        if r.status_code >= 400:
            raise RuntimeError(f"{r.status_code} {r.text[:800]}")
        return r.text


def day_container_data_id(ymd: str) -> str:
    safe = (ymd or "").strip().replace("-", "")
    return f"noteflow-day-{safe}"


def append_html(
    token: str, page_id: str, *, target: str, inner_html: str, mark_item: bool = False
) -> None:
    """Graph PATCH로 특정 target(body 또는 #data-id)에 HTML을 append."""
    wrap = "nf" + secrets.token_hex(6)
    item_attr = ' data-tag="noteflow-item"' if mark_item else ""
    block = f'<div data-id="{wrap}"{item_attr}>{inner_html}</div>'
    body = [{"target": target, "action": "append", "content": block}]
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GRAPH}/me/onenote/pages/{page_id}/content"
    with httpx.Client() as client:
        r = client.patch(url, headers=headers, json=body, timeout=120)
        if r.status_code >= 400:
            raise RuntimeError(f"{r.status_code} {r.text}")


def ensure_day_container(token: str, page_id: str, *, ymd: str) -> str:
    """
    하루(ymd) 단위 컨테이너 div를 보장하고, 그 컨테이너의 id(= PATCH target #...)를 반환한다.

    Graph PATCH 의 target #foo 는 HTML id 속성을 가리키므로 data-id 가 아니라 id 를 쓴다.
    """
    did = day_container_data_id(ymd)
    try:
        page_html = get_page_content_html(token, page_id)
    except Exception:
        page_html = ""
    if f'id="{did}"' in page_html:
        return did
    dsafe = html.escape(ymd)
    container_html = f"""
<div id="{did}" data-noteflow-day="{ymd}">
  <h2><b>{dsafe}</b></h2>
</div>
""".strip()
    append_html(token, page_id, target="body", inner_html=container_html)
    return did


def first_nonempty_line(text: str, *, max_len: int = 2000) -> str:
    """원문에서 첫 비어 있지 않은 줄만 잘라 날짜 블록 아래 한 줄 요약으로 쓴다."""
    for line in (text or "").splitlines():
        t = line.strip()
        if t:
            return t[:max_len]
    return ""


def _day_container_has_item_entries(page_html: str, *, did: str) -> bool:
    """이미 같은 날짜 블록 안에 붙인 항목(noteflow-item)이 있는지."""
    needle = f'id="{did}"'
    i = page_html.find(needle)
    if i < 0:
        return False
    j = page_html.find(">", i)
    if j < 0:
        return False
    tail = page_html[j + 1 : j + 400_000]
    return 'data-tag="noteflow-item"' in tail


def append_item_under_today_container(
    token: str, page_id: str, *, display_line: str
) -> None:
    """같은 토픽 페이지: 오늘(서울) 날짜 블록 아래에 항목 1개(첫 줄은 짧은 한 줄만, 이어 붙이면 ---- + 시각 + 줄)."""
    ymd = today_ymd_seoul()
    did = ensure_day_container(token, page_id, ymd=ymd)
    try:
        page_html = get_page_content_html(token, page_id)
    except Exception:
        page_html = ""
    cont = _day_container_has_item_entries(page_html, did=did)
    inner = md_to_html_append_day_entry(
        display_line=display_line,
        is_continuation=cont,
        time_hhmm=now_hhmm_seoul() if cont else "",
    )
    append_html(
        token, page_id, target=f"#{did}", inner_html=inner, mark_item=True
    )


def append_to_page_content(token: str, page_id: str, inner_html: str) -> None:
    """기본: body 끝에 append."""
    append_html(token, page_id, target="body", inner_html=inner_html)


def md_to_html_append_day_entry(
    *,
    display_line: str,
    is_continuation: bool = False,
    time_hhmm: str = "",
) -> str:
    """
    날짜 블록 아래 한 항목.
    - 당일 첫 줄: 본문 한 줄만.
    - 같은 날 이어 붙이기: ----, 다음 줄에 HH:MM, 그 다음에 본문 한 줄. (탭/메타/pre 없음)
    """
    line = (display_line or "").strip() or "(빈 메모)"
    p_body = "nf" + secrets.token_hex(4)
    esc = html.escape(line)
    if not is_continuation:
        return f'<p data-id="{p_body}">{esc}</p>'
    sep = "nf" + secrets.token_hex(4)
    p_time = "nf" + secrets.token_hex(4)
    t = html.escape((time_hhmm or "00:00").strip()[:5])
    return f"""<p data-id="{sep}">----</p>
<p data-id="{p_time}"><b>{t}</b></p>
<p data-id="{p_body}">{esc}</p>"""


def new_topic_page_with_day_entry(
    *, page_title: str, ymd: str, display_line: str
) -> str:
    """
    토픽(페이지 키워드)이 있는 **첫** 저장: 본문은 H1(원노트 목록·제목 매칭용)만 두고,
    그 아래 당일 날짜 블록 + 첫 줄 1개. (탭/메타/전문 pre 없음 — 이어 붙이기와 동일한 형식)
    """
    tesc = html.escape((page_title or "").strip() or "NoteFlow")
    did = day_container_data_id(ymd)
    dsafe = html.escape(ymd)
    yattr = html.escape(ymd)
    item_inner = md_to_html_append_day_entry(display_line=display_line)
    wrap = "nf" + secrets.token_hex(6)
    item_block = f'<div data-id="{wrap}" data-tag="noteflow-item">{item_inner}</div>'
    return f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>{tesc}</title>
</head>
<body>
  <h1 style="font-size:11px;font-weight:500;color:#9ca3af;letter-spacing:0.01em;">{tesc}</h1>
  <div id="{did}" data-noteflow-day="{yattr}">
    <h2><b>{dsafe}</b></h2>
    {item_block}
  </div>
</body>
</html>"""


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--section", default=os.environ.get("ONENOTE_SECTION_NAME"))
    p.add_argument("--raw-json", action="store_true")
    args = p.parse_args()

    if not OUT_JSON.is_file():
        print("out.json 이 없습니다. 먼저 python process.py 를 실행하세요.", file=sys.stderr)
        sys.exit(1)
    data = json.loads(OUT_JSON.read_text(encoding="utf-8"))
    title = str(data.get("title") or "NoteFlow 정리")
    if args.raw_json or not OUT_MD.is_file():
        html_doc = json_to_html(data)
    else:
        html_doc = md_to_html(title=title, md_text=OUT_MD.read_text(encoding="utf-8"), meta=data)

    try:
        token = get_token(interactive=True)
    except OneNoteConfigError as e:
        print(str(e), file=sys.stderr)
        sys.exit(1)
    if not token:
        print("토큰을 획득하지 못했습니다.", file=sys.stderr)
        sys.exit(1)
    notebook_id = pick_notebook_id(token, None)
    section_name = args.section or str(data.get("recommended_section") or "").strip() or "ax 과제"
    section_id = ensure_section_id(token, notebook_id, section_name)
    result = create_page(token, section_id, html_doc)
    link = ((result.get("links") or {}).get("oneNoteClientUrl", {}) or {}).get("href")
    print("완료:", link or result.get("id", result))


if __name__ == "__main__":
    main()
