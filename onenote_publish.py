"""
out.json 내용을 Microsoft 원노트에 새 페이지로 만든다 (Microsoft Graph).

사전 준비:
1) Azure Portal에서 앱 등록 → 공용 클라이언트(모바일 및 데스크톱) 허용.
2) API 사용 권한(Microsoft Graph 위임): Notes.ReadWrite, User.Read — 관리자 동의 필요 시 승인.
3) 이 터미널에서: export ONENOTE_CLIENT_ID='앱(클라이언트) ID'

원노트는 클라우드에 동기화된 노트북이 있어야 한다. 로컬만·비로그인이면 Graph로 접근할 수 없다.

선택: 특정 섹션에만 넣으려면 export ONENOTE_SECTION_ID='섹션 GUID'
"""

from __future__ import annotations

import html
import json
import os
import sys
from pathlib import Path

import httpx
import msal

CLIENT_ID = os.environ.get("ONENOTE_CLIENT_ID")
AUTHORITY = os.environ.get(
    "ONENOTE_AUTHORITY", "https://login.microsoftonline.com/common"
)
SCOPES = [
    "https://graph.microsoft.com/Notes.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]
GRAPH = "https://graph.microsoft.com/v1.0"
OUT_JSON = Path(__file__).resolve().parent / "out.json"


def get_token() -> str:
    if not CLIENT_ID:
        print(
            "ONENOTE_CLIENT_ID 환경 변수를 설정하세요. (Azure 앱 등록의 애플리케이션(클라이언트) ID)",
            file=sys.stderr,
        )
        sys.exit(1)
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(flow.get("error_description", flow))
    print(flow["message"], flush=True)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", result))
    return result["access_token"]


def first_section_id(token: str) -> str:
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client() as client:
        r = client.get(f"{GRAPH}/me/onenote/notebooks", headers=headers, timeout=60)
        r.raise_for_status()
        values = r.json().get("value", [])
        if not values:
            raise RuntimeError(
                "원노트 노트북이 없습니다. 원노트에서 노트북을 하나 만든 뒤 다시 시도하세요."
            )
        nb_id = values[0]["id"]
        r = client.get(
            f"{GRAPH}/me/onenote/notebooks/{nb_id}/sections",
            headers=headers,
            timeout=60,
        )
        r.raise_for_status()
        secs = r.json().get("value", [])
        if not secs:
            raise RuntimeError("노트북에 섹션이 없습니다. 원노트에서 섹션을 추가하세요.")
        return secs[0]["id"]


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


def create_page(token: str, section_id: str, html_doc: str) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH}/me/onenote/sections/{section_id}/pages"
    files = {"Presentation": (None, html_doc, "text/html; charset=utf-8")}
    with httpx.Client() as client:
        r = client.post(url, headers=headers, files=files, timeout=120)
        if r.status_code >= 400:
            raise RuntimeError(f"{r.status_code} {r.text}")
        return r.json()


def main() -> None:
    if not OUT_JSON.is_file():
        print("out.json 이 없습니다. 먼저 python process.py 를 실행하세요.", file=sys.stderr)
        sys.exit(1)
    data = json.loads(OUT_JSON.read_text(encoding="utf-8"))
    html_doc = json_to_html(data)
    token = get_token()
    section_id = os.environ.get("ONENOTE_SECTION_ID") or first_section_id(token)
    if not os.environ.get("ONENOTE_SECTION_ID"):
        print("첫 번째 노트북의 첫 번째 섹션에 페이지를 만듭니다.", flush=True)
    result = create_page(token, section_id, html_doc)
    link = (
        result.get("links", {})
        .get("oneNoteClientUrl", {})
        .get("href")
    )
    print("완료:", link or result.get("id", result))


if __name__ == "__main__":
    main()
