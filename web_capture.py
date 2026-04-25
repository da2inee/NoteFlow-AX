from __future__ import annotations

import html
import json
import os
import re
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs

import onenote_publish
import process
import render_md


def html_page(body: str, *, title: str = "NoteFlow AX") -> bytes:
    doc = f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{title}</title>
  <style>
    :root {{
      --onenote-purple: #7719aa;
      --bg: #f5f6f8;
      --panel: #ffffff;
      --muted: #6b7280;
      --border: #e5e7eb;
      --text: #111827;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      margin: 0;
      background: var(--bg);
      color: var(--text);
    }}
    .topbar {{
      height: 52px;
      background: var(--onenote-purple);
      color: #fff;
      display: flex;
      align-items: center;
      padding: 0 16px;
      font-weight: 600;
      letter-spacing: 0.2px;
    }}
    .topbar small {{
      opacity: 0.85;
      font-weight: 500;
      margin-left: 10px;
    }}
    .wrap {{
      display: grid;
      grid-template-columns: 260px 1fr;
      gap: 16px;
      padding: 16px;
      max-width: 1200px;
      margin: 0 auto;
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 12px;
      overflow: hidden;
    }}
    .sidebar {{
      padding: 12px;
    }}
    .sidebar h3 {{
      margin: 6px 0 10px;
      font-size: 14px;
      color: var(--muted);
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.06em;
    }}
    .secbtn {{
      width: 100%;
      text-align: left;
      padding: 10px 10px;
      border-radius: 10px;
      border: 1px solid transparent;
      background: transparent;
      cursor: pointer;
      font-size: 14px;
    }}
    .secbtn:hover {{
      background: #f3f4f6;
      border-color: #eef0f2;
    }}
    .main {{
      padding: 14px;
    }}
    .row {{ margin: 12px 0; }}
    label {{ font-size: 13px; font-weight: 600; }}
    textarea {{
      width: 100%;
      min-height: 320px;
      font-size: 14px;
      line-height: 1.55;
      padding: 12px;
      border-radius: 12px;
      border: 1px solid var(--border);
      background: #fff;
      resize: vertical;
    }}
    input[type="text"] {{
      width: 100%;
      padding: 10px;
      border-radius: 10px;
      border: 1px solid var(--border);
      font-size: 14px;
    }}
    .hint {{ color: var(--muted); font-size: 13px; line-height: 1.45; }}
    code {{ background: #f3f4f6; padding: 2px 6px; border-radius: 6px; }}
    .actions {{
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
    }}
    .primary {{
      background: var(--onenote-purple);
      border: 1px solid rgba(0,0,0,0.08);
      color: #fff;
      padding: 10px 14px;
      border-radius: 10px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
    }}
    .primary:disabled {{ opacity: 0.6; cursor: not-allowed; }}
    .pill {{
      display: inline-flex;
      gap: 8px;
      align-items: center;
      padding: 8px 10px;
      border: 1px solid var(--border);
      border-radius: 999px;
      background: #fff;
      font-size: 13px;
      color: var(--muted);
    }}
    pre {{ background: #0b1020; color: #e6e6e6; padding: 12px; border-radius: 10px; overflow: auto; }}
    a {{ color: #0b57d0; }}
    .overlay {{
      position: fixed;
      inset: 0;
      background: rgba(17,24,39,0.40);
      display: none;
      align-items: center;
      justify-content: center;
      padding: 16px;
      z-index: 1000;
    }}
    .overlay .card {{
      width: min(520px, 100%);
      background: #fff;
      border-radius: 14px;
      border: 1px solid var(--border);
      padding: 16px;
    }}
    .spinner {{
      width: 22px;
      height: 22px;
      border-radius: 50%;
      border: 3px solid #e5e7eb;
      border-top-color: var(--onenote-purple);
      display: inline-block;
      animation: spin 0.9s linear infinite;
      vertical-align: -5px;
      margin-right: 10px;
    }}
    @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
  </style>
</head>
<body>
<div class="topbar">OneNote Capture <small>NoteFlow AX</small></div>
{body}
<div id="overlay" class="overlay" aria-hidden="true">
  <div class="card">
    <div style="font-weight:800; margin-bottom:6px;"><span class="spinner"></span>저장 중…</div>
    <div class="hint">
      처음 1회는 로그인/권한 동의로 오래 걸릴 수 있어요. 이후엔 훨씬 빨라집니다.<br/>
      빠르게 끝내려면 <b>빠른 모드(LLM 생략)</b>를 켜고 섹션을 직접 선택해도 됩니다.
    </div>
  </div>
</div>
<script>
  function showOverlay() {{
    const o = document.getElementById('overlay');
    if (o) o.style.display = 'flex';
  }}
  function pickSection(name) {{
    const el = document.querySelector('input[name="section"]');
    if (el) el.value = name;
    const ta = document.querySelector('textarea[name="text"]');
    if (ta) ta.focus();
  }}
</script>
</body>
</html>"""
    return doc.encode("utf-8")


SECTION_CHOICES = ["cos 개발", "ax 과제", "문의", "to-do"]
SIDEBAR_BUTTONS = "\n".join(
    [
        f"""<button class="secbtn" type="button" onclick="pickSection('{s}')">{s}</button>"""
        for s in SECTION_CHOICES
    ]
)

FORM_HTML = """
<div class="wrap">
  <div class="panel sidebar">
    <h3>섹션(탭)</h3>
    {SIDEBAR_BUTTONS}
    <div class="row hint">
      위 버튼을 누르면 섹션 입력칸에 자동으로 채워집니다.<br/>
      비워두면 자동 분류(룰→LLM)를 사용합니다.
    </div>
  </div>

  <div class="panel main">
    <div class="row hint">
      노트북: <code>Note_20260416_171212</code> (고정) · 필수: <code>ONENOTE_CLIENT_ID</code> (또는 프로젝트 루트 <code>.env</code>)
      · <a href="/auth" target="_blank" rel="noreferrer">로그인/권한 동의</a>
    </div>
    <form id="capForm" method="post" action="/capture" onsubmit="showOverlay(); document.getElementById('saveBtn').disabled=true;">
      <div class="row">
        <label>텍스트</label><br/>
        <textarea name="text" placeholder="여기에 메모/대화 내용을 붙여넣으세요."></textarea>
      </div>
      <div class="row actions">
        <label class="pill"><input type="checkbox" name="raw" value="1" checked /> 원문 그대로 저장</label>
        <label class="pill"><input type="checkbox" name="fast" value="1" /> 빠른 모드(LLM 생략)</label>
        <span class="hint">빠른 모드는 룰/선택만 사용해서 훨씬 빠릅니다(애매한 경우 기본은 ax 과제).</span>
      </div>
      <div class="row">
        <label>섹션(탭) 이름(선택)</label><br/>
        <input type="text" name="section" placeholder="예: cos 개발 (비우면 자동 분류)" />
      </div>
      <div class="row actions">
        <button id="saveBtn" class="primary" type="submit">원노트에 저장</button>
      </div>
    </form>
  </div>
</div>
""".replace("{SIDEBAR_BUTTONS}", SIDEBAR_BUTTONS)


def rule_based_section(text: str) -> str | None:
    section, _reason = diagnose_section(text)
    return section


def diagnose_section(text: str) -> tuple[str | None, str]:
    """
    rule 기반 섹션 분류 + 이유(어떤 규칙/키워드가 매칭됐는지).

    반환:
      (섹션 또는 None, 이유 문자열)
    """
    t = (text or "").casefold()

    # 이름/고정 라우팅 (섹션 결정에만 사용)
    name_routes: list[tuple[str, list[str]]] = [
        ("cos 개발", ["이정훈", "이주현"]),
    ]
    for section, keys in name_routes:
        for k in keys:
            if k.casefold() in t:
                return section, f"name_routes:{section}:{k}"

    # AX 본과제·역공학·과제정의 등 (넓은 '가능'보다 먼저 매칭)
    ax_project_signals: list[str] = [
        "역공학",
        "기술자산",
        "과제정의서",
        "과제정의",
        "task 정의",
        "task",
        "중요단말",
        "안전보건",
        "기술문서 자동생성",
        "spring rest docs",
        "rest docs",
        "adoc",
        "mdx",
    ]
    for k in ax_project_signals:
        if k.casefold() in t:
            return "ax 과제", f"ax_project_signals:{k}"

    # 일반 키워드 (구체적인 섹션부터)
    rules: list[tuple[str, list[str]]] = [
        (
            "to-do",
            [
                "to-do",
                "to do",
                "todo",
                "체크리스트",
                "할 일 목록",
                "작업 목록",
                "해야 할 일",
            ],
        ),
        # '가능' 단독은 기술문서에 흔해 문의 오탐 → 제외.
        ("문의", ["문의", "질문", "어떻게", "되나요", "확인 부탁", "알려주세요", "가이드 요청", "요청 드립니다"]),
        (
            "cos 개발",
            [
                "오류",
                "에러",
                "버그",
                "장애",
                "크래시",
                "예외",
                "exception",
                "stack trace",
                "로그",
                "실패",
                "timeout",
                "재현",
                "원인",
                "해결",
                "픽스",
                "cos",
                "개발",
                "구현",
                "코드",
                "스펙",
                "api",
                "rest",
                "react",
                "fastapi",
                "ssg",
                "docusaurus",
                "next.js",
                "hugo",
                "vuepress",
                "redoc",
                "테스트",
                "빌드",
                "배포",
                "리팩토링",
            ],
        ),
        (
            "ax 과제",
            [
                "본사",
                "전사",
                "공문",
                "필수",
                "제출",
                "마감",
                "정책",
                "가이드라인",
                "요청사항",
                "ax 그룹",
                "ax 개인",
                "ax 과제",
                "회의",
                "일정",
                "공유",
                "결정",
                "안건",
                "개인",
                "내가",
                "나만",
                "혼자",
                "개인메모",
                "개인 메모",
                "메모",
                "정리",
                "아이디어",
            ],
        ),
    ]
    for section, keys in rules:
        for k in keys:
            if k.casefold() in t:
                return section, f"rules:{section}:{k}"

    return None, "no_match"


def pick_page_keyword(section: str, text: str) -> str | None:
    # "캠페인/미션" 같은 건 섹션이 아니라 페이지 키워드로만 반영
    t = (text or "").casefold()
    t_nospace = re.sub(r"\s+", "", t)
    section = (section or "").strip()
    keywords_by_section: dict[str, list[str]] = {
        "to-do": [
            "마감",
            "오늘",
            "내일",
            "긴급",
            "우선",
            "처리",
            "회신",
        ],
        "cos 개발": [
            "캠페인 미션",
            "캠페인",
            "미션",
            "API 명세",
            "기술문서",
            "문서 자동생성",
            "SSG",
            "Docusaurus",
            "Next.js",
            "Hugo",
            "VuePress",
            "Redoc",
            "FastAPI",
            "React",
            "테스트",
            "배포",
            "빌드",
            "설계",
            "리써치",
            "리서치",
        ],
        "문의": ["문의", "질문", "확인", "요청", "가이드"],
        "ax 과제": [
            "회의",
            "일정",
            "배포",
            "공유",
            "결정",
            "안건",
            "메모",
            "정리",
            "아이디어",
            "역공학",
            "과제정의",
            "본사",
        ],
    }
    for k in keywords_by_section.get(section, []):
        kk = k.casefold()
        if kk in t:
            return k
        # 사용자가 "캠페인  미션"처럼 띄어쓰기를 흔들어도 같은 토픽으로 묶이게 함
        if re.sub(r"\s+", "", kk) in t_nospace:
            return k
    return None


def build_page_title(section: str, keyword: str | None, fallback_text: str) -> str:
    base = section.strip() or "메모"
    first = (fallback_text.strip().splitlines()[0] if (fallback_text or "").strip() else "메모").strip()
    if len(first) < 2:
        first = "메모"
    if len(first) > 40:
        first = first[:40] + "…"
    if keyword and str(keyword).strip():
        return f"[{base}][{str(keyword).strip()}] {first}"
    return f"[{base}] {first}"


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        if self.path in ("/", "/index.html"):
            data = html_page(FORM_HTML)
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        if self.path == "/auth":
            try:
                token = onenote_publish.get_token(interactive=True)
                if not token:
                    raise RuntimeError("토큰을 획득하지 못했습니다.")
                body = """
<div class="wrap">
  <div class="panel main">
    <h2>로그인 완료</h2>
    <div class="hint">이제부터 저장은 재로그인 없이(대부분) 동작합니다. 원래 탭으로 돌아가서 저장을 눌러보세요.</div>
    <div class="row"><a href="/">입력 화면으로 돌아가기</a></div>
  </div>
</div>
"""
                return self._respond_html(body, status=200)
            except Exception as e:
                return self._respond_error(str(e))

        if self.path == "/health":
            token = onenote_publish.get_token(interactive=False)
            payload = json.dumps({"ok": bool(token)}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return

        if self.path == "/favicon.ico":
            self.send_error(404)
            return

        self.send_error(404)

    def do_POST(self) -> None:
        if self.path != "/capture":
            self.send_error(404)
            return

        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length).decode("utf-8", errors="replace")
        form = parse_qs(raw)
        text = (form.get("text") or [""])[0]
        raw_mode = bool((form.get("raw") or [""])[0])
        fast_mode = bool((form.get("fast") or [""])[0])
        section_override = (form.get("section") or [""])[0].strip() or None

        if not text.strip():
            return self._respond_error("텍스트가 비어 있습니다. 내용을 붙여넣고 다시 시도하세요.")
        if not onenote_publish.get_client_id():
            msg = (
                "ONENOTE_CLIENT_ID를 찾지 못했습니다.\n"
                "이 스크립트가 찾는 위치(NoteFlow AX 폴더 = onenote_publish.py와 같은 곳)에\n"
                "  · .env 또는 noteflow.env  →  ONENOTE_CLIENT_ID=클라이언트ID\n"
                "  · 또는 onenote_client_id.txt 에 ID 한 줄만\n"
                "을 넣은 뒤 web_capture 를 완전히 끄고 다시 실행하세요.\n"
                "터미널 export 도 가능합니다.\n\n"
                + onenote_publish.env_file_hint()
            )
            return self._respond_error(html.escape(msg))

        try:
            token = onenote_publish.get_token(interactive=False)
            if not token:
                body = """
<div class="wrap">
  <div class="panel main">
    <h2>로그인이 필요합니다</h2>
    <div class="hint">
      처음 1회는 Microsoft 로그인/권한 동의가 필요해요.<br/>
      아래 링크를 새 탭에서 열어 로그인 후, 다시 저장을 눌러주세요.
    </div>
    <div class="row"><a href="/auth" target="_blank" rel="noreferrer">로그인/권한 동의로 이동</a></div>
    <div class="row"><a href="/">돌아가기</a></div>
  </div>
</div>
"""
                return self._respond_html(body, status=401)

            notebook_id = onenote_publish.pick_notebook_id(token, None)

            # 섹션(탭) 결정: override > rule-based > LLM
            rb = rule_based_section(text)
            section_name = section_override or rb
            if not section_name and not fast_mode:
                section_name = process.classify_section_llm(raw=text, model=process.MODEL)
            if not section_name and fast_mode:
                section_name = "ax 과제"

            # 페이지 키워드/제목
            page_kw = pick_page_keyword(section_name, text)
            title = build_page_title(section_name, page_kw, text)

            # 본문/메타
            if raw_mode:
                md = text
                data = {"title": title, "recommended_section": section_name}
            else:
                data = process.process_text(raw=text, model=process.MODEL, temperature=0)
                data["recommended_section"] = section_name
                data["title"] = title
                md = render_md.render_markdown(data)

            section_id = onenote_publish.ensure_section_id(token, notebook_id, section_name)
            marker: str | None
            if page_kw:
                marker = onenote_publish.topic_title_marker(section_name, page_kw)
            else:
                marker = None

            existing_id = onenote_publish.find_page_id_for_topic_marker(
                token, section_id, marker) if marker else None

            if existing_id:
                display = onenote_publish.first_nonempty_line(text)
                onenote_publish.append_item_under_today_container(
                    token, existing_id, display_line=display
                )
                link = onenote_publish.get_onenote_client_url_for_page(
                    token, existing_id) or ""
                op_line = f"""같은 토픽으로 보는 기존 페이지(제목에
<code>{html.escape(marker or "")}</code> 포함)에
<strong>이어 붙였습니다</strong>. 새로고침해 확인하세요."""
            else:
                display = onenote_publish.first_nonempty_line(text)
                # 모든 새 페이지는 날짜 블록 포맷으로 통일한다.
                ymd = onenote_publish.today_ymd_seoul()
                html_doc = onenote_publish.new_topic_page_with_day_entry(
                    page_title=title, ymd=ymd, display_line=display
                )
                result = onenote_publish.create_page(token, section_id, html_doc)
                link = ((result.get("links") or {}).get("oneNoteClientUrl", {}) or {}).get("href")
                if marker:
                    op_line = f"""<code>{html.escape(section_name)}</code> 탭에
<strong>새 페이지</strong>를 만들었습니다. 다음 메모는 제목에
<code>{html.escape(marker)}</code>가 있으면 이 페이지에 합쳐집니다."""
                else:
                    op_line = f"""<code>{html.escape(section_name)}</code> 탭에
<strong>새 페이지</strong>를 만들었습니다. (페이지 키워드가 없으면
항상 새 페이지입니다.)"""
            body = f"""
<h2>완료</h2>
<div class="hint">
  {op_line}
</div>
<div class="row"><b>탭</b>: {section_name}</div>
<div class="row"><b>페이지 키워드</b>: {page_kw or '(없음)'}</div>
<div class="row"><b>이번 항목 제목</b>: {html.escape(title)}</div>
<div class="row"><b>원노트</b>: {'<a href="' + link + '" target="_blank" rel="noreferrer">열기</a>' if link else '(링크 없음)'}</div>
<div class="row"><a href="/">다시 입력</a></div>
"""
            return self._respond_html(body)
        except Exception as e:  # pragma: no cover
            return self._respond_error(str(e))

    def _respond_html(self, body: str, *, status: int = 200) -> None:
        data = html_page(body)
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        try:
            self.wfile.write(data)
        except BrokenPipeError:
            return

    def _respond_error(self, message: str) -> None:
        body = f"""
<h2>실패</h2>
<pre>{message}</pre>
<div class="row"><a href="/">돌아가기</a></div>
"""
        self._respond_html(body, status=400)


def main() -> None:
    host = os.environ.get("NOTEFLOW_HOST", "127.0.0.1")
    port = int(os.environ.get("NOTEFLOW_PORT", "8787"))
    httpd = HTTPServer((host, port), Handler)
    print(f"웹 입력창: http://{host}:{port}")
    httpd.serve_forever()


if __name__ == "__main__":
    main()

