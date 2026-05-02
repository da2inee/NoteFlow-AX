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
      --onenote-dark: #4c0b78;
      --accent: #ede3f7;
      --accent-strong: #f6edff;
      --bg: #f7f5fb;
      --panel: #ffffff;
      --muted: #667085;
      --border: #e7e0ef;
      --text: #171321;
      --shadow: 0 18px 45px rgba(76, 11, 120, 0.10);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      margin: 0;
      background:
        radial-gradient(circle at top left, rgba(119,25,170,0.16), transparent 34rem),
        linear-gradient(180deg, #fbf8ff 0%, var(--bg) 46%, #ffffff 100%);
      color: var(--text);
      min-height: 100vh;
    }}
    .topbar {{
      min-height: 60px;
      background: linear-gradient(135deg, var(--onenote-dark), var(--onenote-purple));
      color: #fff;
      display: flex;
      align-items: center;
      padding: 0 22px;
      font-weight: 800;
      letter-spacing: 0.2px;
      box-shadow: 0 10px 30px rgba(76, 11, 120, 0.22);
    }}
    .topbar small {{
      opacity: 0.88;
      font-weight: 500;
      margin-left: 10px;
    }}
    .wrap {{
      display: grid;
      grid-template-columns: 280px minmax(0, 1fr);
      gap: 18px;
      padding: 22px;
      max-width: 1180px;
      margin: 0 auto;
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 18px;
      overflow: hidden;
      box-shadow: var(--shadow);
    }}
    .sidebar {{
      padding: 16px;
      align-self: start;
      position: sticky;
      top: 18px;
    }}
    .sidebar h3 {{
      margin: 4px 0 12px;
      font-size: 14px;
      color: var(--muted);
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.06em;
    }}
    .secbtn {{
      width: 100%;
      text-align: left;
      padding: 11px 12px;
      border-radius: 12px;
      border: 1px solid transparent;
      background: #fff;
      cursor: pointer;
      font-size: 14px;
      color: var(--text);
      margin-bottom: 8px;
      transition: transform 0.15s ease, background 0.15s ease, border-color 0.15s ease;
    }}
    .secbtn:hover {{
      background: var(--accent-strong);
      border-color: #dcc6ee;
      transform: translateY(-1px);
    }}
    .main {{
      padding: 18px;
    }}
    .wrap > .main:only-child {{
      grid-column: 1 / -1;
    }}
    .hero {{
      padding: 18px 18px 16px;
      margin: -18px -18px 18px;
      background: linear-gradient(135deg, #fff 0%, var(--accent-strong) 100%);
      border-bottom: 1px solid var(--border);
    }}
    .hero h1 {{
      margin: 0 0 8px;
      font-size: 24px;
      line-height: 1.25;
      letter-spacing: -0.03em;
    }}
    .row {{ margin: 14px 0; }}
    label {{ font-size: 13px; font-weight: 700; color: #34263f; }}
    textarea {{
      width: 100%;
      min-height: 340px;
      font-size: 14px;
      line-height: 1.55;
      padding: 14px;
      border-radius: 14px;
      border: 1px solid var(--border);
      background: #fff;
      resize: vertical;
      box-shadow: inset 0 1px 2px rgba(17,24,39,0.03);
    }}
    textarea:focus,
    input[type="text"]:focus {{
      border-color: var(--onenote-purple);
      box-shadow: 0 0 0 4px rgba(119,25,170,0.12);
      outline: none;
    }}
    input[type="text"] {{
      width: 100%;
      padding: 11px 12px;
      border-radius: 12px;
      border: 1px solid var(--border);
      font-size: 14px;
      transition: border-color 0.15s ease, box-shadow 0.15s ease;
    }}
    .hint {{ color: var(--muted); font-size: 13px; line-height: 1.55; }}
    code {{ background: var(--accent); color: var(--onenote-dark); padding: 2px 6px; border-radius: 7px; }}
    .actions {{
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
    }}
    .primary {{
      background: linear-gradient(135deg, var(--onenote-purple), #9b45cf);
      border: 1px solid rgba(76,11,120,0.18);
      color: #fff;
      padding: 11px 18px;
      border-radius: 12px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      box-shadow: 0 10px 22px rgba(119,25,170,0.22);
      transition: transform 0.15s ease, box-shadow 0.15s ease;
    }}
    .primary:hover {{ transform: translateY(-1px); box-shadow: 0 12px 26px rgba(119,25,170,0.28); }}
    .primary:disabled {{ opacity: 0.6; cursor: not-allowed; }}
    .pill {{
      display: inline-flex;
      gap: 8px;
      align-items: center;
      padding: 9px 11px;
      border: 1px solid var(--border);
      border-radius: 999px;
      background: #fff;
      font-size: 13px;
      color: #4b4157;
      box-shadow: 0 4px 12px rgba(76,11,120,0.05);
    }}
    pre {{ background: #14051f; color: #f5effa; padding: 14px; border-radius: 14px; overflow: auto; }}
    a {{ color: var(--onenote-purple); font-weight: 700; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .status {{
      border: 1px solid var(--border);
      border-radius: 14px;
      background: #fff;
      padding: 12px 14px;
    }}
    .result-table {{
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      overflow: hidden;
      border: 1px solid var(--border);
      border-radius: 14px;
      background: #fff;
    }}
    .result-table th,
    .result-table td {{
      text-align: left;
      border-bottom: 1px solid var(--border);
      padding: 10px;
      vertical-align: top;
    }}
    .result-table th {{
      background: var(--accent-strong);
      color: var(--onenote-dark);
      font-size: 13px;
    }}
    .result-table tr:last-child td {{ border-bottom: 0; }}
    .overlay {{
      position: fixed;
      inset: 0;
      background: rgba(23, 19, 33, 0.50);
      backdrop-filter: blur(4px);
      display: none;
      align-items: center;
      justify-content: center;
      padding: 16px;
      z-index: 1000;
    }}
    .overlay .card {{
      width: min(520px, 100%);
      background: #fff;
      border-radius: 18px;
      border: 1px solid var(--border);
      padding: 18px;
      box-shadow: 0 22px 60px rgba(17,24,39,0.24);
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
    @media (max-width: 760px) {{
      .wrap {{ grid-template-columns: 1fr; padding: 14px; }}
      .sidebar {{ position: static; }}
      .topbar {{ padding: 0 16px; }}
      .hero h1 {{ font-size: 21px; }}
    }}
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
    <div class="hero">
      <h1>메모를 원노트로 빠르게 정리</h1>
      <div class="hint">
        원문 저장, 자동 분류, 여러 메모 일괄 처리를 한 화면에서 실행합니다.
      </div>
    </div>
    <div class="row status hint">
      노트북: <code>Note_20260416_171212</code> (고정) · 필수: <code>ONENOTE_CLIENT_ID</code> (또는 프로젝트 루트 <code>.env</code>)
      · <a href="/auth" target="_blank" rel="noreferrer">로그인/권한 동의</a>
    </div>
    <form id="capForm" method="post" action="/capture" onsubmit="showOverlay(); document.getElementById('saveBtn').disabled=true;">
      <div class="row">
        <label>텍스트</label><br/>
        <textarea name="text" placeholder="여기에 메모/대화 내용을 붙여넣으세요."></textarea>
      </div>
      <div class="row actions">
        <label class="pill"><input type="checkbox" name="multi" value="1" /> 여러 메모 한 번에</label>
        <span class="hint">구분선으로 분리해 N건을 한 번 클릭으로 처리합니다.</span>
      </div>
      <div class="row">
        <label>메모 구분선(여러 메모일 때)</label><br/>
        <input type="text" name="delimiter" value="---memo---" />
        <div class="hint">예: 텍스트에 <code>---memo---</code> 가 단독 한 줄로 있으면 그 기준으로 나눕니다.</div>
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
        <button id="saveBtn" class="primary" type="submit">원노트에 저장하기</button>
      </div>
    </form>
  </div>
</div>
""".replace("{SIDEBAR_BUTTONS}", SIDEBAR_BUTTONS)


def split_memos(raw: str, *, delimiter: str) -> list[str]:
    text = (raw or "").strip()
    if not text:
        return []
    d = (delimiter or "").strip()
    if not d:
        return [text]
    pattern = re.compile(rf"(?:^|\n)\s*{re.escape(d)}\s*(?:\n|$)", re.MULTILINE)
    parts = pattern.split(text)
    return [p.strip() for p in parts if p.strip()]


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
    # to-do는 한 페이지 인박스로 고정(항상 같은 페이지에 쌓기)
    if section == "to-do":
        return "inbox"
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


def extract_todos(text: str) -> list[str]:
    """
    메모에서 '할 일'만 가볍게 추출(룰 기반). 원본 저장은 그대로 하고, to-do 인박스에만 추가로 쌓는다.

    목표: 오탐을 줄이기 위해 '해야/해주세요/올려/확인/회신/정리' 같은 실행 동사 중심으로만 뽑는다.
    """
    out: list[str] = []
    seen: set[str] = set()
    for raw_line in (text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        # 채팅 머리말/메타 제외
        if line.startswith("[") and line.endswith("]") and len(line) < 30:
            continue
        line = re.sub(r"^[\-\*\u2022]\s*", "", line)  # bullet 제거
        if len(line) < 3:
            continue

        triggers = [
            "해야",
            "해줘",
            "해주세요",
            "부탁",
            "올려",
            "올려줘",
            "정리",
            "확인",
            "회신",
            "전달",
            "메일",
            "공유",
            "리뷰",
            "체크",
            "추가",
            "반영",
        ]
        if not any(t in line for t in triggers):
            continue

        # 질문형은 할 일로 뽑지 않음(문의/확인요청은 원문만 저장)
        if line.endswith("?") or "가능할까요" in line or "되나요" in line:
            continue

        s = re.sub(r"\s+", " ", line).strip()
        if s in seen:
            continue
        seen.add(s)
        out.append(s)
    return out


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
        multi = bool((form.get("multi") or [""])[0])
        delimiter = (form.get("delimiter") or ["---memo---"])[0]
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

            memos = split_memos(text, delimiter=delimiter) if multi else [text.strip()]
            if not memos:
                return self._respond_error("텍스트가 비어 있습니다. 내용을 붙여넣고 다시 시도하세요.")

            results: list[dict[str, str]] = []
            for idx, memo in enumerate(memos, start=1):
                # 섹션(탭) 결정: override > rule-based > LLM
                rb = rule_based_section(memo)
                section_name = section_override or rb
                if not section_name and not fast_mode:
                    section_name = process.classify_section_llm(raw=memo, model=process.MODEL)
                if not section_name and fast_mode:
                    section_name = "ax 과제"

                # 페이지 키워드/제목
                page_kw = pick_page_keyword(section_name, memo)
                title = build_page_title(section_name, page_kw, memo)

                # 본문/메타 (원노트에는 '첫 줄 + 전문'으로 저장)
                if raw_mode:
                    full_text = memo
                    full_label = "원문"
                else:
                    data = process.process_text(raw=memo, model=process.MODEL, temperature=0)
                    data["recommended_section"] = section_name
                    data["title"] = title
                    full_text = render_md.render_markdown(data)
                    full_label = "정리본"

                section_id = onenote_publish.ensure_section_id(token, notebook_id, section_name)
                marker: str | None = onenote_publish.topic_title_marker(section_name, page_kw) if page_kw else None
                existing_id = (
                    onenote_publish.find_page_id_for_topic_marker(token, section_id, marker) if marker else None
                )

                if existing_id:
                    display = onenote_publish.first_nonempty_line(memo)
                    onenote_publish.append_fulltext_under_today_container(
                        token,
                        existing_id,
                        display_line=display,
                        full_text=full_text,
                        label=full_label,
                    )
                    link = onenote_publish.get_onenote_client_url_for_page(token, existing_id) or ""
                    op_line = f"""기존 페이지(제목에 <code>{html.escape(marker or "")}</code> 포함)에 <strong>이어 붙임</strong>"""
                else:
                    display = onenote_publish.first_nonempty_line(memo)
                    ymd = onenote_publish.today_ymd_seoul()
                    html_doc = onenote_publish.new_topic_page_with_day_entry_fulltext(
                        page_title=title,
                        ymd=ymd,
                        display_line=display,
                        full_text=full_text,
                        label=full_label,
                    )
                    created = onenote_publish.create_page(token, section_id, html_doc)
                    link = ((created.get("links") or {}).get("oneNoteClientUrl", {}) or {}).get("href") or ""
                    op_line = "<strong>새 페이지 생성</strong>"

                # 할 일이 보이면 to-do 인박스에도 추가
                todos = extract_todos(memo)
                if todos:
                    todo_section_id = onenote_publish.ensure_section_id(token, notebook_id, "to-do")
                    todo_marker = onenote_publish.topic_title_marker("to-do", "inbox")
                    todo_page_id = onenote_publish.find_page_id_for_topic_marker(token, todo_section_id, todo_marker)
                    todo_title = build_page_title("to-do", "inbox", "to-do inbox")
                    if not todo_page_id:
                        ymd2 = onenote_publish.today_ymd_seoul()
                        html_doc2 = onenote_publish.new_topic_page_with_day_entry(
                            page_title=todo_title, ymd=ymd2, display_line=todos[0]
                        )
                        created2 = onenote_publish.create_page(token, todo_section_id, html_doc2)
                        todo_page_id = created2.get("id") or todo_page_id
                        rest = todos[1:]
                    else:
                        rest = todos
                    if todo_page_id:
                        for t in rest:
                            onenote_publish.append_item_under_today_container(token, todo_page_id, display_line=t)

                results.append(
                    {
                        "idx": str(idx),
                        "section": section_name,
                        "kw": page_kw or "",
                        "title": title,
                        "op": op_line,
                        "link": link,
                    }
                )

            rows = []
            for r in results:
                link_html = (
                    f'<a href="{html.escape(r["link"])}" target="_blank" rel="noreferrer">열기</a>'
                    if r.get("link")
                    else "(링크 없음)"
                )
                rows.append(
                    "<tr>"
                    f"<td>{html.escape(r['idx'])}</td>"
                    f"<td>{html.escape(r['section'])}</td>"
                    f"<td>{html.escape(r['kw']) or '(없음)'}</td>"
                    f"<td>{html.escape(r['title'])}</td>"
                    f"<td>{r['op']}</td>"
                    f"<td>{link_html}</td>"
                    "</tr>"
                )

            body = f"""
<div class="wrap">
  <div class="panel main">
    <div class="hero">
      <h1>저장 완료</h1>
      <div class="hint">총 <b>{len(results)}</b>건 처리했습니다.</div>
    </div>
<div class="row">
  <table class="result-table">
    <tr>
      <th>#</th>
      <th>탭</th>
      <th>키워드</th>
      <th>제목</th>
      <th>처리</th>
      <th>원노트</th>
    </tr>
    {''.join(rows)}
  </table>
</div>
<div class="row"><a href="/">다시 입력</a></div>
  </div>
</div>
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
<div class="wrap">
  <div class="panel main">
    <div class="hero">
      <h1>실패</h1>
      <div class="hint">아래 내용을 확인한 뒤 다시 시도하세요.</div>
    </div>
    <pre>{message}</pre>
    <div class="row"><a href="/">돌아가기</a></div>
  </div>
</div>
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

