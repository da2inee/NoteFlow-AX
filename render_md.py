import json
from datetime import datetime
from pathlib import Path


IN_PATH = Path(__file__).resolve().parent / "out.json"
OUT_PATH = Path(__file__).resolve().parent / "out.md"


def md_escape(text: str) -> str:
    return text.replace("\r\n", "\n").strip()


def bullet_lines(items: list[str]) -> str:
    if not items:
        return "- (없음)"
    return "\n".join(f"- {md_escape(str(x))}" for x in items if str(x).strip()) or "- (없음)"


def action_items_table(items: list[dict]) -> str:
    if not items:
        return "| 할 일 | 담당 | 기한 |\n|---|---|---|\n| (없음) |  |  |\n"
    lines = ["| 할 일 | 담당 | 기한 |", "|---|---|---|"]
    for it in items:
        task = md_escape(str(it.get("task", "")))
        owner = md_escape(str(it.get("owner", "")))
        due = md_escape(str(it.get("due", "")))
        if not task:
            continue
        lines.append(f"| {task} | {owner} | {due} |")
    if len(lines) == 2:
        lines.append("| (없음) |  |  |")
    return "\n".join(lines) + "\n"


def main() -> None:
    data = json.loads(IN_PATH.read_text(encoding="utf-8"))
    OUT_PATH.write_text(render_markdown(data), encoding="utf-8")
    print(str(OUT_PATH))


def render_markdown(data: dict) -> str:
    """out.json dict → 마크다운 본문 (process.py / web 캡처용)."""
    title = md_escape(str(data.get("title") or "")).strip()
    if not title:
        title = "정리 노트"

    summary = md_escape(str(data.get("summary") or ""))
    categories = data.get("categories") or []
    decisions = data.get("decisions") or []
    action_items = data.get("action_items") or []
    open_questions = data.get("open_questions") or []
    references = data.get("references") or []

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    cat_line = ", ".join(str(c).strip() for c in categories if str(c).strip()) or "(없음)"

    md = [
        f"# {title}",
        "",
        f"- 생성시간: {now}",
        f"- 카테고리: {cat_line}",
        "",
        "## 요약",
        summary if summary else "(없음)",
        "",
        "## 결정사항",
        bullet_lines([str(x) for x in decisions]),
        "",
        "## 액션 아이템",
        action_items_table(list(action_items)),
        "",
        "## 열린 질문",
        bullet_lines([str(x) for x in open_questions]),
        "",
        "## 참고/링크",
        bullet_lines([str(x) for x in references]),
        "",
    ]
    return "\n".join(md)


if __name__ == "__main__":
    main()
