import json
import re
import argparse
import sys
from pathlib import Path

import ollama

MODEL = "llama3"
INPUT_PATH = Path(__file__).resolve().parent / "sample.txt"
OUTPUT_PATH = Path(__file__).resolve().parent / "out.json"
SECTION_CANDIDATES = ["cos 개발", "ax 과제", "문의", "to-do"]

# 예전 프롬프트/모델 출력 호환
_SECTION_ALIASES = {
    "본사과제": "ax 과제",
    "이슈 및 오류": "cos 개발",
}


def normalize_recommended_section(value: object) -> str:
    s = str(value or "").strip()
    s = _SECTION_ALIASES.get(s, s)
    if s not in SECTION_CANDIDATES:
        return "ax 과제"
    return s

SYSTEM_PROMPT = """너는 비정형 대화/메모를 구조화하는 도우미다.
사용자가 준 텍스트만 보고 아래 JSON 스키마에 맞춰 정리해라.

규칙:
- 출력은 JSON만. 설명 문장/마크다운 코드펜스/불릿 등을 JSON 바깥에 쓰지 마라.
- 모든 문자열은 한국어로 작성한다. (영어 금지)
- 추측하지 말고, 텍스트에 근거해서만 쓴다.
- 빈 항목은 빈 문자열 "" 또는 빈 배열 []로 둔다.
- title은 반드시 비어 있지 않은 한국어 한 줄 제목이다. (최소 5자) 애매하면 "업무 메모 정리"처럼 일반 제목을 쓴다.
- summary는 반드시 비어 있지 않게 1~4문장으로 쓴다. 정보가 적어도 "무엇에 대한 메모인지"를 한 문장으로 정리한다.
- categories는 아래 후보 중에서만 1~3개 고른다.
  후보: ["업무", "과제", "일정", "회의", "기술", "개인", "이슈", "문의", "기타"]
- recommended_section은 원노트에서 넣을 "탭(섹션) 이름" 1개를 추천한다.
  - 아래 후보 중에서만 반드시 1개를 고른다. (다른 문자열 금지, 철자·하이픈 그대로)
  - 후보: ["cos 개발", "ax 과제", "문의", "to-do"]
  - cos 개발: 코드·API·COS·SSG·테스트·빌드·버그/로그/재현 등 구현·기술 중심.
  - ax 과제: 역공학·과제정의서·TASK·중요단말·본사·전사 과제·회의·결정 등 내부 AX 과제·기획.
  - 문의: 외부·타팀에 묻는 질문·요청만.
  - to-do: 개인 할 일·체크리스트·마감만 있는 짧은 실행 목록(과제 본문 설명이 아닐 때).
- decisions는 문자열 배열만 담는다. 객체를 넣지 마라. (할 일은 action_items에만 넣는다.)

스키마:
{
  "title": "한 줄 제목",
  "summary": "전체 요약 (2~4문장)",
  "categories": ["업무", "과제" 등 후보에서 선택],
  "recommended_section": "추천 탭(섹션) 이름 1개",
  "decisions": ["합의·확정 내용을 문자열로만 나열"],
  "action_items": [{"task": "할 일", "owner": "담당자 또는 빈 문자열", "due": "기한 또는 빈 문자열"}],
  "open_questions": ["아직 불명확한 점"],
  "references": ["파일경로, 티켓명 등 텍스트에 나온 참조"]
}
"""

SECTION_CLASSIFIER_PROMPT = f"""너는 사용자가 준 텍스트를 원노트 탭(섹션)으로 분류하는 분류기다.

규칙:
- 출력은 후보 중 1개 문자열만 출력한다. (따옴표/JSON/설명 금지)
- 애매하면 기본값은 "ax 과제"로 한다.
- 후보: {SECTION_CANDIDATES}
- cos 개발: 코드·API·COS·Spring REST Docs·SSG·테스트·빌드·버그 재현·로그 등 기술 구현 중심.
- ax 과제: 역공학·기술자산·과제정의서·TASK·중요단말·본사·전사·회의·결정 등 AX 과제·기획 본문.
- 문의: 외부·타팀에 보내는 질문·요청만.
- to-do: 실행 항목만 나열·체크리스트·짧은 할 일(과제 설명·기술 스펙이 주가 아닐 때). 영어 표기는 정확히 to-do (하이픈 포함).
"""


def _strip_code_fence(text: str) -> str:
    text = text.strip()
    m = re.match(r"^```(?:json)?\s*\n?(.*)\n?```\s*$", text, re.DOTALL | re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return text


def _first_json_object(text: str) -> str:
    """모델이 앞뒤에 한 줄 설명을 붙인 경우 첫 번째 {...}만 잘라낸다."""
    start = text.find("{")
    if start == -1:
        raise ValueError("JSON 객체 시작 '{'를 찾을 수 없습니다.")
    depth = 0
    for i in range(start, len(text)):
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
            if depth == 0:
                return text[start : i + 1]
    raise ValueError("닫히지 않은 JSON 객체입니다.")


def _parse_json_strict(text: str) -> dict:
    cleaned = _strip_code_fence(text.strip())
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        return json.loads(_first_json_object(cleaned))


def classify_section_llm(*, raw: str, model: str = MODEL) -> str:
    response = ollama.chat(
        model=model,
        messages=[
            {"role": "system", "content": SECTION_CLASSIFIER_PROMPT},
            {"role": "user", "content": raw},
        ],
        options={"temperature": 0},
    )
    section = str(response["message"]["content"]).strip()
    return normalize_recommended_section(section)


def process_text(*, raw: str, model: str = MODEL, temperature: float = 0) -> dict:
    response = ollama.chat(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": raw},
        ],
        options={"temperature": temperature},
    )
    content = response["message"]["content"]
    data = _parse_json_strict(content)
    if isinstance(data, dict):
        data["recommended_section"] = normalize_recommended_section(data.get("recommended_section"))
    return data


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument(
        "--in",
        dest="in_path",
        default=str(INPUT_PATH),
        help="입력 텍스트 파일 경로. '-' 이면 stdin에서 읽습니다.",
    )
    p.add_argument(
        "--out",
        dest="out_path",
        default=str(OUTPUT_PATH),
        help="출력 JSON 파일 경로",
    )
    p.add_argument("--model", default=MODEL)
    p.add_argument("--temperature", type=float, default=0)
    args = p.parse_args()

    if args.in_path == "-":
        raw = sys.stdin.read()
    else:
        raw = Path(args.in_path).read_text(encoding="utf-8")

    data = process_text(raw=raw, model=args.model, temperature=args.temperature)
    Path(args.out_path).write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
