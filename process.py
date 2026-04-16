import json
import re
from pathlib import Path

import ollama

MODEL = "llama3"
INPUT_PATH = Path(__file__).resolve().parent / "sample.txt"
OUTPUT_PATH = Path(__file__).resolve().parent / "out.json"

SYSTEM_PROMPT = """너는 비정형 대화/메모를 구조화하는 도우미다.
사용자가 준 텍스트만 보고 아래 JSON 스키마에 맞춰 정리해라.
- title, summary, categories, decisions, action_items의 task/owner/due, open_questions, references 안의 모든 문자열은 한국어로 작성한다.
- 빈 항목은 빈 문자열 "" 또는 빈 배열 []로 둔다.
- title은 반드시 한국어 한 줄 제목이며 빈 문자열로 두지 않는다.
- decisions는 문자열만 담은 배열이다. 객체를 넣지 마라. (할 일은 action_items에만 넣는다.)
- 추측은 하지 말고, 텍스트에 근거해 요약한다.
- 응답은 JSON만 출력한다. 앞뒤로 설명 문장, 마크다운 코드펜스를 넣지 마라.

스키마:
{
  "title": "한 줄 제목",
  "summary": "전체 요약 (2~4문장)",
  "categories": ["업무", "과제" 등 키워드 배열],
  "decisions": ["합의·확정 내용을 문자열로만 나열"],
  "action_items": [{"task": "할 일", "owner": "담당자 또는 빈 문자열", "due": "기한 또는 빈 문자열"}],
  "open_questions": ["아직 불명확한 점"],
  "references": ["파일경로, 티켓명 등 텍스트에 나온 참조"]
}
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


def main() -> None:
    raw = INPUT_PATH.read_text(encoding="utf-8")
    response = ollama.chat(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": raw},
        ],
        options={"temperature": 0},
    )
    content = response["message"]["content"]
    data = _parse_json_strict(content)
    OUTPUT_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
