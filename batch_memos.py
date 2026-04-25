#!/usr/bin/env python3
"""
하루치 등 여러 메모를 한 번에 구조화(out.json 스키마 + .md).

  # 폴더 안의 각 .txt 파일 = 메모 1개 (이름 순)
  python batch_memos.py --dir ./memos_today

  # 한 파일 안에서 줄 단위 구분자로 여러 메모 (기본 구분: ---memo---)
  python batch_memos.py --file ./day_dump.txt --split

  # 구분 문자열 바꾸기
  python batch_memos.py --file dump.txt --split --delimiter "===CUT==="

출력: --out-dir (기본 ./outputs) 아래에 memo_001.json, memo_001.md …
원노트 업로드는 각 결과를 web_capture 로 보내거나, 추후 스크립트로 확장.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import process
import render_md


def _slug(title: str, index: int) -> str:
    t = (title or "").strip()
    t = re.sub(r"[^\w\s\-가-힣]+", " ", t, flags=re.UNICODE)
    t = re.sub(r"\s+", "_", t).strip("_")
    if len(t) < 2:
        return f"memo_{index:03d}"
    return t[:100] if len(t) > 100 else t


def _chunks_from_file(path: Path, delimiter: str, split: bool) -> list[str]:
    raw = path.read_text(encoding="utf-8")
    if not split:
        return [raw]
    d = (delimiter or "").strip()
    if not d:
        raise SystemExit("--delimiter 가 비어 있습니다.")
    # 구분 문자열이 단독 한 줄(앞뒤 공백 허용)일 때만 잘라냄
    pattern = re.compile(rf"(?:^|\n)\s*{re.escape(d)}\s*(?:\n|$)", re.MULTILINE)
    parts = pattern.split(raw)
    return [p.strip() for p in parts if p.strip()]


def _chunks_from_dir(dir_path: Path) -> list[tuple[str, str]]:
    """(파일 stem, 본문) 리스트."""
    txts = sorted(dir_path.glob("*.txt"))
    if not txts:
        raise SystemExit(f"{dir_path} 안에 .txt 가 없습니다.")
    out: list[tuple[str, str]] = []
    for p in txts:
        body = p.read_text(encoding="utf-8")
        if body.strip():
            out.append((p.stem, body))
    if not out:
        raise SystemExit(f"{dir_path} 의 .txt 가 모두 비어 있습니다.")
    return out


def main() -> None:
    p = argparse.ArgumentParser(description="여러 메모 배치 구조화 (Ollama)")
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--dir", type=Path, metavar="DIR", help="폴더 내 각 .txt = 메모 1개")
    g.add_argument("--file", type=Path, metavar="FILE", help="단일 텍스트 파일")
    p.add_argument(
        "--split",
        action="store_true",
        help="--file 과 함께: 본문을 --delimiter 줄로 나눔",
    )
    p.add_argument(
        "--delimiter",
        default="---memo---",
        help="--split 일 때만 사용. 이 문자열이 단독 한 줄이면 경계(앞뒤 공백 무시).",
    )
    p.add_argument(
        "--out-dir",
        type=Path,
        default=Path("outputs"),
        help="결과 저장 폴더 (없으면 생성)",
    )
    p.add_argument("--model", default=process.MODEL)
    p.add_argument("--temperature", type=float, default=0)
    args = p.parse_args()

    if args.split and not args.file:
        p.error("--split 은 --file 과 함께만 사용할 수 있습니다.")
    if args.split and args.dir:
        p.error("--split 은 단일 --file 전용입니다. 폴더는 파일마다 메모 1개입니다.")

    if args.file and not args.split:
        items: list[tuple[str, str]] = [(args.file.stem, args.file.read_text(encoding="utf-8"))]
    elif args.file and args.split:
        chunks = _chunks_from_file(args.file, args.delimiter, split=True)
        items = [(f"part_{i+1:03d}", c) for i, c in enumerate(chunks)]
    else:
        assert args.dir
        items = _chunks_from_dir(args.dir)

    out_dir: Path = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    batch_dir = out_dir / f"batch_{stamp}"
    batch_dir.mkdir(parents=True, exist_ok=True)

    used: set[str] = set()
    for i, (label, raw) in enumerate(items, start=1):
        raw = raw.strip()
        if not raw:
            continue
        print(f"[{i}/{len(items)}] 처리 중… ({label})", file=sys.stderr, flush=True)
        data = process.process_text(raw=raw, model=args.model, temperature=args.temperature)
        title = str(data.get("title") or label)
        base = _slug(title, i)
        name = base
        n = 1
        while name in used:
            n += 1
            name = f"{base}_{n}"
        used.add(name)

        jpath = batch_dir / f"{name}.json"
        mpath = batch_dir / f"{name}.md"
        jpath.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        mpath.write_text(render_md.render_markdown(data), encoding="utf-8")
        print(str(jpath))
        print(str(mpath))

    print(f"완료: {batch_dir}", file=sys.stderr)


if __name__ == "__main__":
    main()
