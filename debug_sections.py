#!/usr/bin/env python3
"""
입력 텍스트를 web_capture 룰로 분류하고, 어떤 키워드로 매칭됐는지 보여준다.

사용:
  python debug_sections.py --text "..."
  python debug_sections.py --file sample.txt
"""

from __future__ import annotations

import argparse
from pathlib import Path

import web_capture


def main() -> None:
    p = argparse.ArgumentParser()
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--text")
    g.add_argument("--file", type=Path)
    args = p.parse_args()

    if args.file:
        text = args.file.read_text(encoding="utf-8")
    else:
        text = args.text or ""

    sec, reason = web_capture.diagnose_section(text)
    print("section:", sec)
    print("reason :", reason)


if __name__ == "__main__":
    main()

