#!/usr/bin/env python3
"""Extract Markdown heading hierarchy and nearby content into JSON."""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

HEADING = re.compile(r"^(#{1,6})\s+(.+?)\s*$")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    source = Path(args.input)
    text = source.read_text(encoding="utf-8-sig")
    nodes, stack = [], []
    for line_number, raw in enumerate(text.splitlines(), start=1):
        match = HEADING.match(raw)
        if match:
            level, title = len(match.group(1)), match.group(2).strip()
            node = {"level": level, "title": title, "line": line_number, "content": [], "children": []}
            while stack and stack[-1]["level"] >= level:
                stack.pop()
            if stack:
                stack[-1]["children"].append(node)
            else:
                nodes.append(node)
            stack.append(node)
        elif stack and raw.strip():
            stack[-1]["content"].append(raw.strip())
    result = {"input": str(source.resolve()), "sections": nodes}
    destination = Path(args.out)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(destination)


if __name__ == "__main__":
    main()
