#!/usr/bin/env python3
"""Create a conservative, editable first-pass deck plan from Markdown headings."""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def walk(nodes):
    for node in nodes:
        yield node
        yield from walk(node.get("children", []))


def concise(node):
    lines = list(node.get("content", []))
    if not lines:
        return "【待补充：此页的正文内容】"
    body = "\n".join(lines)
    return body[:360] + ("…" if len(body) > 360 else "")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--outline-structure", required=True)
    parser.add_argument("--template-analysis", required=True)
    parser.add_argument("--template-path", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    structure = json.loads(Path(args.outline_structure).read_text(encoding="utf-8"))
    template = json.loads(Path(args.template_analysis).read_text(encoding="utf-8"))
    width = float(template["canvas"]["width_in"])
    height = float(template["canvas"]["height_in"])
    count = int(template["slide_count"])
    if count < 1:
        raise SystemExit("Visual template has no slides.")
    nodes = list(walk(structure.get("sections", [])))
    title = nodes[0]["title"] if nodes else Path(structure["input"]).stem
    slides = [{
        "template_slide": 1,
        "clear_template_text": True,
        "blocks": [
            {"kind": "text", "role": "title", "text": title, "x": width * 0.14, "y": height * 0.31, "w": width * 0.72, "h": height * 0.17, "font_size": 28, "bold": True, "align": "center"},
            {"kind": "text", "role": "subtitle", "text": "【待补充：汇报人、单位、日期】", "x": width * 0.18, "y": height * 0.58, "w": width * 0.64, "h": height * 0.07, "font_size": 14, "align": "center"},
        ],
    }]
    content_templates = list(range(2, count + 1)) or [1]
    for index, node in enumerate(nodes):
        template_slide = content_templates[index % len(content_templates)]
        slides.append({
            "template_slide": template_slide,
            "clear_template_text": True,
            "blocks": [
                {"kind": "text", "role": "title", "text": node["title"], "x": width * 0.10, "y": height * 0.08, "w": width * 0.80, "h": height * 0.10, "font_size": 25, "bold": True, "align": "left"},
                {"kind": "text", "role": "body", "text": concise(node), "x": width * 0.12, "y": height * 0.24, "w": width * 0.76, "h": height * 0.53, "font_size": 16, "align": "left"},
            ],
        })
    plan = {"title": title, "template_path": str(Path(args.template_path).resolve()), "generated_from": str(Path(structure["input"]).resolve()), "slides": slides}
    destination = Path(args.out)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    print(destination)


if __name__ == "__main__":
    main()
