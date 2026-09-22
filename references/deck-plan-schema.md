# 页面计划格式

`deck_plan.json` 是唯一的可复建页面描述。它由输入大纲和模板分析生成，不保存到 Skill 内。

```json
{
  "title": "汇报标题",
  "template_path": "绝对模板路径",
  "slides": [
    {
      "template_slide": 1,
      "clear_template_text": true,
      "preserve_text": ["校徽", "页码"],
      "blocks": [
        {"kind": "text", "role": "title", "text": "标题", "x": 1.2, "y": 1.0, "w": 10.8, "h": 0.7, "font_size": 28, "bold": true, "align": "center"},
        {"kind": "text", "role": "body", "text": "正文", "x": 1.5, "y": 2.0, "w": 10.2, "h": 2.5, "font_size": 18, "align": "left"},
        {"kind": "table", "role": "table", "x": 1.4, "y": 2.0, "w": 10.4, "h": 3.1, "headers": ["项目", "说明"], "rows": [["A", "B"]]}
      ]
    }
  ]
}
```

- Coordinates are inches and must stay inside `template_analysis.json` canvas.
- `template_slide` is a 1-based index from the supplied visual template.
- Use `clear_template_text: true` for a copied sample/content slide. Keep it
  false for a cover/chapter slide whose template text should remain.
- `preserve_text` contains literal fragments that must not be cleared, such as
  a logo wordmark or department name.
- Prefer one title and one to three content blocks per slide. Split dense
  content into additional pages rather than using text smaller than the
  template's smallest normal body font.
