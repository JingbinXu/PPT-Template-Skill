# 交付质量检查

Run all available checks after every build:

1. `qa_ppt.py` checks PPTX readability, slide count, object bounds, likely
   object overlap, font inventory, source-text extraction, and `【待补充】`
   markers.
2. `Render-PptPreview.ps1` exports every page to `renders/` and creates a
   contact sheet when WPS/PowerPoint automation is available.
3. Visually inspect the contact sheet and any flagged pages. If render export
   is unavailable, report that as a limitation rather than claiming review.

Warnings are acceptable only for deliberate overlays (for example, a title on
a background shape) or explicit `【待补充】` blocks. The QA report must list both.
Any object outside the canvas, an unreadable PPTX, or unexpected residual
template sample text requires rebuilding from `deck_plan.json`.
