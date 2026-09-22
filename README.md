# Template PPT

一个面向正式汇报的 Codex Skill：根据内容大纲和指定视觉模板生成可编辑的 PowerPoint，而不是重新设计一套主题。

适用于开题答辩、毕业答辩、项目汇报、课程展示和工作汇报等场景。

## 特性

- 内容大纲是唯一的事实来源，不虚构数据、引用、成果或结论。
- 视觉模板是必填项，复用其比例、配色、字体、页眉页脚、装饰和既有版式。
- 可选结构参考 PPT 仅用于提取章节逻辑与页面节奏，不复制其中的研究或业务内容。
- 生成文本框和表格均保持可编辑，不输出图片化 PPT。
- 自动保存独立工作区：内容结构、模板分析、页面计划、构建脚本、预览图和 QA 报告均可追溯。
- 检查 PPTX 可读性、页面边界、文本重叠、字体、待补充内容，并可通过 WPS 或 PowerPoint 导出预览图。

## 安装

将本目录复制到 Codex Skills 目录：

```text
%USERPROFILE%\.codex\skills\template-ppt
```

需要的本地能力：

- Windows 上的 WPS 演示或 Microsoft PowerPoint，用于复制模板版式和导出预览；
- Python 及 `python-pptx`，用于模板分析、页面结构解析与 QA；
- 可选：已安装的 `$markitdown`，用于把非 Markdown 内容大纲转换为 Markdown。

## 使用

显式调用 Skill：

```text
$template-ppt
内容大纲：D:\\path\\to\\outline.md
视觉模板：D:\\path\\to\\template.pptx
结构参考：D:\\path\\to\\reference.pptx
输出目录：D:\\path\\to\\output
```

其中：

| 输入 | 是否必填 | 说明 |
|---|---:|---|
| 内容大纲 | 是 | 唯一内容依据；优先使用 Markdown。 |
| 视觉模板 | 是 | 唯一视觉来源，必须是 `.pptx`。 |
| 结构参考 | 否 | 用于章节叙事与页面节奏，可与视觉模板是同一文件。 |
| 输出目录 | 否 | 缺省时写入内容大纲同级的 `PPT输出`。 |

缺少内容大纲或视觉模板时，Skill 会停止并列出缺失项；不会替换为默认主题或商业模板。

## 输出内容

每次生成都会建立独立工作区：

```text
PPT输出/<大纲名称>/
├── source/                 # 输入文件副本
├── template_analysis.json  # 模板页面、文字和字体盘点
├── outline_structure.json  # 大纲层级
├── deck_plan.json          # 可复建的页面计划
├── build/                  # 构建产物
├── qa/                     # QA 报告
├── renders/                # 单页预览和联系表
└── <大纲名称>-正式汇报.pptx
```

## 质量边界

- 内容不足时使用 `【待补充：原因】`，并在 QA 报告中列出；不会擅自检索或补写事实。
- 页面内容过密、越出模板内容区或与装饰冲突时，QA 会标记问题，必须调整页面计划后重新构建。
- 不将学校模板、个人汇报内容、学生数据或内部资料提交到本仓库。

## 仓库内容

```text
SKILL.md       # Codex Skill 指令
agents/        # 显式调用策略和界面元数据
references/    # 输入契约、页面计划和 QA 规则
scripts/       # 模板分析、构建、预览与 QA 工具
```

本项目不包含任何 PPT 模板、示例汇报或个人数据。
