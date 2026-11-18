---
name: thesis-formatter
description: 按学校论文格式 Word 模板对论文文本/初稿做精准排版，逐项满足字号、字体、行距、缩进、页边距、页码、页眉页脚、标题层级、图表题注、参考文献等全部硬性要求，并生成可勾选的验收报告。适用于“按这个学校模板排论文”“把论文按模板调整格式”“页眉页脚页码字号都要符合模板”等任务。当用户提供 .doc/.docx 论文格式模板与论文文本并要求据此调整格式时必须使用本 skill。This skill should be used when a user provides a school thesis Word format template plus thesis text and asks to apply the template's exact formatting — including font sizes, page numbering, headers and footers — to produce a compliant .docx.
agent_created: true
---

# 论文模板排版（Thesis Formatter）

## 用途

把用户提供的**学校论文格式模板（.docx）**作为唯一版式来源，对**论文文本/初稿**做排版，产出格式逐项达标的 .docx，并附验收报告。核心保证：**字号、页码、页眉页脚等全部要求都达到**，因为交付文档是从模板原包克隆而来，只替换正文内容，样式、页面设置、页眉页脚、页码域原样保留。

## 触发场景

- “按学校模板给论文排版”“把论文按这个模板调整格式”
- 上传 .docx 格式模板 + 论文文本/初稿，要求字号、页眉页脚、页码全部符合模板
- 明确强调“所有格式要求都必须达到”

## 所需输入

| 输入 | 必需 | 说明 |
| --- | --- | --- |
| 学校论文格式模板（.docx） | 是 | 唯一版式基线；.doc 先另存为 .docx |
| 论文文本/初稿 | 是 | .docx/.md/.txt 或对话内文本 |
| 硬性附加要求 | 否 | 如“正文宋体小四、页脚居中页码从 1 开始” |

缺少模板或论文文本时一次性请求补充；齐全则不提问，直接执行。

## 工作流

### 步骤 1：解析模板，建立版式基线

- 运行 `python scripts/inspect_template.py <模板.docx> -o spec.json --summary summary.md`。
- 阅读 `summary.md`，确认：页面设置（页边距/纸张/页码格式与起始/首页不同）、样式表（各样式字体字号行距缩进）、页眉页脚文字与域（PAGE/NUMPAGES）。
- 阅读 `references/format-checklist.md`，记住全部硬性项。
- 确认模板里的样式角色：正文、一级/二级/三级标题、摘要、参考文献、图表题注分别对应哪个样式 ID（summary 第 3 节有自动推测）。

### 步骤 2：整理论文文本为内容模型

- 把论文文本按段落拆成 `content.json`（格式见 `assets/content.example.json`）。
- 每段用 `style` 指定模板中的样式 ID（如 `Heading1`、`Normal`、`Caption`、`Bibliography`）；未指定则继承模板默认样式。
- 章节标题务必用对应 Heading 样式并设置 `keep_next: true`，保证大纲级别正确、标题不与正文分离。
- 需要分页处用 `page_break` 或 `page_break_before`；封面后另起页用 `page_break_before`。
- 若用户另给了硬性文字要求（如正文字号），写入 `requirements.json`（格式见 `assets/requirements.example.json`）；仅写需要额外校验或覆盖的项，其余以模板为准。
- **绝不编造论文内容**；原文缺失处用明确占位并列入交付说明，不用虚构内容填充。

### 步骤 3：克隆模板并替换正文

- 运行 `python scripts/build_document.py <模板.docx> content.json <输出.docx>`。
- 该脚本复制模板整包（样式、编号、页眉页脚、页面设置、主题、字体原样保留），仅替换 `<w:body>` 中的正文段落，保留所有 `sectPr`。这一步从机制上保证字号、页码、页眉页脚与模板一致。
- 若样式 ID 与模板不一致，用 `--stylesheet-map roles.json` 做角色名→真实样式 ID 映射。

### 步骤 4：逐项验收（必须通过）

- 运行 `python scripts/verify_requirements.py <输出.docx> --spec spec.json --requirements requirements.json --report report.md`。
- 打开 `report.md`，要求 **FAIL=0**。任何 FAIL 回到步骤 2/3 修正后重新验证，直到全部通过。
- 打开生成文档目视复核：页眉页脚显示、页码连续、章节标题层级、目录页码、图表题注、参考文献格式。
- 逐项对照 `references/format-checklist.md` 勾选；A/B/C 类为硬性项，不得遗漏。

### 步骤 5：交付

- 提供可下载、可编辑的 .docx 与 `report.md`。
- 交付说明简短：使用的模板、验收结果（FAIL 数）、需人工复核的少量位置（如目录页码需在 Word 中按 F9 更新、占位符）。

## 红线

- 不得新建空白文档拼版；必须从模板原包克隆，否则无法保证页眉页脚/页码/字号零损。
- 不得跳过 `verify_requirements.py`；FAIL 未清零不得交付。
- 不得篡改页眉页脚中的域为静态文字；PAGE / NUMPAGES 必须保持为域。
- 不得编造论文内容、数据或参考文献。
- 不得修改用户提供的原始模板文件。

## 脚本依赖

全部脚本仅用 Python 标准库（zipfile + xml.etree），无需 pip 安装；Python 3.8+ 可用。

## 资源

- `scripts/inspect_template.py`：解析模板 → spec.json + 人类可读摘要。
- `scripts/build_document.py`：克隆模板 + 替换正文 → 输出 docx。
- `scripts/verify_requirements.py`：逐项验收 → report.md（FAIL 非零退出）。
- `scripts/docx_utils.py`：共享 OOXML 读写工具（无第三方依赖）。
- `references/format-checklist.md`：全部格式硬性要求清单。
- `assets/content.example.json`：内容模型示例。
- `assets/requirements.example.json`：额外硬性要求示例。
- `tests/selftest.py`：一键自测（正向 FAIL=0 + 负向拦截验证）。

## 免责声明

本 Skill 仅提供排版格式辅助，不替代学术判断，不保证查重、评审或答辩结果；内容真实性、原创性与学术规范由使用者负责。如与学校最新要求不一致，以学校要求为准。
