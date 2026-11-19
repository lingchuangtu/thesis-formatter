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

## 第一步：判断走哪条路径（关键，必做）

先看论文形态，选错路径会丢内容：

| 论文形态 | 路径 | 核心脚本 |
| --- | --- | --- |
| 纯文本/简单初稿，**无表格/图片/代码** | **A. 克隆法** | `build_document.py` |
| 已完成的论文，**含表格/图片/代码**，或模板为多分节+自定义样式 | **B. 移植法** | `transform_paper.py`（+ `rebuild_cover.py`） |

判断方法：检查论文 docx 是否含 `word/media/`、`w:tbl`、`w:drawing`；模板 `styles.xml` 是否有非标准样式名（如 `CMheading1`）、`document.xml` 是否有多个 `w:sectPr`。**任一命中即走路径 B。**

- 路径 B 的完整方法、诊断清单、避坑经验见 `references/real-world-workflow.md`，配置示例见 `assets/transform-config.example.json` 和 `assets/cover-config.example.json`。
- 下面步骤 1~5 是**路径 A（克隆法）**；路径 B 按 real-world-workflow.md 执行「诊断 → 配置 → transform_paper → **rebuild_cover（必做）** → 验证三件套」。封面默认必做，见步骤 4。

## 工作流（路径 A：克隆法）

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

## 路径 B：移植法速查（真实论文）

1. **诊断**（只读，不写代码）：按 `references/real-world-workflow.md` 第二节摸清模板分节/样式/页码和论文段落/标题编号/内容指纹。
2. **配置**：`python scripts/transform_paper.py --print-template` 出配置骨架，按诊断填 `assets/transform-config.example.json`。
3. **移植**：`python scripts/transform_paper.py <配置.json>` —— 样式移植 + 清直接格式 + 标题重编号 + 分节重建 + 页眉页脚移植，输出时自带结构零丢失自检。
4. **封面（默认必做，零配置）**：`python scripts/rebuild_cover.py <第3步成品.docx> <学校模板.docx> <带封面成品.docx>`。
   - 脚本自动识别模板封面范围、取论文题目、搬迁校徽、读取封面页边距；**只替换题目，个人信息（姓名/学号/导师/专业/学院/日期）一律清空只留标签**——不强制用户提供隐私信息。
   - 仅当自动识别不适用时才用配置覆盖：`python scripts/rebuild_cover.py <config.json>`（骨架 `--print-template`）。
   - 判定封面是否成功：第一页应有校徽 + "毕业论文" + 题目（新）+ 空的下划线字段；若仍是模板示例题目/示例姓名，说明漏跑了这一步。
5. **验证三件套**：
   - `python scripts/audit_content.py <原论文> <成品> --normalize-headings` —— 内容零丢失（全 PASS）。
   - `python scripts/inspect_template.py <模板> -o spec.json` + `python scripts/verify_requirements.py <成品> --spec spec.json --report report.md` —— 格式达标（区分真实缺陷与示例文字/自定义样式误报，见参考文档第五节）。
   - 目视复核封面、页码起算、奇偶页眉、目录 F9 更新。

## 红线

- **封面默认必做**：路径 B 排版后必须跑 `rebuild_cover.py` 生成带封面的第一页；封面个人信息只清空、不向用户索取（属隐私），题目必须替换为论文题目。不得交付"页面仍是模板示例题目/示例姓名"的文档。
- **含表格/图片/代码的论文绝不用克隆法**，必须走移植法，否则丢图丢表。
- 不得新建空白文档拼版；路径 A 必须从模板原包克隆，路径 B 必须以论文原文件为基线。
- 不得跳过内容零丢失验证；`audit_content.py` 非全 PASS 不得交付。
- 不得跳过 `verify_requirements.py`；FAIL 未清零（或误报未说明）不得交付。
- 不得篡改页眉页脚中的域为静态文字；PAGE / NUMPAGES 必须保持为域。
- 标题/目录编号必须与模板统一；改目录时只改章号文字，页码域与制表位不动。
- 不得编造论文内容、数据或参考文献。
- 不得修改用户提供的原始模板文件。

## 脚本依赖

全部脚本仅用 Python 标准库（zipfile + xml.etree），无需 pip 安装；Python 3.8+ 可用。

## 资源

- `scripts/inspect_template.py`：解析模板 → spec.json + 人类可读摘要。
- `scripts/build_document.py`：**路径A** 克隆模板 + 替换正文 → 输出 docx。
- `scripts/transform_paper.py`：**路径B** 配置驱动的格式移植（样式/分节/页眉页脚/标题编号）→ 输出 docx，含结构自检。
- `scripts/rebuild_cover.py`：**路径B 必做** 零配置重建封面（自动识别封面/题目/校徽，只替换题目，个人信息清空留空，独立节）。用法：`rebuild_cover.py <无封面.docx> <模板.docx> <输出.docx>`。
- `scripts/audit_content.py`：内容零丢失核验（文字哈希/图片/表格/绘图/XML），`--normalize-headings` 容忍编号规范化。
- `scripts/verify_requirements.py`：逐项验收 → report.md（FAIL 非零退出）。
- `scripts/docx_utils.py`：共享 OOXML 读写工具（无第三方依赖）。
- `references/format-checklist.md`：全部格式硬性要求清单。
- `references/real-world-workflow.md`：**真实模板/复杂论文完整方法论**（诊断、五步移植、验证三件套、10 条避坑）。
- `assets/content.example.json`：路径A 内容模型示例。
- `assets/requirements.example.json`：额外硬性要求示例。
- `assets/transform-config.example.json`：路径B 移植配置示例。
- `assets/cover-config.example.json`：封面**可选覆盖**配置（零配置命令通常不需要）。
- `tests/selftest.py`：一键自测（正向 FAIL=0 + 负向拦截验证）。

## 免责声明

本 Skill 仅提供排版格式辅助，不替代学术判断，不保证查重、评审或答辩结果；内容真实性、原创性与学术规范由使用者负责。如与学校最新要求不一致，以学校要求为准。
