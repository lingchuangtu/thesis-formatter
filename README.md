# thesis-formatter（论文模板排版 Skill）

一个 WorkBuddy / Claude Skill：按学校论文格式 Word 模板对论文做精准排版，**逐项满足字号、字体、行距、缩进、页边距、页码、页眉页脚、标题层级、图表题注、参考文献等全部硬性要求**，并生成可勾选的验收报告。

## 核心思路

交付文档从模板 `.docx` **整包克隆**而来——样式、编号、页眉页脚、页面设置、主题、字体原样保留，仅替换正文段落（`<w:body>` 内容并保留所有 `sectPr`）。页码 `PAGE/NUMPAGES` 保持为域而非静态文字。因此字号、页码、页眉页脚与模板零偏差，并通过 `verify_requirements.py` 逐项验收，FAIL 必须清零才交付。

## 安装

将本目录放到：

- WorkBuddy 用户级：`~/.workbuddy/skills/thesis-formatter/`
- 项目级：`.workbuddy/skills/thesis-formatter/`

无需 pip 依赖，Python 3.8+ 即可（仅用标准库 `zipfile` + `xml.etree`）。

## 使用

1. 提供学校 `.docx` 格式模板与论文文本/初稿。
2. Skill 自动执行：
   - `inspect_template.py` 解析模板 → `spec.json` + `summary.md`
   - 整理论文文本为 `content.json`（每段指定模板样式 ID）
   - `build_document.py` 克隆模板并替换正文 → 输出 `.docx`
   - `verify_requirements.py` 逐项验收 → `report.md`（FAIL 必须为 0）
3. 交付可编辑 `.docx` + 验收报告。

## 脚本

| 脚本 | 作用 |
| --- | --- |
| `scripts/inspect_template.py` | 解析模板模板样式/页面设置/页眉页脚，生成 spec 与摘要 |
| `scripts/build_document.py` | 克隆模板整包，替换正文，保留全部格式定义 |
| `scripts/verify_requirements.py` | 对照 spec 与 requirements 逐项验收，输出报告与退出码 |
| `scripts/docx_utils.py` | 共享 OOXML 读写工具 |

## 文件结构

```
thesis-formatter/
├── SKILL.md
├── README.md
├── scripts/
│   ├── docx_utils.py
│   ├── inspect_template.py
│   ├── build_document.py
│   └── verify_requirements.py
├── references/
│   └── format-checklist.md
└── assets/
    ├── content.example.json
    └── requirements.example.json
```

## 免责声明

仅提供排版格式辅助，不替代学术判断，不保证查重/评审/答辩结果；内容真实性由使用者负责。与学校最新要求不一致时以学校要求为准。
