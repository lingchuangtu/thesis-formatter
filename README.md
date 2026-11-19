# thesis-formatter（论文模板排版 Skill）

一个 WorkBuddy / Claude Skill：按学校论文格式 Word 模板对论文做精准排版，**逐项满足字号、字体、行距、缩进、页边距、页码、页眉页脚、标题层级、图表题注、参考文献等全部硬性要求**，并生成可勾选的验收报告。

## 两条路径

| | 路径 A：克隆法 | 路径 B：移植法 |
| --- | --- | --- |
| 适用 | 纯文本初稿，无表格/图片/代码 | **已完成的论文**（含表格/图片/代码）+ **真实学校模板**（多分节/自定义样式/奇偶页眉） |
| 做法 | 从模板 `.docx` 整包克隆，只替换正文段落 | 以论文原文件为基线，把模板样式/页面设置/分节/页眉页脚移植进来 |
| 保证 | 字号、页码、页眉页脚与模板零偏差 | 内容零丢失的前提下格式达标（含标题重编号、封面重建） |
| 脚本 | `build_document.py` | `transform_paper.py` + `rebuild_cover.py` |

两条路径都通过 `verify_requirements.py` 验收，路径 B 额外用 `audit_content.py` 证明内容零丢失。页码 `PAGE/NUMPAGES` 始终保持为域而非静态文字。复杂场景的方法论见 `references/real-world-workflow.md`。

## 安装

将本目录放到：

- WorkBuddy 用户级：`~/.workbuddy/skills/thesis-formatter/`
- 项目级：`.workbuddy/skills/thesis-formatter/`

无需 pip 依赖，Python 3.8+ 即可（仅用标准库 `zipfile` + `xml.etree`）。

## 测试

一键自测（生成演示模板，跑通解析→排版→验收全流程，并做负向测试确认验收器能拦截错误）：

```bash
python tests/selftest.py
```

预期输出：`[5/5] 全部通过`，正向验收 `FAIL=0`，负向测试 `FAIL=1`（正确拦截）。

真实测试：提供你自己的学校 `.docx` 模板与论文文本，让 WorkBuddy 按 SKILL.md 工作流执行，检查生成的 `report.md` 中 FAIL 是否为 0。

## 使用

1. 提供学校 `.docx` 格式模板与论文（纯文本走路径 A，完整论文走路径 B）。
2. Skill 判断路径并执行。
3. 交付可编辑 `.docx` + 验收报告。

路径 B 流程（真实论文）：

```bash
python scripts/inspect_template.py <模板> -o spec.json --summary summary.md   # 诊断
python scripts/transform_paper.py --print-template                            # 看配置骨架
# 按诊断填好 transform 配置（参考 assets/transform-config.example.json）
python scripts/transform_paper.py <配置.json>                                # 格式移植
python scripts/rebuild_cover.py <无封面.docx> <模板.docx> <带封面.docx>       # 封面必做，零配置
python scripts/audit_content.py <原论文> <成品> --normalize-headings         # 内容零丢失
python scripts/verify_requirements.py <成品> --spec spec.json --report report.md
```

## 脚本

| 脚本 | 作用 |
| --- | --- |
| `scripts/inspect_template.py` | 解析模板样式/页面设置/页眉页脚，生成 spec 与摘要 |
| `scripts/build_document.py` | 路径A：克隆模板整包，替换正文，保留全部格式定义 |
| `scripts/transform_paper.py` | 路径B：配置驱动的格式移植（样式/分节/页眉页脚/标题编号），含结构自检 |
| `scripts/rebuild_cover.py` | 路径B 必做：零配置重建封面（自动识别封面/题目/校徽，只替换题目，个人信息清空留空，独立节） |
| `scripts/audit_content.py` | 内容零丢失核验（文字/图片/表格/绘图/XML） |
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
│   ├── build_document.py          # 路径A
│   ├── transform_paper.py         # 路径B
│   ├── rebuild_cover.py           # 路径B
│   ├── audit_content.py
│   └── verify_requirements.py
├── tests/
│   └── selftest.py
├── references/
│   ├── format-checklist.md
│   └── real-world-workflow.md     # 真实模板/复杂论文方法论
└── assets/
    ├── content.example.json
    ├── requirements.example.json
    ├── transform-config.example.json
    └── cover-config.example.json
```

## 免责声明

仅提供排版格式辅助，不替代学术判断，不保证查重/评审/答辩结果；内容真实性由使用者负责。与学校最新要求不一致时以学校要求为准。
