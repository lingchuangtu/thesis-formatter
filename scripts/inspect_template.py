#!/usr/bin/env python3
"""Inspect a Word template (.docx) and emit a machine-readable format spec.

Usage:
    python inspect_template.py <template.docx> [-o spec.json] [--summary summary.md]

The spec captures styles (fonts/sizes/spacing/indents), section page setup
(margins, page size, page-number format/start) and header/footer content
including PAGE/NUMPAGES fields.  No third-party dependencies.
"""

from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_utils import (build_spec, read_docx, size_name)  # noqa: E402


ROLE_KEYWORDS = {
    "heading1": ["heading 1", "标题 1", "标题1", "一级标题", "章标题"],
    "heading2": ["heading 2", "标题 2", "标题2", "二级标题", "节标题"],
    "heading3": ["heading 3", "标题 3", "标题3", "三级标题"],
    "body": ["normal", "正文", "body text", "正文文本"],
    "caption": ["caption", "题注", "图题", "表题"],
    "reference": ["bibliography", "reference", "参考文献"],
    "toc": ["toc ", "目录"],
    "abstract": ["abstract", "摘要"],
}


def guess_roles(styles: dict) -> dict:
    roles = {k: [] for k in ROLE_KEYWORDS}
    for sid, info in styles.items():
        if sid.startswith("__"):
            continue
        name = (info.get("name") or sid).lower()
        for role, keys in ROLE_KEYWORDS.items():
            if any(k in name or k in sid.lower() for k in keys):
                roles[role].append(sid)
                break
    return roles


def human_summary(spec: dict) -> str:
    lines = ["# 模板格式解析摘要", ""]
    lines.append("> 本文件由 `inspect_template.py` 自动生成，供人工核对与撰写 requirements.json 参考。")
    lines.append("")
    lines.append("## 1. 页面设置（每个分节）")
    for s in spec["sections"]:
        lines.append(f"- 分节 {s['index']}:")
        if "pgSz" in s:
            p = s["pgSz"]
            lines.append(f"  - 纸张: {p['w']}×{p['h']} twips ({int(p['w'])/1440:.2f}×{int(p['h'])/1440:.2f} in), 方向={p['orient']}")
        if "pgMar" in s:
            m = s["pgMar"]
            lines.append("  - 页边距(twips): " + ", ".join(f"{k}={v}" for k, v in m.items() if v is not None))
        if "pgNumType" in s:
            lines.append("  - 页码: " + ", ".join(f"{k}={v}" for k, v in s["pgNumType"].items()))
        lines.append(f"  - 首页不同: {s.get('titlePg')}")
        for kind in ("header", "footer"):
            for ref in s.get(kind + "Refs", []):
                paras = ref.get("paragraphs", [])
                descr = " | ".join(
                    (p["text"] or "(空)") + (" [域:" + ",".join(p["fields"]) + "]" if p["fields"] else "")
                    for p in paras
                )
                lines.append(f"  - {kind}[{ref['type']}]: {descr or '(无内容)'}")
    lines.append("")
    lines.append("## 2. 样式（字体 / 字号 / 对齐 / 行距 / 缩进）")
    lines.append("")
    lines.append("| 样式ID | 名称 | 中文字体 | 西文字体 | 字号 | 加粗 | 对齐 | 行距 | 首行缩进 |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- | --- | --- |")
    for sid, info in spec["styles"].items():
        if sid.startswith("__"):
            continue
        r, p = info["run"], info["paragraph"]
        sp = p.get("spacing", {}) if isinstance(p.get("spacing"), dict) else {}
        row = [
            sid, info.get("name") or "", r.get("font_eastAsia", ""), r.get("font_ascii", ""),
            size_name(r.get("size_pt")), "是" if r.get("bold") else "", p.get("alignment", ""),
            (sp.get("line") or "") + ("(" + sp["lineRule"] + ")" if sp.get("lineRule") else ""),
            p.get("firstLine") or p.get("firstLineChars") or "",
        ]
        lines.append("| " + " | ".join(str(x) for x in row) + " |")
    lines.append("")
    lines.append("## 3. 样式角色推测（供内容映射参考）")
    for role, sids in guess_roles(spec["styles"]).items():
        lines.append(f"- {role}: {', '.join(sids) if sids else '(未找到，需人工指定)'}")
    lines.append("")
    lines.append("## 4. 段落示例（前 20 段）")
    for p in spec["paragraphs"][:20]:
        lines.append(f"- [{p['index']}] style={p['style'] or '(默认)'} :: {p['text'][:80]}")
    lines.append("")
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser(description="Inspect a .docx template and produce a format spec.")
    ap.add_argument("template")
    ap.add_argument("-o", "--output", help="spec JSON output path (default: stdout)")
    ap.add_argument("--summary", help="optional human-readable markdown summary path")
    args = ap.parse_args()

    parts = read_docx(args.template)
    spec = build_spec(parts, os.path.basename(args.template))
    payload = json.dumps(spec, ensure_ascii=False, indent=2)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(payload)
        print(f"SPEC_WRITTEN={args.output}")
    else:
        print(payload)
    if args.summary:
        with open(args.summary, "w", encoding="utf-8") as f:
            f.write(human_summary(spec))
        print(f"SUMMARY_WRITTEN={args.summary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
