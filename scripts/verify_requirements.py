#!/usr/bin/env python3
"""Verify a formatted thesis .docx against a format spec / requirements file.

Usage:
    python verify_requirements.py <output.docx> --spec <spec.json> \
        [--requirements requirements.json] [--report report.md] [--strict]

Exit code 0 = all checks PASS, 1 = at least one FAIL (or WARN with --strict).

requirements.json (optional, all keys optional — adds/overrides checks):
{
  "styles": {
    "Normal": {"font_eastAsia": "宋体", "size_pt": 12, "alignment": "both"},
    "Heading1": {"size_pt": 16, "bold": true}
  },
  "sections": [{"pgMar": {"top": "1440", "left": "1800"}, "pgNumType": {"fmt": "decimal", "start": "1"}}],
  "required_text": ["摘要", "Abstract", "参考文献"],
  "forbidden_text": ["【填写题目】", "XXX", "待补充"],
  "require_page_field_in_footer": true,
  "require_page_field_in_header": false,
  "min_paragraphs": 30
}
"""

from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_utils import build_spec, read_docx, size_name  # noqa: E402


class Checks:
    def __init__(self):
        self.rows = []

    def add(self, ok, category, name, detail=""):
        self.rows.append({"status": "PASS" if ok else "FAIL", "category": category,
                          "name": name, "detail": detail})

    def warn(self, category, name, detail=""):
        self.rows.append({"status": "WARN", "category": category, "name": name, "detail": detail})

    @property
    def failed(self):
        return any(r["status"] == "FAIL" for r in self.rows)

    @property
    def warned(self):
        return any(r["status"] == "WARN" for r in self.rows)

    def markdown(self) -> str:
        lines = ["# 排版格式验收报告", ""]
        counts = {}
        for r in self.rows:
            counts[r["status"]] = counts.get(r["status"], 0) + 1
        lines.append(f"PASS={counts.get('PASS', 0)}  FAIL={counts.get('FAIL', 0)}  WARN={counts.get('WARN', 0)}")
        lines.append("")
        lines.append("| 结果 | 类别 | 检查项 | 说明 |")
        lines.append("| --- | --- | --- | --- |")
        icon = {"PASS": "通过", "FAIL": "未通过", "WARN": "注意"}
        for r in self.rows:
            lines.append(f"| {icon[r['status']]} | {r['category']} | {r['name']} | {r['detail']} |")
        lines.append("")
        lines.append("## 结论")
        lines.append("")
        if self.failed:
            lines.append("**未通过**。请按上表 FAIL 项逐条修正后重新运行验证。")
        elif self.warned:
            lines.append("**通过（含提示）**。所有硬性要求已满足，请人工复核 WARN 项。")
        else:
            lines.append("**全部通过**。字号、页码、页眉页脚等要求与模板一致。")
        return "\n".join(lines)


def deep_get(d, path, default=None):
    cur = d
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return default
        cur = cur[part]
    return cur


def compare_value(actual, expected) -> bool:
    if actual is None:
        return False
    try:
        return abs(float(actual) - float(expected)) < 0.01
    except (TypeError, ValueError):
        return str(actual).strip() == str(expected).strip()


def verify(output_spec: dict, template_spec: dict | None, req: dict | None) -> Checks:
    c = Checks()

    # ---- Page setup (derived from template spec) ----
    if template_spec:
        tsec = template_spec.get("sections", [{}])[0]
        osec = (output_spec.get("sections") or [{}])[0]
        for key in ("pgSz", "pgMar", "pgNumType", "cols"):
            tv, ov = tsec.get(key, {}), osec.get(key, {})
            if not isinstance(tv, dict):
                continue
            for sub, exp in tv.items():
                if exp is None:
                    continue
                got = ov.get(sub)
                ok = compare_value(got, exp)
                c.add(ok, "页面设置", f"{key}.{sub}", f"模板={exp} 实际={got}")
        if tsec.get("titlePg") is not None:
            c.add(bool(osec.get("titlePg")) == bool(tsec.get("titlePg")),
                  "页面设置", "首页不同(titlePg)",
                  f"模板={tsec.get('titlePg')} 实际={osec.get('titlePg')}")

    # ---- Header / footer fields ----
    sections = output_spec.get("sections", [])
    footer_fields, header_fields = set(), set()
    for s in sections:
        for kind, acc in (("footer", footer_fields), ("header", header_fields)):
            for ref in s.get(kind + "Refs", []):
                for f in ref.get("fields", []):
                    acc.add(f)
    if req and req.get("require_page_field_in_footer"):
        c.add("PAGE" in footer_fields, "页码", "页脚包含 PAGE 域",
              "发现域: " + (",".join(sorted(footer_fields)) or "无"))
    if req and req.get("require_page_field_in_header"):
        c.add("PAGE" in header_fields, "页码", "页眉包含 PAGE 域",
              "发现域: " + (",".join(sorted(header_fields)) or "无"))
    if template_spec:
        tsec = (template_spec.get("sections") or [{}])[0]
        osec = (sections or [{}])[0]
        for kind in ("header", "footer"):
            trefs = tsec.get(kind + "Refs", [])
            orefs = osec.get(kind + "Refs", [])
            for i, tr in enumerate(trefs):
                if i >= len(orefs):
                    c.add(False, "页眉页脚", f"{kind}[{tr.get('type')}]存在性", "输出缺失该引用")
                    continue
                tf = set(tr.get("fields", []))
                of = set(orefs[i].get("fields", []))
                c.add(tf.issubset(of), "页眉页脚", f"{kind}[{tr.get('type')}]域保持",
                      f"模板={sorted(tf)} 实际={sorted(of)}")
                tt = [p["text"] for p in tr.get("paragraphs", []) if p.get("text")]
                ot = [p["text"] for p in orefs[i].get("paragraphs", []) if p.get("text")]
                if tt:
                    c.add(tt == ot, "页眉页脚", f"{kind}[{tr.get('type')}]文字",
                          f"模板={tt} 实际={ot}")

    # ---- Explicit requirements ----
    req = req or {}
    out_styles = output_spec.get("styles", {})
    for sid, expected in (req.get("styles") or {}).items():
        info = out_styles.get(sid)
        if info is None:
            c.add(False, "样式", f"样式 {sid} 存在", "输出文档无此样式")
            continue
        for key, exp in expected.items():
            if key in ("font_eastAsia", "font_ascii", "font_hAnsi", "color"):
                got = info.get("run", {}).get(key)
            elif key in ("size_pt", "bold", "italic"):
                got = info.get("run", {}).get(key)
                if key == "size_pt" and got is not None:
                    c.add(compare_value(got, exp), "样式", f"{sid}.字号",
                          f"要求={size_name(exp)} 实际={size_name(got)}")
                    continue
            elif key in ("alignment", "firstLine", "firstLineChars", "line", "lineRule",
                         "before", "after", "left", "right"):
                got = info.get("paragraph", {}).get(key,
                      (info.get("paragraph", {}).get("spacing") or {}).get(key) if
                      key in ("line", "lineRule", "before", "after") else None)
            else:
                got = None
            c.add(compare_value(got, exp), "样式", f"{sid}.{key}", f"要求={exp} 实际={got}")

    for i, expected in enumerate(req.get("sections", [])):
        s = output_spec.get("sections", [{}])[i] if i < len(output_spec.get("sections", [])) else {}
        for key, exp in expected.items():
            got = s.get(key)
            if isinstance(exp, dict):
                for sub, subexp in exp.items():
                    c.add(compare_value((got or {}).get(sub), subexp), "页面设置",
                          f"section[{i}].{key}.{sub}", f"要求={subexp} 实际={(got or {}).get(sub)}")
            else:
                c.add(compare_value(got, exp), "页面设置", f"section[{i}].{key}",
                      f"要求={exp} 实际={got}")

    all_text = "\n".join(p.get("text", "") for p in output_spec.get("paragraphs", []))
    for t in req.get("required_text", []):
        c.add(t in all_text, "内容", f"包含“{t}”", "")
    for t in req.get("forbidden_text", []):
        c.add(t not in all_text, "内容", f"不含“{t}”", "")
    if req.get("min_paragraphs"):
        n = len(output_spec.get("paragraphs", []))
        c.add(n >= req["min_paragraphs"], "内容", "段落数量",
              f"要求≥{req['min_paragraphs']} 实际={n}")

    used = {p.get("style") for p in output_spec.get("paragraphs", []) if p.get("style")}
    for sid in used:
        if sid not in out_styles:
            c.warn("样式", f"段落引用未定义样式 {sid}", "样式未在 styles.xml 中找到")
    if not footer_fields and not header_fields:
        c.warn("页码", "未发现任何页码域", "如模板要求页码，请检查页眉页脚")
    return c


def main() -> int:
    ap = argparse.ArgumentParser(description="Verify formatted docx against template spec.")
    ap.add_argument("output")
    ap.add_argument("--spec", help="JSON spec generated by inspect_template.py")
    ap.add_argument("--requirements", help="optional explicit requirements.json")
    ap.add_argument("--report", help="write markdown report to this path")
    ap.add_argument("--strict", action="store_true", help="treat WARN as failure")
    args = ap.parse_args()

    template_spec = None
    if args.spec:
        with open(args.spec, "r", encoding="utf-8") as f:
            template_spec = json.load(f)
    req = None
    if args.requirements:
        with open(args.requirements, "r", encoding="utf-8") as f:
            req = json.load(f)

    out_spec = build_spec(read_docx(args.output), os.path.basename(args.output))
    checks = verify(out_spec, template_spec, req)
    report = checks.markdown()
    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            f.write(report)
        print(f"REPORT_WRITTEN={args.report}")
    print(report)
    if checks.failed or (args.strict and checks.warned):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
