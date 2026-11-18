#!/usr/bin/env python3
"""One-command end-to-end self-test for thesis-formatter.

Generates a demo school template, runs the full pipeline
(inspect -> build -> verify), asserts all checks pass, then runs a negative
test to confirm the verifier actually catches a broken format.

Usage:
    python tests/selftest.py

No third-party dependencies.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
SCRIPTS = SKILL / "scripts"
ASSETS = SKILL / "assets"
sys.path.insert(0, str(SCRIPTS))


def make_demo_template(path: Path) -> None:
    """Create a minimal but realistic school template .docx with a PAGE field footer."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ct = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    docrels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''
    styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr><w:pPr><w:jc w:val="both"/><w:ind w:firstLine="420"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:outlineLvl w:val="0"/><w:keepNext/></w:pPr><w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="21"/></w:rPr><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
</w:styles>'''
    footer = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:t>第 </w:t></w:r>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:r><w:t> 页</w:t></w:r></w:p></w:ftr>'''
    doc = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>
<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>模板占位正文，将被替换。</w:t></w:r></w:p>
<w:sectPr>
<w:footerReference w:type="default" r:id="rIdF"/>
<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
<w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:num="1" w:space="720"/><w:pgNumType w:fmt="decimal" w:start="1"/>
</w:sectPr></w:body></w:document>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", docrels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/footer1.xml", footer)


def run(script: str, *args: str) -> subprocess.CompletedProcess:
    cmd = [sys.executable, str(SCRIPTS / script), *[str(a) for a in args]]
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")


def fail_count(report_path: Path) -> int:
    text = report_path.read_text(encoding="utf-8")
    for line in text.splitlines():
        if line.startswith("PASS="):
            return int(line.split("FAIL=")[1].split()[0])
    return -1


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="thesis-formatter-selftest-") as td:
        out = Path(td)
        tpl = out / "demo_template.docx"
        spec = out / "spec.json"
        summary = out / "summary.md"
        built = out / "output.docx"
        report = out / "report.md"

        print("=" * 60)
        print("论文排版 Skill 一键自测")
        print("=" * 60)
        make_demo_template(tpl)
        print("[1/5] 生成演示模板 ............................ OK")

        r = run("inspect_template.py", tpl, "-o", spec, "--summary", summary)
        if r.returncode != 0:
            print("解析模板失败:\n", r.stdout, r.stderr)
            return 1
        print("[2/5] 解析模板 -> spec.json .................. OK")

        r = run("build_document.py", tpl, ASSETS / "content.example.json", built)
        if r.returncode != 0:
            print("生成文档失败:\n", r.stdout, r.stderr)
            return 1
        print("[3/5] 克隆模板 + 替换正文 -> output.docx ..... OK")

        r = run("verify_requirements.py", built, "--spec", spec,
                "--requirements", ASSETS / "requirements.example.json", "--report", report)
        nfail = fail_count(report)
        if r.returncode != 0 or nfail != 0:
            print(f"正向验收失败 (FAIL={nfail}):\n", r.stdout, r.stderr)
            return 1
        header = report.read_text(encoding="utf-8").splitlines()[2]
        print(f"[4/5] 逐项验收 ............................... {header}")

        # Negative test: deliberately break a hard requirement (wrong font size),
        # verify must report FAIL and exit non-zero.
        bad_req = json.loads((ASSETS / "requirements.example.json").read_text(encoding="utf-8"))
        bad_req["styles"]["Normal"]["size_pt"] = 99
        bad_path = out / "bad_requirements.json"
        bad_path.write_text(json.dumps(bad_req, ensure_ascii=False), encoding="utf-8")
        bad_report = out / "bad_report.md"
        r = run("verify_requirements.py", built, "--spec", spec,
                "--requirements", bad_path, "--report", bad_report)
        if r.returncode == 0 or fail_count(bad_report) < 1:
            print("负向测试失败：验收器未能拦截错误的字号要求！")
            return 1
        print(f"[5/5] 负向测试（故意写错字号）.............. 正确拦截 (FAIL={fail_count(bad_report)})")

        print("=" * 60)
        print("全部通过：skill 可正常解析模板、克隆排版并逐项验收。")
        print("验收报告样例：")
        print("=" * 60)
        lines = report.read_text(encoding="utf-8").splitlines()
        print("\n".join(lines[:9]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
