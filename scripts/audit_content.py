# -*- coding: utf-8 -*-
"""内容零丢失核验：对比排版前后两个 .docx 的文字/图片/表格/绘图对象是否一致。

格式移植（transform_paper.py）会重写 document.xml，必须用本工具证明
「只改格式、没动内容」。纯标准库实现。

用法：
    python audit_content.py <原论文.docx> <排版后.docx> [--json 报告.json]

退出码：0 = 内容完全一致；1 = 检测到差异（不可交付）。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def signature(path: str) -> dict:
    z = zipfile.ZipFile(path)
    doc = ET.fromstring(z.read("word/document.xml"))
    # 文字：按文档顺序拼接所有 w:t，做 SHA256（含表格内文字）
    texts = [t.text or "" for t in doc.iter(W + "t")]
    joined = "".join(texts)
    # 图片：按名字排序的 md5 列表（比较字节，不看文件名）
    media = sorted(
        hashlib.md5(z.read(n)).hexdigest()
        for n in z.namelist()
        if n.startswith("word/media/")
    )
    return {
        "text_sha256": hashlib.sha256(joined.encode("utf-8")).hexdigest(),
        "text_chars": len(joined),
        "text_runs": len(texts),
        "tables": len(doc.findall(".//" + W + "tbl")),
        "drawings": len(doc.findall(".//" + W + "drawing")),
        "media_count": len(media),
        "media_hashes": media,
        "zip_ok": z.testzip() is None,
    }


def all_xml_wellformed(path: str):
    z = zipfile.ZipFile(path)
    bad = []
    for n in z.namelist():
        if n.endswith((".xml", ".rels")):
            try:
                ET.fromstring(z.read(n))
            except ET.ParseError as e:
                bad.append({"part": n, "error": str(e)})
    return bad


def normalized_text(path: str) -> str:
    """把章标题的中文数字/阿拉伯数字编号统一成占位符，用于容忍编号规范化。"""
    import re
    z = zipfile.ZipFile(path)
    doc = ET.fromstring(z.read("word/document.xml"))
    joined = "".join(t.text or "" for t in doc.iter(W + "t"))
    joined = re.sub(r"第[一二三四五六七八九十百]+章", "第N章", joined)
    joined = re.sub(r"第\d+章", "第N章", joined)
    return joined


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("original")
    ap.add_argument("formatted")
    ap.add_argument("--json", dest="json_out")
    ap.add_argument("--normalize-headings", action="store_true",
                    help="忽略章标题编号的中/阿拉伯数字差异（编号规范化属预期变更）")
    args = ap.parse_args()

    a = signature(args.original)
    b = signature(args.formatted)
    bad_xml = all_xml_wellformed(args.formatted)

    heading_norm_ok = None
    if args.normalize_headings and a["text_sha256"] != b["text_sha256"]:
        heading_norm_ok = normalized_text(args.original) == normalized_text(args.formatted)

    # 注意：章标题中文数字→阿拉伯数字属于「预期的文字变更」，
    # 由 transform_paper.py 在 rename 时用 --allow-text-delta 记录原始哈希。
    # 本工具默认严格比对；若需允许标题重编号，传入基线文件作为 original。
    text_ok = a["text_sha256"] == b["text_sha256"]
    if not text_ok and heading_norm_ok is True:
        text_ok = True  # 差异仅来自标题编号规范化
    checks = {
        "text_sha256": text_ok,
        "tables": a["tables"] == b["tables"],
        "drawings": a["drawings"] == b["drawings"],
        "media_count": a["media_count"] == b["media_count"],
        "media_bytes": a["media_hashes"] == b["media_hashes"],
        "zip_ok": b["zip_ok"],
        "xml_wellformed": not bad_xml,
    }
    passed = all(checks.values())

    report = {
        "original": args.original,
        "formatted": args.formatted,
        "original_signature": a,
        "formatted_signature": b,
        "checks": checks,
        "xml_errors": bad_xml,
        "PASS": passed,
        "note": ("文字哈希不同通常是章标题编号被规范化（中文数字→阿拉伯数字）。"
                 "若确认只改了标题编号，此差异属预期；图片/表格/绘图一致即可交付。"),
    }
    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

    print("=== 内容零丢失核验 ===")
    for k, v in checks.items():
        label = k
        if k == "text_sha256" and a["text_sha256"] != b["text_sha256"] and v:
            label = "text_sha256 (归一化章号后一致)"
        print(f"  {'PASS' if v else 'FAIL'}  {label}")
    print(f"  文字: 原 {a['text_chars']} 字 / 新 {b['text_chars']} 字")
    print(f"  图片: 原 {a['media_count']} 张 / 新 {b['media_count']} 张")
    print(f"  表格: 原 {a['tables']} / 新 {b['tables']}   绘图: 原 {a['drawings']} / 新 {b['drawings']}")
    if not checks.get("text_sha256"):
        print("  [提示] " + report["note"])
    print("结果:", "PASS 内容一致" if passed else "FAIL 检测到差异，请复核")
    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
