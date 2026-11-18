#!/usr/bin/env python3
"""Build a formatted thesis .docx by cloning a school template package.

Strategy (format-safe): the template .docx package is copied verbatim — styles,
numbering, headers, footers, page setup, theme, fonts — and only the body
content is replaced.  This guarantees that font sizes, page numbers, header
and footer definitions match the template exactly.

Usage:
    python build_document.py <template.docx> <content.json> <output.docx> \
        [--stylesheet-map styles.json]

content.json format (UTF-8):
{
  "title": "论文题目",
  "paragraphs": [
    {"style": "Heading1", "text": "第1章 绪论"},
    {"style": "Normal", "text": "正文段落……", "indent_first_line": true},
    {"style": "Normal", "text": "换页后的段落", "page_break_before": true},
    {"style": "Caption", "text": "图1-1 示例"},
    {"page_break": true},
    {"style": "Normal", "text": "参考文献条目", "numbering": {"numId": "1", "ilvl": "0"}}
  ]
}

Any property not specified is inherited from the template style, so the
template remains the single source of truth for all format requirements.
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_utils import (NS, PKG_REL, child, children, parse_xml, qn,  # noqa: E402
                        read_docx, serialize_xml, write_docx)

W = NS["w"]


def make_element(tag: str, **attrs) -> ET.Element:
    el = ET.Element(qn(tag))
    for k, v in attrs.items():
        el.set(qn("w:" + k), str(v))
    return el


def add_run(p: ET.Element, text: str, rpr: dict | None = None) -> ET.Element:
    run = ET.SubElement(p, qn("w:r"))
    if rpr:
        rpr_el = ET.SubElement(run, qn("w:rPr"))
        if rpr.get("bold"):
            ET.SubElement(rpr_el, qn("w:b"))
        if rpr.get("italic"):
            ET.SubElement(rpr_el, qn("w:i"))
        fonts = rpr.get("fonts")
        if fonts:
            f = ET.SubElement(rpr_el, qn("w:rFonts"))
            for k, v in fonts.items():
                f.set(qn("w:" + k), v)
        if rpr.get("size_pt"):
            ET.SubElement(rpr_el, qn("w:sz"), {qn("w:val"): str(int(float(rpr["size_pt"]) * 2))})
    for i, seg in enumerate(text.split("\n")):
        if i:
            ET.SubElement(run, qn("w:br"))
        t = ET.SubElement(run, qn("w:t"))
        if seg.startswith(" ") or seg.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = seg
    return run


def build_paragraph(item: dict, style_map: dict) -> ET.Element:
    p = ET.Element(qn("w:p"))
    style = item.get("style")
    if style:
        style = style_map.get(style, style)
    ppr = ET.SubElement(p, qn("w:pPr"))
    if style:
        ET.SubElement(ppr, qn("w:pStyle"), {qn("w:val"): str(style)})

    if item.get("page_break_before"):
        ET.SubElement(ppr, qn("w:pageBreakBefore"))
    if item.get("keep_next"):
        ET.SubElement(ppr, qn("w:keepNext"))
    if item.get("keep_lines"):
        ET.SubElement(ppr, qn("w:keepLines"))
    if item.get("alignment"):
        ET.SubElement(ppr, qn("w:jc"), {qn("w:val"): str(item["alignment"])})
    if item.get("indent_first_line"):
        ind = child(ppr, "w:ind") or ET.SubElement(ppr, qn("w:ind"))
        ind.set(qn("w:firstLine"), str(item.get("first_line_twips", 420)))
    if item.get("indent"):
        ind = child(ppr, "w:ind") or ET.SubElement(ppr, qn("w:ind"))
        for k, v in item["indent"].items():
            ind.set(qn("w:" + k), str(v))
    if item.get("spacing"):
        sp = child(ppr, "w:spacing") or ET.SubElement(ppr, qn("w:spacing"))
        for k, v in item["spacing"].items():
            sp.set(qn("w:" + k), str(v))

    num = item.get("numbering")
    if num:
        numpr = ET.SubElement(ppr, qn("w:numPr"))
        ET.SubElement(numpr, qn("w:ilvl"), {qn("w:val"): str(num.get("ilvl", 0))})
        ET.SubElement(numpr, qn("w:numId"), {qn("w:val"): str(num.get("numId", 1))})

    text = item.get("text", "")
    add_run(p, text, item.get("run"))
    return p


def replace_body(template_parts: dict, content: dict, style_map: dict) -> dict:
    doc = parse_xml(template_parts["word/document.xml"])
    body = child(doc, "w:body")
    if body is None:
        raise ValueError("template word/document.xml has no <w:body>")

    sect_prs = [el for el in body if el.tag == qn("w:sectPr")]

    for el in list(body):
        if el.tag != qn("w:sectPr"):
            body.remove(el)

    for item in content.get("paragraphs", []):
        if item.get("page_break") and not item.get("text"):
            p = ET.Element(qn("w:p"))
            run = ET.SubElement(p, qn("w:r"))
            ET.SubElement(run, qn("w:br"), {qn("w:type"): "page"})
            body.append(p)
            continue
        body.append(build_paragraph(item, style_map))

    for sect in sect_prs:
        body.append(sect)

    out = dict(template_parts)
    out["word/document.xml"] = serialize_xml(doc)
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Clone a template and replace its body content.")
    ap.add_argument("template")
    ap.add_argument("content")
    ap.add_argument("output")
    ap.add_argument("--stylesheet-map", help="optional JSON mapping role names to real style IDs")
    args = ap.parse_args()

    with open(args.content, "r", encoding="utf-8") as f:
        content = json.load(f)
    style_map = {}
    if args.stylesheet_map:
        with open(args.stylesheet_map, "r", encoding="utf-8") as f:
            style_map = json.load(f)

    parts = read_docx(args.template)
    result = replace_body(parts, content, style_map)
    write_docx(args.output, result)
    print(f"BUILT={args.output}")
    print(f"PARAGRAPHS={len(content.get('paragraphs', []))}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
