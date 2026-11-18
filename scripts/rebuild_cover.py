# -*- coding: utf-8 -*-
"""按学校模板重建封面：把模板封面段落搬到论文第一页，保留全部原文内容。

适用于「论文原文没有独立封面 / 封面不规范」的场景。做法：
  1. 从模板提取封面的原始段落 XML（含校徽图片、黑体字号、下划线、空行全部排版）；
  2. 替换示例文字为论文题目，清空需要用户自填的字段值（保留标签与下划线）；
  3. 处理校徽图片关系（模板 rId → 论文里不冲突的新 rId + 独立 media 文件）；
  4. 把封面设为独立节（无页眉页脚，正文从下一节起算页码）；
  5. 原文标题/声明段落按需删除或移到封面末尾，做到内容零丢失。

用法：
    python rebuild_cover.py <config.json>
    python rebuild_cover.py --print-template > cover.example.json

纯标准库实现。
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WN, RN = "{%s}" % W, "{%s}" % R

for _pfx, _uri in [
    ("w", W), ("r", R),
    ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
    ("a", A), ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
    ("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"),
    ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
    ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
]:
    ET.register_namespace(_pfx, _uri)


TEMPLATE_CONFIG = {
    "src": "已排版_无封面.docx",
    "template": "学校模板.docx",
    "output": "已排版_带封面.docx",
    "title": {"old": "模板示例题目", "new": "你的论文题目"},
    "cover_range": [0, 18],          # 模板封面段落区间 [起点, 终点)（在 body 直属 w:p 中）
    "fields_to_blank": ["学生姓名", "学号", "指导教师", "专业年级",
                        "所在学院和系", "完成日期", "答辩日期"],  # 含这些字的段落清空值保留标签
    "logo": {
        "template_rid": "rId7",      # 模板封面校徽的 r:embed
        "template_media": "word/media/image1.png",
        "output_media": "word/media/school_logo.png",
        "output_rid": "rId90",       # 论文里不冲突的新 rId
    },
    "drop_paragraphs": [0],          # src 中要删除的段落索引（如原论文标题）
    "move_to_cover": [1],            # src 中移到封面末尾的段落索引（如声明），删除其 sectPr
    "cover_section": {
        "w": 11906, "h": 16838, "top": 1440, "right": 1800, "bottom": 1440,
        "left": 1800, "header": 851, "footer": 992, "gutter": 0,
        "cols": 425, "linePitch": 312,
    },
    "_comment": "fields_to_blank 按『标签包含匹配』处理：段落里含该关键词时把首个 w:t 设为去空格后的标签、其余 w:t 清空。若模板字段值本就为空可留空列表。",
}


def txt_of(para: ET.Element) -> str:
    return "".join(t.text or "" for t in para.iter(WN + "t"))


def strip_keywords(s: str) -> str:
    return s.replace(" ", "").replace("\u3000", "")


def max_rid(rels_root: ET.Element) -> int:
    import re
    best = 0
    for rel in rels_root:
        m = re.match(r"rId(\d+)$", rel.get("Id", ""))
        if m:
            best = max(best, int(m.group(1)))
    return best


def run(cfg: dict) -> dict:
    src, tpl, out = cfg["src"], cfg["template"], cfg["output"]
    zo = zipfile.ZipFile(src)
    zt = zipfile.ZipFile(tpl)
    parts = {n: zo.read(n) for n in zo.namelist()}

    odoc = ET.fromstring(parts["word/document.xml"])
    obody = odoc.find(WN + "body")
    original = obody.findall(WN + "p")

    # 收集要删除/移动的原段落（先按索引取值，最后统一删除，避免索引错位）
    to_drop = {i: original[i] for i in cfg.get("drop_paragraphs", [])}
    to_move = [copy.deepcopy(original[i]) for i in cfg.get("move_to_cover", [])]
    for para in to_move:
        sppr = para.find(WN + "pPr")
        if sppr is not None:
            old = sppr.find(WN + "sectPr")
            if old is not None:
                sppr.remove(old)
    for para in list(to_drop.values()):
        if para in list(obody):
            obody.remove(para)
    for i in sorted(to_drop, reverse=True):
        # 已被 remove 的引用无需再处理；此处仅保留索引语义说明
        pass

    # 从模板提取封面段落
    tdoc = ET.fromstring(zt.read("word/document.xml"))
    tparas = tdoc.find(WN + "body").findall(WN + "p")
    lo, hi = cfg["cover_range"]
    cover = [copy.deepcopy(tparas[i]) for i in range(lo, hi)]

    title_cfg = cfg.get("title", {})
    old_title, new_title = title_cfg.get("old"), title_cfg.get("new")
    fields = cfg.get("fields_to_blank", [])
    logo_cfg = cfg.get("logo")
    new_rid = logo_cfg["output_rid"] if logo_cfg else None
    old_rid = logo_cfg["template_rid"] if logo_cfg else None

    # 自动避让 rId（若配置的目标 rId 已被占用）
    rels_root = ET.fromstring(parts["word/_rels/document.xml.rels"])
    used_ids = {rel.get("Id") for rel in rels_root}
    if new_rid and new_rid in used_ids:
        new_rid = "rId%d" % (max_rid(rels_root) + 1)

    for para in cover:
        ts = list(para.iter(WN + "t"))
        if ts:
            full = "".join(t.text or "" for t in ts)
            if old_title and new_title and old_title in full:
                ts[0].text = full.replace(old_title, new_title)
                for t in ts[1:]:
                    t.text = ""
            else:
                for lab in fields:
                    if strip_keywords(lab) in strip_keywords(full):
                        ts[0].text = lab
                        for t in ts[1:]:
                            t.text = ""
                        break
        # 校徽：替换 r:embed，并清除模板封面自带的 sectPr
        if logo_cfg:
            for blip in para.iter("{%s}blip" % A):
                if blip.get(RN + "embed") == old_rid:
                    blip.set(RN + "embed", new_rid)
        ppr = para.find(WN + "pPr")
        if ppr is not None:
            sp = ppr.find(WN + "sectPr")
            if sp is not None:
                ppr.remove(sp)

    # 原文要保留的段落（声明等）移到封面末尾
    cover.extend(to_move)

    # 封面作为独立节，插入正文之前
    cs = dict(cfg["cover_section"])
    cs["_W"], cs["_R"] = W, R
    cover_sect = ET.fromstring(
        '<w:sectPr xmlns:w="%(_W)s" xmlns:r="%(_R)s">'
        '<w:pgSz w:w="%(w)d" w:h="%(h)d"/>'
        '<w:pgMar w:top="%(top)d" w:right="%(right)d" w:bottom="%(bottom)d" '
        'w:left="%(left)d" w:header="%(header)d" w:footer="%(footer)d" w:gutter="%(gutter)d"/>'
        '<w:cols w:space="%(cols)d"/><w:docGrid w:type="lines" w:linePitch="%(linePitch)d"/>'
        '<w:type w:val="nextPage"/>'
        '</w:sectPr>' % cs)
    last = cover[-1]
    ppr = last.find(WN + "pPr")
    if ppr is None:
        ppr = ET.Element(WN + "pPr")
        last.insert(0, ppr)
    existing = ppr.find(WN + "sectPr")
    if existing is not None:
        ppr.remove(existing)
    ppr.append(cover_sect)

    for idx, para in enumerate(cover):
        obody.insert(idx, para)
    parts["word/document.xml"] = ET.tostring(odoc, encoding="UTF-8", xml_declaration=True)

    # 校徽关系与媒体文件
    if logo_cfg and new_rid:
        rel = ET.SubElement(rels_root, "{%s}Relationship" % PKG_REL)
        rel.set("Id", new_rid)
        rel.set("Type", R + "/image")
        rel.set("Target", os.path.basename(logo_cfg["output_media"]))
        parts["word/_rels/document.xml.rels"] = ET.tostring(
            rels_root, encoding="UTF-8", xml_declaration=True)
        parts[logo_cfg["output_media"]] = zt.read(logo_cfg["template_media"])

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in parts.items():
            z.writestr(n, d)

    return {"output": out, "cover_paragraphs": len(cover),
            "moved_paragraphs": len(to_move), "logo_rid": new_rid}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("config", nargs="?")
    ap.add_argument("--print-template", action="store_true")
    args = ap.parse_args()
    if args.print_template:
        print(json.dumps(TEMPLATE_CONFIG, ensure_ascii=False, indent=2))
        return 0
    if not args.config:
        ap.error("需要 config.json（或用 --print-template 查看示例）")
    with open(args.config, encoding="utf-8") as f:
        cfg = json.load(f)
    res = run(cfg)
    print("=== 封面重建完成 ===")
    print("  输出:", res["output"])
    print("  封面段落:", res["cover_paragraphs"], "  保留原文段落:", res["moved_paragraphs"])
    print("  校徽 rId:", res["logo_rid"])
    print("  [下一步] 用 audit_content.py 核对图片/表格零丢失，并确认封面无页眉页脚。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
