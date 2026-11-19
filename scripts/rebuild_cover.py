# -*- coding: utf-8 -*-
"""按学校模板重建封面（默认自动化，零配置）。

设计目标：换一个 AI / 换一份模板也能稳定生成正确封面，不依赖人工填配置。
默认行为（零配置命令）：
    python rebuild_cover.py <已排版无封面.docx> <学校模板.docx> <输出.docx>

自动完成：
  1. 自动识别模板封面范围：从 body 第一个 w:p 到第一个带 sectPr 的段落（含）；
  2. 自动取论文题目：src 第一个非空段落；
  3. 只替换题目：封面「题目」栏的示例值换成论文题目；
  4. 个人信息一律清空（含姓名/学号/导师/专业/学院/日期），只保留标签与下划线，
     —— 不强制用户提供任何个人信息（属隐私，由使用者自行填写）；
  5. 自动搬迁校徽：从模板 rels 找到封面图片，分配不冲突的新 rId 与独立 media 文件；
  6. 封面设为独立节（无页眉页脚），页面设置自动从模板封面节读取。

可选覆盖：若自动识别不适用，可传 JSON 覆盖任意字段（见 --print-template）。
    python rebuild_cover.py <config.json>

纯标准库实现。
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import re
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

# 个人信息标签：命中即清空其值（保留标签与下划线）。题目不在此列。
PERSONAL_LABELS = [
    "学生姓名", "学号", "指导教师", "专业年级", "所在学院和系", "完成日期", "答辩日期",
    "学生姓名", "姓名", "专业", "学院", "系别", "班级", "完成时间", "答辩时间",
]
TITLE_LABEL_RE = re.compile(r"^\s*题\s*目\s*[:：]")


def txt_of(para):
    return "".join(t.text or "" for t in para.iter(WN + "t"))


def norm(s):
    return s.replace(" ", "").replace("\u3000", "")


def max_rid(rels_root):
    best = 0
    for rel in rels_root:
        m = re.match(r"rId(\d+)$", rel.get("Id", ""))
        if m:
            best = max(best, int(m.group(1)))
    return best


def para_has_sect(para):
    ppr = para.find(WN + "pPr")
    return ppr is not None and ppr.find(WN + "sectPr") is not None


def detect_cover_range(tparas):
    """封面 = body 直属段落 [0, 第一个带 sectPr 的段落索引+1)。"""
    for i, p in enumerate(tparas):
        if para_has_sect(p):
            return 0, i + 1
    last = 0
    for i, p in enumerate(tparas[:40]):
        if any(k in norm(txt_of(p)) for k in ("答辩日期", "完成日期")):
            last = i
    return 0, max(last + 1, 1)


def detect_cover_section(tparas):
    """从模板封面节的 sectPr 读取页面设置。"""
    default = dict(w=11906, h=16838, top=1440, right=1800, bottom=1440,
                   left=1800, header=851, footer=992, gutter=0,
                   cols=425, linePitch=312)
    for p in tparas[:40]:
        ppr = p.find(WN + "pPr")
        if ppr is None:
            continue
        sect = ppr.find(WN + "sectPr")
        if sect is None:
            continue
        sz = sect.find(WN + "pgSz")
        mar = sect.find(WN + "pgMar")
        cols = sect.find(WN + "cols")
        grid = sect.find(WN + "docGrid")
        if sz is not None:
            default["w"] = int(sz.get(WN + "w", default["w"]))
            default["h"] = int(sz.get(WN + "h", default["h"]))
        if mar is not None:
            for k in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
                v = mar.get(WN + k)
                if v is not None:
                    default[k] = int(v)
        if cols is not None and cols.get(WN + "space") is not None:
            default["cols"] = int(cols.get(WN + "space"))
        if grid is not None and grid.get(WN + "linePitch") is not None:
            default["linePitch"] = int(grid.get(WN + "linePitch"))
        break
    return default


def detect_logo(tparas, tpl_rels_root):
    """自动找封面第一张图片的 rId 与对应 media 目标。"""
    rid = None
    for p in tparas[:25]:
        for blip in p.iter("{%s}blip" % A):
            e = blip.get(RN + "embed")
            if e:
                rid = e
                break
        if rid:
            break
    if not rid:
        return None
    target = None
    for rel in tpl_rels_root:
        if rel.get("Id") == rid:
            target = rel.get("Target")
            break
    if not target:
        return None
    media = target if target.startswith("word/media/") else "word/media/" + target.split("/")[-1]
    return {"template_rid": rid, "template_media": media,
            "output_media": "word/media/school_logo.png"}


def auto_config(src, template, output):
    """零配置：仅凭 src/template/output 推断全部封面参数。"""
    zo = zipfile.ZipFile(src)
    zt = zipfile.ZipFile(template)
    odoc = ET.fromstring(zo.read("word/document.xml"))
    obody = odoc.find(WN + "body")
    opar = obody.findall(WN + "p")

    new_title = ""
    title_idx = None
    for i, p in enumerate(opar):
        t = txt_of(p).strip()
        if t:
            new_title = t
            title_idx = i
            break
    if not new_title:
        raise RuntimeError("无法从已排版文档提取论文题目（第一个非空段落为空）")

    tdoc = ET.fromstring(zt.read("word/document.xml"))
    tparas = tdoc.find(WN + "body").findall(WN + "p")
    lo, hi = detect_cover_range(tparas)
    tpl_rels = ET.fromstring(zt.read("word/_rels/document.xml.rels"))
    src_rels = ET.fromstring(zo.read("word/_rels/document.xml.rels"))

    logo = detect_logo(tparas[:hi], tpl_rels)
    if logo:
        logo["output_rid"] = "rId%d" % (max_rid(src_rels) + 1)

    move = []
    if title_idx is not None and title_idx + 1 < len(opar):
        nxt = txt_of(opar[title_idx + 1]).strip()
        if nxt and len(nxt) <= 300 and any(
            k in nxt for k in ("声明", "只可用于", "不允许", "抄袭", "代写", "挪用")
        ):
            move.append(title_idx + 1)

    return {
        "src": src, "template": template, "output": output,
        "new_title": new_title,
        "cover_range": [lo, hi],
        "fields_to_blank": PERSONAL_LABELS,
        "logo": logo,
        "drop_paragraphs": [title_idx] if title_idx is not None else [],
        "move_to_cover": move,
        "cover_section": detect_cover_section(tparas[:hi]),
    }


def run(cfg):
    src, tpl, out = cfg["src"], cfg["template"], cfg["output"]
    zo = zipfile.ZipFile(src)
    zt = zipfile.ZipFile(tpl)
    parts = {n: zo.read(n) for n in zo.namelist()}

    odoc = ET.fromstring(parts["word/document.xml"])
    obody = odoc.find(WN + "body")
    original = obody.findall(WN + "p")

    drop_idx = set(cfg.get("drop_paragraphs", []))
    move_idx = set(cfg.get("move_to_cover", []))
    # 移到封面的段落也要从正文删除（否则正文出现重复声明）
    to_drop = {i: original[i] for i in (drop_idx | move_idx) if 0 <= i < len(original)}
    to_move = [copy.deepcopy(original[i]) for i in sorted(move_idx) if 0 <= i < len(original)]
    for para in to_move:
        ppr = para.find(WN + "pPr")
        if ppr is not None:
            old = ppr.find(WN + "sectPr")
            if old is not None:
                ppr.remove(old)
    for para in to_drop.values():
        if para in list(obody):
            obody.remove(para)

    tdoc = ET.fromstring(zt.read("word/document.xml"))
    tparas = tdoc.find(WN + "body").findall(WN + "p")
    lo, hi = cfg["cover_range"]
    cover = [copy.deepcopy(tparas[i]) for i in range(lo, min(hi, len(tparas)))]

    new_title = cfg.get("new_title", "")
    title_cfg = cfg.get("title", {})
    old_title = title_cfg.get("old")
    if not new_title:
        new_title = title_cfg.get("new", "")
    fields = cfg.get("fields_to_blank", [])
    logo_cfg = cfg.get("logo")

    rels_root = ET.fromstring(parts["word/_rels/document.xml.rels"])
    used_ids = {rel.get("Id") for rel in rels_root}
    new_rid = logo_cfg.get("output_rid") if logo_cfg else None
    old_rid = logo_cfg.get("template_rid") if logo_cfg else None
    if new_rid and new_rid in used_ids:
        new_rid = "rId%d" % (max_rid(rels_root) + 1)

    labels_sorted = sorted(fields, key=len, reverse=True)
    for para in cover:
        ts = list(para.iter(WN + "t"))
        if ts:
            full = "".join(t.text or "" for t in ts)
            if TITLE_LABEL_RE.match(full) and new_title:
                m = re.match(r"^(\s*题\s*目\s*[:：])\s*.*$", full, re.S)
                ts[0].text = (m.group(1) + " " + new_title) if m else new_title
                for t in ts[1:]:
                    t.text = ""
            elif old_title and new_title and old_title in full:
                ts[0].text = full.replace(old_title, new_title)
                for t in ts[1:]:
                    t.text = ""
            else:
                hit = None
                for lab in labels_sorted:
                    if norm(lab) in norm(full):
                        hit = lab
                        break
                if hit:
                    m = re.match(r"^(.*?[:：])", full, re.S)
                    ts[0].text = m.group(1) if m else hit
                    for t in ts[1:]:
                        t.text = ""
        if logo_cfg:
            for blip in para.iter("{%s}blip" % A):
                if blip.get(RN + "embed") == old_rid:
                    blip.set(RN + "embed", new_rid)
        ppr = para.find(WN + "pPr")
        if ppr is not None:
            sp = ppr.find(WN + "sectPr")
            if sp is not None:
                ppr.remove(sp)

    cover.extend(to_move)
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
            "moved_paragraphs": len(to_move), "logo_rid": new_rid,
            "title": new_title}


def main():
    ap = argparse.ArgumentParser(
        description="按模板重建封面：零配置(src template out) 或配置模式(config.json)")
    ap.add_argument("args", nargs="*", help="零配置: <src.docx> <template.docx> <out.docx>；或: <config.json>")
    ap.add_argument("--print-template", action="store_true")
    args = ap.parse_args()
    if args.print_template:
        print(json.dumps({
            "_说明": "可选覆盖配置。零配置命令无需此文件：python rebuild_cover.py src.docx template.docx out.docx",
            "title": {"old": "模板示例题目", "new": "论文题目"},
            "cover_range": [0, 18],
            "fields_to_blank": PERSONAL_LABELS,
            "logo": {"template_rid": "rId7", "template_media": "word/media/image1.png",
                     "output_media": "word/media/school_logo.png", "output_rid": "rId99"},
            "drop_paragraphs": [0], "move_to_cover": [1],
            "cover_section": {"w": 11906, "h": 16838, "top": 1440, "right": 1800,
                              "bottom": 1440, "left": 1800, "header": 851, "footer": 992,
                              "gutter": 0, "cols": 425, "linePitch": 312},
        }, ensure_ascii=False, indent=2))
        return 0

    argv = args.args
    if len(argv) == 1 and argv[0].lower().endswith(".json"):
        with open(argv[0], encoding="utf-8") as f:
            cfg = json.load(f)
        if not all(k in cfg for k in ("src", "template", "output", "new_title", "cover_range")):
            base = auto_config(cfg["src"], cfg["template"], cfg["output"])
            base.update(cfg)
            cfg = base
        res = run(cfg)
    elif len(argv) == 3:
        cfg = auto_config(argv[0], argv[1], argv[2])
        res = run(cfg)
    else:
        ap.error("用法: rebuild_cover.py <src.docx> <template.docx> <out.docx>  或  rebuild_cover.py <config.json>")

    print("=== 封面重建完成 ===")
    print("  输出:", res["output"])
    print("  题目:", res["title"])
    print("  封面段落:", res["cover_paragraphs"], " 保留原文段落:", res["moved_paragraphs"])
    print("  校徽 rId:", res["logo_rid"])
    print("  个人信息: 已清空（仅保留标签，由使用者自行填写）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
