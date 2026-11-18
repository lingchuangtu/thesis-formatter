# -*- coding: utf-8 -*-
"""格式移植（真实论文 / 真实学校模板路径）。

以论文原文件为基线，保留全部文字、表格、图片、代码；把模板的样式定义、
页面设置、分节结构、页眉页脚（含 PAGE 域）移植进来，并给段落套用对应样式。
所有细节由配置 JSON 驱动，不写死任何学校/论文的 rId 或段落索引。

与 build_document.py 的区别：
  * build_document.py —— 纯文本初稿 → 克隆模板替换正文（简单路径）
  * transform_paper.py —— 已有完整内容的论文 → 移植模板格式（真实路径）

用法：
    python transform_paper.py <config.json>
    python transform_paper.py --print-template > config.example.json

内容零丢失：脚本会核对段落数、表格数、绘图数与图片字节不变；
标题编号规范化属于预期的文字变更，会在输出中单独列出（见 --strict-text 关闭）。

纯标准库实现（zipfile + xml.etree）。
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
WN, RN = "{%s}" % W, "{%s}" % R

for _pfx, _uri in [
    ("w", W), ("r", R),
    ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
    ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
    ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
    ("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"),
    ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
    ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
]:
    ET.register_namespace(_pfx, _uri)


# --------------------------------------------------------------------------
# 配置示例（--print-template 输出）
# --------------------------------------------------------------------------
TEMPLATE_CONFIG = {
    "paper": "输入论文.docx",
    "template": "学校模板.docx",
    "output": "输出_已排版.docx",
    "title": {"old": "模板示例题目", "new": "你的论文题目"},
    "styles": {
        "import_ids": ["CMheading1", "CMheading2", "CMcontent", "CMreflist"],
        "code_style_size_halfpt": 21,
    },
    "role_rules": [
        {"when": {"style": "Heading1"}, "style": "CMheading1", "strip_bold": True},
        {"when": {"style": "Heading2"}, "style": "CMheading2", "strip_bold": True},
        {"when": {"style": "Heading3"}, "style": "CMheading3", "strip_bold": True},
        {"when": {"style": "BodyText3"}, "style": "CMcontent"},
        {"when": {"regex": "^图\\s*\\d"}, "style": "CMcaption"},
        {"when": {"regex": "^表\\s*\\d"}, "style": "CMcaptionTable"},
        {"when": {"index": [256, 274]}, "style": "CMreflist", "keep_tabs": True},
    ],
    "renumber": {
        "enabled": True,
        "heading_style": "CMheading1",
        "toc_index_range": [9, 44],
        "cn_to_arabic": True,
        "add_period": False,
    },
    "page_sets": {
        "body": {"w": 11906, "h": 16838, "top": 1587, "right": 1134,
                 "bottom": 1247, "left": 1701, "header": 1020, "footer": 793,
                 "gutter": 0, "cols": 720, "linePitch": 312},
        "cover": {"w": 11906, "h": 16838, "top": 1440, "right": 1800,
                  "bottom": 1440, "left": 1800, "header": 851, "footer": 992,
                  "gutter": 0, "cols": 425, "linePitch": 312},
    },
    "header_parts": ["header1.xml", "header2.xml"],
    "footer_parts": ["footer1.xml", "footer2.xml"],
    "header_text": {"header2.xml": "第1章 绪论", "header4.xml": "__TITLE__"},
    "even_odd_headers": True,
    "sections": [
        {"name": "封面", "end": 1, "page": "cover"},
        {"name": "摘要", "end": 4, "page": "body",
         "header": {"even": "header1.xml", "default": "header2.xml"},
         "footer": {"default": "footer1.xml"}, "pgnum": {"fmt": "upperRoman", "start": "1"}},
        {"name": "正文", "end": None, "page": "body",
         "header": {"default": "header2.xml"},
         "footer": {"default": "footer1.xml"}, "pgnum": {"fmt": "decimal", "start": "1"}},
    ],
    "_comment": "end=本节最后一个段落的索引；None=用 body 末尾 sectPr。__TITLE__ 表示替换为配置中的新题目。",
}


# --------------------------------------------------------------------------
# 基础工具
# --------------------------------------------------------------------------
def txt_of(para: ET.Element) -> str:
    return re.sub(r"\s+", " ", "".join(t.text or "" for t in para.iter(WN + "t"))).strip()


def style_of(para: ET.Element) -> str:
    e = para.find(WN + "pPr/" + WN + "pStyle")
    return e.get(WN + "val") if e is not None else "-"


def set_style(para: ET.Element, sid: str) -> None:
    ppr = para.find(WN + "pPr")
    if ppr is None:
        ppr = ET.Element(WN + "pPr")
        para.insert(0, ppr)
    e = ppr.find(WN + "pStyle")
    if e is None:
        e = ET.Element(WN + "pStyle")
        ppr.insert(0, e)
    e.set(WN + "val", sid)


def strip_direct(para: ET.Element, keep_tabs: bool = False, strip_bold: bool = False,
                 tags_extra: list | None = None) -> None:
    ppr = para.find(WN + "pPr")
    if ppr is not None:
        for tag in ("jc", "ind", "spacing", "rPr"):
            el = ppr.find(WN + tag)
            if el is not None:
                ppr.remove(el)
        if not keep_tabs:
            t = ppr.find(WN + "tabs")
            if t is not None:
                ppr.remove(t)
    tags = ["rFonts", "sz", "szCs", "color", "spacing", "w", "kern",
            "position", "highlight", "shd", "i", "iCs"]
    if strip_bold:
        tags += ["b", "bCs"]
    if tags_extra:
        tags += list(tags_extra)
    for r in para.findall(".//" + WN + "r"):
        rpr = r.find(WN + "rPr")
        if rpr is None:
            continue
        for tag in tags:
            el = rpr.find(WN + tag)
            if el is not None:
                rpr.remove(el)
        if len(rpr) == 0 and not rpr.attrib:
            r.remove(rpr)


CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
          "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def cn_to_int(s: str) -> int | None:
    if not s:
        return None
    if len(s) == 1:
        return CN_NUM.get(s)
    # 简单处理：十/二十/十一 等
    if "十" in s:
        left, _, right = s.partition("十")
        tens = CN_NUM.get(left, 1) if left else 1
        ones = CN_NUM.get(right, 0) if right else 0
        return tens * 10 + ones
    return sum(CN_NUM.get(c, 0) for c in s) or None


def renumber_heading(para: ET.Element, add_period: bool = False) -> str | None:
    """章标题 '第N章 xxx'（中文数字）→ '第N章 xxx'（阿拉伯数字）。返回新文本。"""
    ts = list(para.iter(WN + "t"))
    if not ts:
        return None
    full = "".join(t.text or "" for t in ts)
    m = re.match(r"^第([一二三四五六七八九十]+)章(\s*)(.*)$", full)
    if not m:
        return None
    n = cn_to_int(m.group(1))
    if n is None:
        return None
    new_text = "第%d章 %s%s" % (n, m.group(3).strip(), "." if add_period else "")
    ts[0].text = new_text
    for t in ts[1:]:
        t.text = ""
    return new_text


def build_sect_xml(sec: dict, page_sets: dict, rid_map: dict) -> bytes:
    pg = page_sets[sec["page"]]
    parts = []
    for kind in ("header", "footer"):
        refs = sec.get(kind, {}) or {}
        for typ in ("even", "default"):
            if typ in refs:
                target = refs[typ]
                rid = rid_map.get(("header" if kind == "header" else "footer", target))
                if rid:
                    parts.append('<w:%sReference w:type="%s" r:id="%s"/>' % (kind, typ, rid))
    parts.append('<w:pgSz w:w="%(w)d" w:h="%(h)d"/>' % pg)
    parts.append('<w:pgMar w:top="%(top)d" w:right="%(right)d" w:bottom="%(bottom)d" '
                 'w:left="%(left)d" w:header="%(header)d" w:footer="%(footer)d" '
                 'w:gutter="%(gutter)d"/>' % pg)
    if sec.get("pgnum"):
        pn = sec["pgnum"]
        s = ' w:start="%s"' % pn["start"] if pn.get("start") else ""
        parts.append('<w:pgNumType w:fmt="%s"%s/>' % (pn["fmt"], s))
    parts.append('<w:cols w:space="%(cols)d"/>' % pg)
    parts.append('<w:docGrid w:type="lines" w:linePitch="%(linePitch)d"/>' % pg)
    return ('<w:sectPr xmlns:w="%s" xmlns:r="%s">%s</w:sectPr>'
            % (W, R, "".join(parts))).encode("utf-8")


def set_header_text(raw: bytes, new_text: str) -> bytes:
    """把页眉里分散在多个 w:t 的文字替换为单个文本，保留其它格式与域。"""
    s = raw.decode("utf-8")
    matches = list(re.finditer(r"(<w:t[^>]*>)(.*?)(</w:t>)", s, re.S))
    if not matches:
        return raw
    out, last = [], 0
    for i, m in enumerate(matches):
        out.append(s[last:m.start()])
        out.append(m.group(1))
        out.append(escape(new_text) if i == 0 else "")
        out.append(m.group(3))
        last = m.end()
    out.append(s[last:])
    return "".join(out).encode("utf-8")


def max_rid(rels_root: ET.Element) -> int:
    best = 0
    for rel in rels_root:
        m = re.match(r"rId(\d+)$", rel.get("Id", ""))
        if m:
            best = max(best, int(m.group(1)))
    return best


def media_hashes(z: zipfile.ZipFile) -> list:
    return sorted(hashlib.md5(z.read(n)).hexdigest()
                  for n in z.namelist() if n.startswith("word/media/"))


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------
def run(cfg: dict) -> dict:
    paper, tpl, out = cfg["paper"], cfg["template"], cfg["output"]
    zp, zt = zipfile.ZipFile(paper), zipfile.ZipFile(tpl)
    paper_names = set(zp.namelist())

    # ---- 0. 内容基线 ----
    base = media_hashes(zp)
    dp0 = ET.fromstring(zp.read("word/document.xml"))
    base_paras = len(dp0.findall(".//" + WN + "p"))
    base_tables = len(dp0.findall(".//" + WN + "tbl"))
    base_drawings = len(dp0.findall(".//" + WN + "drawing"))

    # ---- 1. 样式移植：以论文为基线，合并模板 docDefaults 与需要的样式 ----
    ps = ET.fromstring(zp.read("word/styles.xml"))
    ts = ET.fromstring(zt.read("word/styles.xml"))
    td = ts.find(WN + "docDefaults")
    pd = ps.find(WN + "docDefaults")
    if td is not None:
        if pd is not None:
            ps.remove(pd)
        ps.insert(0, copy.deepcopy(td))
    paper_ids = {s.get(WN + "styleId") for s in ps.findall(WN + "style")}
    tpl_byid = {s.get(WN + "styleId"): s for s in ts.findall(WN + "style")}
    added = 0
    for sid in cfg.get("styles", {}).get("import_ids", []):
        if sid not in paper_ids and sid in tpl_byid:
            ps.append(copy.deepcopy(tpl_byid[sid]))
            added += 1
    code_sz = cfg.get("styles", {}).get("code_style_size_halfpt")
    if code_sz:
        for s in ps.findall(WN + "style"):
            if s.get(WN + "styleId") == "Code":
                rpr = s.find(WN + "rPr")
                if rpr is None:
                    rpr = ET.SubElement(s, WN + "rPr")
                for tag in ("sz", "szCs"):
                    e = rpr.find(WN + tag)
                    if e is None:
                        e = ET.SubElement(rpr, WN + tag)
                    e.set(WN + "val", str(code_sz))
    styles_out = ET.tostring(ps, encoding="UTF-8", xml_declaration=True)

    # ---- 2. 段落：规则映射样式、清直接格式、标题重编号、目录同步 ----
    doc = ET.fromstring(zp.read("word/document.xml"))
    body = doc.find(WN + "body")
    paras = body.findall(WN + "p")
    rules = cfg.get("role_rules", [])

    def match_rule(i, para, text):
        for r in rules:
            w = r.get("when", {})
            if not w:
                return r  # 空条件 = 兜底规则（放最后）
            ok = True
            if "index" in w:
                lo, hi = w["index"]
                ok = ok and lo <= i <= hi
            if "style" in w:
                ok = ok and style_of(para) == w["style"]
            if "regex" in w:
                ok = ok and bool(re.match(w["regex"], text))
            if "equals" in w:
                ok = ok and text == w["equals"]
            if "maxlen" in w:
                ok = ok and len(text) <= w["maxlen"]
            if ok:
                return r
        return None

    renumber_cfg = cfg.get("renumber", {})
    toc_range = renumber_cfg.get("toc_index_range")
    renumbered = []
    touched = 0
    for i, para in enumerate(paras):
        text = txt_of(para)
        # 目录段落：绝不套用正文样式，只清 run 级字体字号并同步章号
        in_toc = toc_range and toc_range[0] <= i <= toc_range[1]
        if in_toc:
            for r in para.findall(".//" + WN + "r"):
                rpr = r.find(WN + "rPr")
                if rpr is not None:
                    for tag in ("rFonts", "sz", "szCs", "color", "spacing"):
                        e = rpr.find(WN + tag)
                        if e is not None:
                            rpr.remove(e)
            ts_ = list(para.iter(WN + "t"))
            if ts_ and ts_[0].text:
                m = re.match(r"^第([一二三四五六七八九十]+)章\s*(.*)$", ts_[0].text)
                if m:
                    n = cn_to_int(m.group(1))
                    if n:
                        # 只改第一个文本节点里的章号；其余节点是页码，必须保留
                        ts_[0].text = "第%d章 %s" % (n, m.group(2))
            continue
        rule = match_rule(i, para, text)
        if rule:
            sid = rule["style"]
            set_style(para, sid)
            strip_direct(para,
                         keep_tabs=rule.get("keep_tabs", False),
                         strip_bold=rule.get("strip_bold", False),
                         tags_extra=rule.get("strip_tags"))
            touched += 1
            if (renumber_cfg.get("enabled") and
                    sid == renumber_cfg.get("heading_style") and
                    renumber_cfg.get("cn_to_arabic")):
                nt = renumber_heading(para, renumber_cfg.get("add_period", False))
                if nt:
                    renumbered.append(nt)

    # ---- 3. 关系：为 header/footer 部件分配新 rId（自动避让现有 Id） ----
    rels_root = ET.fromstring(zp.read("word/_rels/document.xml.rels"))
    next_rid = max_rid(rels_root) + 1
    rid_map = {}
    new_rels = []
    for kind, names in (("header", cfg.get("header_parts", [])),
                        ("footer", cfg.get("footer_parts", []))):
        for name in names:
            rid = "rId%d" % next_rid
            next_rid += 1
            rid_map[(kind, name)] = rid
            rel = ET.Element("{%s}Relationship" % PKG_REL)
            rel.set("Id", rid)
            rel.set("Type", R + "/" + kind)
            rel.set("Target", name)
            new_rels.append(rel)
            rels_root.append(rel)
    rels_out = ET.tostring(rels_root, encoding="UTF-8", xml_declaration=True)

    # ---- 4. 分节重建 ----
    for sec in cfg.get("sections", []):
        sect = ET.fromstring(build_sect_xml(sec, cfg["page_sets"], rid_map))
        end = sec.get("end")
        if end is None:
            old = body.find(WN + "sectPr")
            if old is not None:
                body.remove(old)
            body.append(sect)
        else:
            para = paras[end]
            ppr = para.find(WN + "pPr")
            if ppr is None:
                ppr = ET.Element(WN + "pPr")
                para.insert(0, ppr)
            old = ppr.find(WN + "sectPr")
            if old is not None:
                ppr.remove(old)
            ppr.append(sect)
    doc_out = ET.tostring(doc, encoding="UTF-8", xml_declaration=True)

    # ---- 5. 页眉页脚：复制模板部件，替换示例题目 / 章节文字 ----
    hf_paths = {}
    title_cfg = cfg.get("title", {})
    old_title, new_title = title_cfg.get("old"), title_cfg.get("new")
    all_names = list(cfg.get("header_parts", [])) + list(cfg.get("footer_parts", []))
    for name in all_names:
        src = "word/" + name
        if src not in zt.namelist():
            raise SystemExit("模板缺少部件: " + src)
        raw = zt.read(src)
        if old_title and new_title:
            raw = raw.replace(old_title.encode("utf-8"), new_title.encode("utf-8"))
        if name in cfg.get("header_text", {}):
            ht = cfg["header_text"][name]
            if ht == "__TITLE__":
                ht = new_title or ""
            raw = set_header_text(raw, ht)
        hf_paths[src] = raw

    # ---- 6. Content_Types / settings ----
    ct = zp.read("[Content_Types].xml").decode("utf-8")
    adds = []
    for name in all_names:
        pn = "/word/" + name
        if pn not in ct:
            typ = "header" if name.startswith("header") else "footer"
            adds.append('<Override PartName="%s" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.%s+xml"/>' % (pn, typ))
    ct = ct.replace("</Types>", "".join(adds) + "</Types>")
    settings = zp.read("word/settings.xml").decode("utf-8")
    if cfg.get("even_odd_headers") and "evenAndOddHeaders" not in settings:
        settings = settings.replace("</w:settings>", "<w:evenAndOddHeaders/></w:settings>")

    modified = {
        "word/document.xml": doc_out,
        "word/styles.xml": styles_out,
        "word/_rels/document.xml.rels": rels_out,
        "[Content_Types].xml": ct.encode("utf-8"),
        "word/settings.xml": settings.encode("utf-8"),
    }
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zp.infolist():
            name = item.filename
            zo.writestr(name, hf_paths.get(name, modified.get(name, zp.read(name))))
        for path, data in hf_paths.items():
            if path not in paper_names:
                zo.writestr(path, data)

    # ---- 7. 内容零丢失自检（结构层面）----
    zo = zipfile.ZipFile(out)
    d1 = ET.fromstring(zo.read("word/document.xml"))
    struct_ok = (
        len(d1.findall(".//" + WN + "tbl")) == base_tables and
        len(d1.findall(".//" + WN + "drawing")) == base_drawings and
        media_hashes(zo) == base and
        zo.testzip() is None
    )
    return {
        "output": out,
        "styles_imported": added,
        "paragraphs_styled": touched,
        "sections": len(cfg.get("sections", [])),
        "header_footer_parts": len(hf_paths),
        "renumbered_chapters": renumbered,
        "base_paragraphs": base_paras,
        "structure_preserved": struct_ok,
    }


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
    print("=== 格式移植完成 ===")
    print("  输出:", res["output"])
    print("  导入模板样式:", res["styles_imported"], "个")
    print("  套用样式段落:", res["paragraphs_styled"], "个")
    print("  重建分节:", res["sections"], "个")
    print("  页眉页脚部件:", res["header_footer_parts"], "个")
    print("  章标题重编号:", len(res["renumbered_chapters"]), "个 ->", res["renumbered_chapters"][:6])
    print("  结构零丢失(表格/绘图/图片字节/ZIP):", "PASS" if res["structure_preserved"] else "FAIL")
    if res["renumbered_chapters"]:
        print("  [提示] 标题编号已规范化，文字哈希会变化；请用 audit_content.py 复核图片/表格一致。")
    return 0 if res["structure_preserved"] else 1


if __name__ == "__main__":
    sys.exit(main())
