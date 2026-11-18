"""Shared standard-library helpers for reading and writing .docx (OOXML) files.

No third-party dependencies: only zipfile + xml.etree.ElementTree.
"""

from __future__ import annotations

import copy
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Iterable, Optional, Tuple

# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

for _p, _u in NS.items():
    ET.register_namespace(_p, _u)

REL_TYPE = {
    "header": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    "footer": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    "image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    "numbering": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    "styles": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    "settings": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    "fontTable": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
    "theme": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
}

# Chinese 字号 -> half-points (w:sz uses half-points)
SIZE_NAMES = [
    (8, "七号"), (9, "小六"), (10.5, "五号"), (11, "小五"),
    (12, "小四"), (14, "四号"), (15, "小三"), (16, "三号"),
    (18, "小二"), (22, "二号"), (24, "小一"), (28, "一号"),
    (36, "小初"), (42, "初号"),
]


def size_name(pt: Optional[float]) -> str:
    if pt is None:
        return "(继承)"
    best = min(SIZE_NAMES, key=lambda x: abs(x[0] - pt))
    if abs(best[0] - pt) <= 0.6:
        return f"{best[1]}({pt:g}pt)"
    return f"{pt:g}pt"


def qn(tag: str) -> str:
    """Resolve a prefix:local tag to a Clark-notation tag."""
    if ":" in tag:
        p, local = tag.split(":", 1)
        return "{%s}%s" % (NS.get(p, p), local)
    return "{%s}%s" % (NS["w"], tag)


def local(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


# ---------------------------------------------------------------------------
# Package read / write
# ---------------------------------------------------------------------------
def read_docx(path: str) -> Dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as z:
        return {n: z.read(n) for n in z.namelist()}


def write_docx(path: str, parts: Dict[str, bytes]) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def serialize_xml(root: ET.Element, default_uri: Optional[str] = None) -> bytes:
    if default_uri:
        ET.register_namespace("", default_uri)
    else:
        # Prevent a stale default namespace from emitting a misleading prefix.
        ET.register_namespace("", "http://www.w3.org/XML/1998/namespace")
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def rels_path(part_name: str) -> str:
    d, b = os.path.split(part_name)
    return f"{d}/_rels/{b}.rels" if d else f"_rels/{b}.rels"


def rels_default_rels_path(part_name: str) -> str:
    return rels_path(part_name)


# ---------------------------------------------------------------------------
# Generic element helpers
# ---------------------------------------------------------------------------
def child(parent: ET.Element, tag: str) -> Optional[ET.Element]:
    return parent.find(qn(tag))


def children(parent: ET.Element, tag: str) -> list:
    return list(parent.findall(qn(tag)))


def attr(el: Optional[ET.Element], name: str, default=None):
    """Read an attribute.

    Prefixed names (e.g. 'w:val') resolve via the wordprocessingml namespace;
    unprefixed names (e.g. relationship 'Id', 'Type', 'Target') are read as-is.
    """
    if el is None:
        return default
    key = qn(name) if ":" in name else name
    return el.get(key, default)


def get_val(parent: ET.Element, path: str, default=None):
    """Walk a slash-separated path of w:-prefixed tags and read w:val.

    Uses manual traversal instead of ElementTree.find() so that the 'w'
    prefix does not require a namespaces map at query time.
    """
    el = parent
    for token in path.split("/"):
        if token == "" or token == ".":
            continue
        el = child(el, token)
        if el is None:
            return default
    return attr(el, "w:val", default)


def text_of(el: ET.Element) -> str:
    """Collect all w:t text under el, ignoring deleted text."""
    out = []
    for node in el.iter():
        if node.tag == qn("w:t"):
            out.append(node.text or "")
        elif node.tag == qn("w:tab"):
            out.append("\t")
        elif node.tag in (qn("w:br"), qn("w:cr")):
            out.append("\n")
    return "".join(out)


# ---------------------------------------------------------------------------
# Field detection (page number etc.)
# ---------------------------------------------------------------------------
_FIELD_RE = re.compile(r"\b(PAGE|NUMPAGES|SECTIONPAGES|CHAPTERPAGES)\b", re.I)


def fields_in(el: ET.Element) -> list:
    found = set()
    for node in el.iter():
        if node.tag == qn("w:instrText"):
            found.update(_FIELD_RE.findall(node.text or ""))
        if node.tag == qn("w:fldSimple"):
            found.update(_FIELD_RE.findall(attr(node, "w:instr", "") or ""))
        if node.tag == qn("w:fldChar"):
            pass
    return sorted(found)


def normalize_header_text(el: ET.Element) -> str:
    """Text with field results removed but field tokens kept, for comparison."""
    parts = []
    skip_depth = 0
    for node in el.iter():
        if node.tag == qn("w:fldChar"):
            t = attr(node, "w:fldCharType")
            if t == "begin":
                skip_depth += 1
        if skip_depth == 0 and node.tag == qn("w:t"):
            parts.append(node.text or "")
        if node.tag == qn("w:fldChar") and attr(node, "w:fldCharType") == "end":
            skip_depth = max(0, skip_depth - 1)
    return re.sub(r"\s+", " ", "".join(parts)).strip()


# ---------------------------------------------------------------------------
# Format extraction
# ---------------------------------------------------------------------------
def extract_run_props(rPr: Optional[ET.Element]) -> dict:
    out = {}
    if rPr is None:
        return out
    fonts = child(rPr, "w:rFonts")
    if fonts is not None:
        for k in ("ascii", "hAnsi", "eastAsia", "cs"):
            v = attr(fonts, "w:" + k)
            if v:
                out["font_" + k] = v
    sz = child(rPr, "w:sz")
    if sz is not None and attr(sz, "w:val"):
        out["size_pt"] = int(attr(sz, "w:val")) / 2
    b = child(rPr, "w:b")
    out["bold"] = attr(b, "w:val", "1") not in ("0", "false", "off") if b is not None else None
    i = child(rPr, "w:i")
    out["italic"] = attr(i, "w:val", "1") not in ("0", "false", "off") if i is not None else None
    color = child(rPr, "w:color")
    if color is not None:
        out["color"] = attr(color, "w:val")
    u = child(rPr, "w:u")
    if u is not None:
        out["underline"] = attr(u, "w:val", "single")
    return out


def extract_par_props(pPr: Optional[ET.Element]) -> dict:
    out = {}
    if pPr is None:
        return out
    jc = child(pPr, "w:jc")
    if jc is not None:
        out["alignment"] = attr(jc, "w:val")
    spacing = child(pPr, "w:spacing")
    if spacing is not None:
        for k in ("line", "lineRule", "before", "after", "beforeLines", "afterLines"):
            v = attr(spacing, "w:" + k)
            if v is not None:
                out[k] = v
    ind = child(pPr, "w:ind")
    if ind is not None:
        for k in ("firstLine", "firstLineChars", "left", "leftChars", "right", "rightChars", "hanging", "hangingChars"):
            v = attr(ind, "w:" + k)
            if v is not None:
                out[k] = v
    for flag in ("keepNext", "keepLines", "pageBreakBefore", "widowControl", "suppressAutoHyphens"):
        el = child(pPr, "w:" + flag)
        if el is not None:
            out[flag] = attr(el, "w:val", "1") not in ("0", "false", "off")
    outline = child(pPr, "w:outlineLvl")
    if outline is not None:
        out["outlineLvl"] = attr(outline, "w:val")
    numPr = child(pPr, "w:numPr")
    if numPr is not None:
        out["numId"] = get_val(numPr, "w:numId")
        out["ilvl"] = get_val(numPr, "w:ilvl")
    return out


def extract_styles(styles_root: ET.Element) -> dict:
    result = {}
    defaults = child(styles_root, "w:docDefaults")
    if defaults is not None:
        rpr_default = child(child(defaults, "w:rPrDefault"), "w:rPr")
        ppr_default = child(child(defaults, "w:pPrDefault"), "w:pPr")
        result["__docDefaults__"] = {
            "run": extract_run_props(rpr_default),
            "paragraph": extract_par_props(ppr_default),
        }
    for st in children(styles_root, "w:style"):
        sid = attr(st, "w:styleId")
        if not sid:
            continue
        item = {
            "type": attr(st, "w:type"),
            "name": get_val(st, "w:name"),
            "basedOn": get_val(st, "w:basedOn"),
            "default": attr(st, "w:default") in ("1", "true", "on"),
            "run": extract_run_props(child(st, "w:rPr")),
            "paragraph": extract_par_props(child(st, "w:pPr")),
        }
        result[sid] = item
    return result


def extract_sections(document_root: ET.Element, parts: Dict[str, bytes]) -> list:
    body = child(document_root, "w:body")
    sects = list(body.findall(qn("w:sectPr"))) if body is not None else []
    # sectPr may also be inside a paragraph's pPr for section breaks.
    for ppr in document_root.iter(qn("w:pPr")):
        s = child(ppr, "w:sectPr")
        if s is not None and s not in sects:
            sects.append(s)
    rels = {}
    try:
        rroot = parse_xml(parts["word/_rels/document.xml.rels"])
        for rel in rroot:
            rels[attr(rel, "Id")] = (attr(rel, "Type"), attr(rel, "Target"))
    except (KeyError, ET.ParseError):
        pass

    out = []
    for idx, sect in enumerate(sects):
        info = {"index": idx}
        pgsz = child(sect, "w:pgSz")
        if pgsz is not None:
            info["pgSz"] = {
                "w": attr(pgsz, "w:w"),
                "h": attr(pgsz, "w:h"),
                "orient": attr(pgsz, "w:orient", "portrait"),
            }
        pgmar = child(sect, "w:pgMar")
        if pgmar is not None:
            info["pgMar"] = {k: attr(pgmar, "w:" + k) for k in
                             ("top", "bottom", "left", "right", "header", "footer", "gutter")}
        cols = child(sect, "w:cols")
        if cols is not None:
            info["cols"] = {k: attr(cols, "w:" + k) for k in ("num", "space", "equalWidth")}
        pn = child(sect, "w:pgNumType")
        if pn is not None:
            info["pgNumType"] = {k: attr(pn, "w:" + k) for k in ("fmt", "start", "chapStyle")}
        title_pg = child(sect, "w:titlePg")
        info["titlePg"] = title_pg is not None
        for kind in ("header", "footer"):
            refs = []
            for ref in children(sect, "w:" + kind + "Reference"):
                typ = attr(ref, "w:type", "default")
                rid = attr(ref, "r:id")
                entry = {"type": typ, "rid": rid}
                if rid and rid in rels and rels[rid][1]:
                    target = rels[rid][1].lstrip("/")
                    part_name = target if target.startswith("word/") else "word/" + target
                    part_name = os.path.normpath(part_name).replace("\\", "/")
                    entry["part"] = part_name
                    if part_name in parts:
                        try:
                            root = parse_xml(parts[part_name])
                            paras = []
                            for p in root.findall(qn("w:p")):
                                paras.append({
                                    "text": normalize_header_text(p),
                                    "fields": fields_in(p),
                                    "alignment": get_val(p, "w:pPr/w:jc"),
                                    "style": get_val(p, "w:pPr/w:pStyle"),
                                })
                            entry["paragraphs"] = paras
                            entry["fields"] = fields_in(root)
                        except ET.ParseError:
                            pass
                refs.append(entry)
            info[kind + "Refs"] = refs
        out.append(info)
    return out


def extract_paragraphs(document_root: ET.Element) -> list:
    body = child(document_root, "w:body")
    out = []
    if body is None:
        return out
    for i, p in enumerate(body.findall(qn("w:p"))):
        out.append({
            "index": i,
            "style": get_val(p, "w:pPr/w:pStyle"),
            "text": re.sub(r"\s+", " ", text_of(p)).strip(),
            "props": extract_par_props(child(p, "w:pPr")),
        })
    return out


def build_spec(parts: Dict[str, bytes], filename: str = "") -> dict:
    doc_root = parse_xml(parts["word/document.xml"])
    spec = {
        "source_file": filename,
        "styles": {},
        "sections": [],
        "paragraphs": [],
    }
    if "word/styles.xml" in parts:
        spec["styles"] = extract_styles(parse_xml(parts["word/styles.xml"]))
    spec["sections"] = extract_sections(doc_root, parts)
    spec["paragraphs"] = extract_paragraphs(doc_root)
    return spec
