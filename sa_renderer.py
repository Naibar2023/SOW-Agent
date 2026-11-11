"""Builds a Word (.docx) SOW in the Osool house format.

The renderer opens the base template (a real Osool SOW), keeps its cover page,
header, footer, styles and numbering, removes the old body, and writes the new
content by cloning the template's own paragraphs and tables. That way the output
looks the same as a SOW prepared by hand in the house format.

Input is a plain dict (the "SOW JSON"), described in SOW_SCHEMA_TEXT below.
"""

from __future__ import annotations

import copy
import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def _resource(folder, name):
    """Find a data file whether the program is laid out in folders or as flat files."""
    here = Path(__file__).resolve().parent
    for base in (here.parent, here):
        for p in (base / folder / name, base / name):
            if p.is_file():
                return p
    return here.parent / folder / name


DEFAULT_TEMPLATE = _resource("templates", "osool_sow_base.docx")

MAX_WEEK_COLUMNS = 17    # W0..W16; longer programmes use months
MAX_MONTH_COLUMNS = 24

SOW_SCHEMA_TEXT = """\
A SOW is one JSON object:

{
  "cover_title": "RGB Roof Top Lighting – Granada Business Park",   // shown on the cover page and in the page header
  "document_title": "Scope of Work (SOW)",                           // optional, first line of page 2
  "sections": [ { "heading": "Introduction", "blocks": [ ... ] }, ... ]
}

Section headings are written WITHOUT a letter. The renderer letters them A., B., C. ... in order.
To refer to another section inside any text write {{ref:Penalties}}; it becomes "Section M (Penalties)".

Block types (each block is an object with a "type"):

  {"type": "paragraph", "text": "..."}
  {"type": "bullets",   "items": ["text", {"text": "sub point", "level": 1}]}          // level 0-3
  {"type": "numbered",  "items": ["text", {"text": "a. item", "level": 1},             // level 0 = 1.  level 1 = a.  level 2 = i.  level 3 = 1.
                                  {"text": "a bullet under a number", "level": 1, "bullet": true}],
                        "start": 1}                     // optional: first number, to continue a list interrupted by a paragraph
  {"type": "subheading", "text": "Site Survey and Existing Condition Assessment"}       // numbered 1., 2., 3. within the section
  {"type": "subheading", "text": "Pre-commissioning", "level": 2}                       // lettered a., b. under the current subheading
  {"type": "table", "columns": ["Item / Parameter", "Minimum Requirement"],
                    "rows": [["Colour", "RGBW preferred"], ...],
                    "widths": [25, 75],                 // optional, percentages
                    "align": ["left", "left"],          // optional, per column: left | center | right
                    "bold_last_row": false}             // optional, for a totals row
  {"type": "schedule", "unit": "week",                  // "week" (W0, W1, ... grouped in months) or "month" (M1, M2, ...)
                       "activities": [{"name": "Purchase Order Receipt & Project Kick-Off", "start": 0, "end": 0},
                                      {"name": "Site Survey", "start": 1, "end": 2}]}
                       // week numbers start at 0 (W0 = kick-off week); month numbers start at 1
  {"type": "note", "text": "..."}                       // italic underlined note

Inline formatting inside any text: **bold**, *italic*. Inside a table cell, "\\n" starts a new line.
No other markup (no markdown headings, no HTML).
"""


class SOWValidationError(ValueError):
    """Raised when the SOW JSON does not follow the schema."""


# --------------------------------------------------------------------------- validation

def _is_str(v):
    return isinstance(v, str) and v.strip() != ""


def validate_sow(sow) -> list[str]:
    """Return a list of human-readable problems. Empty list means the SOW is valid."""
    errs: list[str] = []
    if not isinstance(sow, dict):
        return ["The SOW must be a JSON object."]
    if not _is_str(sow.get("cover_title")):
        errs.append('"cover_title" is required (short project title for the cover page).')
    secs = sow.get("sections")
    if not isinstance(secs, list) or not secs:
        errs.append('"sections" must be a non-empty list.')
        return errs
    seen = set()
    for si, sec in enumerate(secs):
        where = f"sections[{si}]"
        if not isinstance(sec, dict) or not _is_str(sec.get("heading")):
            errs.append(f'{where}: needs a "heading".')
            continue
        where = f'section "{sec["heading"]}"'
        key = sec["heading"].strip().lower()
        if key in seen:
            errs.append(f"{where}: duplicate section heading.")
        seen.add(key)
        if re.match(r"^[A-Z]{1,2}[.)]\s", sec["heading"]):
            errs.append(f"{where}: remove the letter prefix; sections are lettered automatically.")
        blocks = sec.get("blocks")
        if not isinstance(blocks, list):
            errs.append(f'{where}: "blocks" must be a list.')
            continue
        for bi, b in enumerate(blocks):
            bw = f"{where}, block {bi + 1}"
            if not isinstance(b, dict):
                errs.append(f"{bw}: must be an object.")
                continue
            t = b.get("type")
            if t in ("paragraph", "note"):
                if not _is_str(b.get("text")):
                    errs.append(f'{bw} ({t}): "text" is required.')
            elif t == "subheading":
                if not _is_str(b.get("text")):
                    errs.append(f'{bw} (subheading): "text" is required.')
                if b.get("level", 1) not in (1, 2):
                    errs.append(f'{bw} (subheading): "level" must be 1 or 2.')
            elif t in ("bullets", "numbered"):
                items = b.get("items")
                if not isinstance(items, list) or not items:
                    errs.append(f'{bw} ({t}): "items" must be a non-empty list.')
                    continue
                st = b.get("start", 1)
                if not isinstance(st, int) or isinstance(st, bool) or st < 1:
                    errs.append(f'{bw} ({t}): "start" must be a whole number, 1 or more.')
                for ii, it in enumerate(items):
                    if isinstance(it, str):
                        if not it.strip():
                            errs.append(f"{bw} ({t}): item {ii + 1} is empty.")
                    elif isinstance(it, dict):
                        if not _is_str(it.get("text")):
                            errs.append(f'{bw} ({t}): item {ii + 1} needs "text".')
                        lv = it.get("level", 0)
                        if not isinstance(lv, int) or isinstance(lv, bool) or not 0 <= lv <= 3:
                            errs.append(f'{bw} ({t}): item {ii + 1} "level" must be 0-3.')
                    else:
                        errs.append(f"{bw} ({t}): item {ii + 1} must be text or an object.")
            elif t == "table":
                cols = b.get("columns")
                rows = b.get("rows")
                if not isinstance(cols, list) or not cols or not all(isinstance(c, str) for c in cols):
                    errs.append(f'{bw} (table): "columns" must be a list of header texts.')
                    continue
                if len(cols) > 8:
                    errs.append(f"{bw} (table): at most 8 columns fit on the page.")
                if not isinstance(rows, list) or not rows:
                    errs.append(f'{bw} (table): "rows" must be a non-empty list.')
                    continue
                for ri, r in enumerate(rows):
                    if not isinstance(r, list) or len(r) != len(cols):
                        errs.append(f"{bw} (table): row {ri + 1} must have {len(cols)} cells.")
                for opt in ("widths", "align"):
                    if opt in b and (not isinstance(b[opt], list) or len(b[opt]) != len(cols)):
                        errs.append(f'{bw} (table): "{opt}" must have one entry per column.')
                if isinstance(b.get("widths"), list) and not all(
                        isinstance(w, (int, float)) and not isinstance(w, bool) and w > 0 for w in b["widths"]):
                    errs.append(f'{bw} (table): "widths" must be positive numbers.')
            elif t == "schedule":
                unit = b.get("unit", "week")
                if unit not in ("week", "month"):
                    errs.append(f'{bw} (schedule): "unit" must be "week" or "month".')
                    continue
                acts = b.get("activities")
                if not isinstance(acts, list) or not acts:
                    errs.append(f'{bw} (schedule): "activities" must be a non-empty list.')
                    continue
                first = 0 if unit == "week" else 1
                last = first
                for ai, a in enumerate(acts):
                    ok = (isinstance(a, dict) and _is_str(a.get("name"))
                          and all(isinstance(a.get(k), int) and not isinstance(a.get(k), bool)
                                  for k in ("start", "end")))
                    if not ok:
                        errs.append(f'{bw} (schedule): activity {ai + 1} needs "name", "start", "end" (whole numbers).')
                        continue
                    if a["start"] < first or a["end"] < a["start"]:
                        errs.append(f"{bw} (schedule): activity {ai + 1} has an invalid start/end "
                                    f"({unit}s start at {first}).")
                    last = max(last, a["end"])
                n = last - first + 1
                limit = MAX_WEEK_COLUMNS if unit == "week" else MAX_MONTH_COLUMNS
                if n > limit:
                    hint = ' Use "unit": "month" for long projects.' if unit == "week" else ""
                    errs.append(f"{bw} (schedule): {n} {unit} columns do not fit on the page (max {limit}).{hint}")
            else:
                errs.append(f'{bw}: unknown block type "{t}".')
    return errs


# --------------------------------------------------------------------------- XML helpers

# Child order required by the Word schema inside <w:rPr> (subset we may insert).
_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
              "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
              "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
              "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
              "specVanish", "oMath"]
# Child order inside <w:pPr> (subset).
_PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
              "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
              "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
              "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap",
              "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId",
              "cnfStyle", "rPr", "sectPr", "pPrChange"]


def _local(el):
    return el.tag.split("}")[1] if "}" in el.tag else el.tag


def _put(parent, child, order):
    """Insert child into parent at the schema-correct position, replacing a same-named child."""
    name = _local(child)
    for old in parent.findall(qn("w:" + name)):
        parent.remove(old)
    idx = order.index(name) if name in order else len(order)
    pos = 0
    for i, c in enumerate(parent):
        n = _local(c)
        if (order.index(n) if n in order else len(order)) <= idx:
            pos = i + 1
    parent.insert(pos, child)
    return child


def _el(name, **attrs):
    e = OxmlElement("w:" + name)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def _text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


_INLINE = re.compile(r"(\*\*.+?\*\*|(?<![\*\w])\*(?!\s)[^*]+?(?<!\s)\*(?![\*\w]))", re.S)


def _segments(text):
    """Split text with **bold** / *italic* markers into (text, bold, italic) pieces."""
    out = []
    for part in _INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            inner = part[2:-2]
            m = re.fullmatch(r"\*(.+)\*", inner, re.S)
            out.append((m.group(1), True, True) if m else (inner, True, False))
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            out.append((part[1:-1], False, True))
        else:
            out.append((part, False, False))
    return out


_BAD_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _clean(text):
    return _BAD_XML.sub("", str(text)).replace("\r\n", "\n").replace("\r", "\n")


class _Builder:
    def __init__(self, template_path):
        self.doc = Document(str(template_path))
        self.body = self.doc.element.body
        self.sect = self.body.find(qn("w:sectPr"))
        if self.sect is None:
            raise RuntimeError("Template has no section properties.")
        self._find_prototypes()
        self._clear_body()
        self.page_width = self._text_width()
        self.refs: dict[str, str] = {}

    # ---- template analysis

    def _find_prototypes(self):
        kids = list(self.body)
        self.cover = kids[0] if kids and kids[0].tag == qn("w:sdt") else None
        p_h2 = p_body = p_bullet = p_h4 = p_spacer = tbl = None
        for el in kids:
            if el.tag == qn("w:p"):
                ppr = el.find(qn("w:pPr"))
                style = ppr.find(qn("w:pStyle")) if ppr is not None else None
                sval = style.get(qn("w:val")) if style is not None else None
                has_num = ppr is not None and ppr.find(qn("w:numPr")) is not None
                txt = _text_of(el).strip()
                if sval == "Heading2" and p_h2 is None:
                    p_h2 = el
                elif sval == "Heading4" and has_num and p_h4 is None:
                    p_h4 = el
                elif has_num and sval in (None, "ListParagraph") and p_bullet is None and txt:
                    p_bullet = el
                elif sval is None and not has_num and txt and p_body is None:
                    p_body = el
                elif sval is None and not has_num and not txt and p_spacer is None \
                        and el.find(".//" + qn("w:drawing")) is None:
                    p_spacer = el
            elif el.tag == qn("w:tbl") and tbl is None:
                rows = el.findall(qn("w:tr"))
                if len(rows) >= 2 and rows[0].find(".//" + qn("w:shd")) is not None:
                    tbl = el
        missing = [n for n, v in (("section heading", p_h2), ("body paragraph", p_body),
                                  ("bullet", p_bullet), ("sub-heading", p_h4),
                                  ("empty line", p_spacer), ("table", tbl)) if v is None]
        if missing:
            raise RuntimeError("Template is missing sample elements: " + ", ".join(missing))
        self.p_h2, self.p_body, self.p_bullet = map(copy.deepcopy, (p_h2, p_body, p_bullet))
        self.p_h4, self.p_spacer = copy.deepcopy(p_h4), copy.deepcopy(p_spacer)
        rows = tbl.findall(qn("w:tr"))
        self.tbl_pr = copy.deepcopy(tbl.find(qn("w:tblPr")))
        self.tr_head, self.tr_body = copy.deepcopy(rows[0]), copy.deepcopy(rows[1])
        self.tc_head = copy.deepcopy(rows[0].find(qn("w:tc")))
        self.tc_body = copy.deepcopy(rows[1].find(qn("w:tc")))
        self.body_rpr = self._first_rpr(self.p_body)
        self.accent = "007889"
        shd = self.tc_head.find(".//" + qn("w:shd"))
        if shd is not None and shd.get(qn("w:fill")):
            self.accent = shd.get(qn("w:fill"))

        # numbering definitions
        self.numbering = self.doc.part.numbering_part.element
        num_to_abs = {n.get(qn("w:numId")): n.find(qn("w:abstractNumId")).get(qn("w:val"))
                      for n in self.numbering.findall(qn("w:num"))}
        self.bullet_num = p_bullet.find(".//" + qn("w:numId")).get(qn("w:val"))
        self.h4_abs = num_to_abs[p_h4.find(".//" + qn("w:numId")).get(qn("w:val"))]
        self.num_abs = None
        want = ["decimal", "lowerLetter", "lowerRoman", "decimal"]
        for a in self.numbering.findall(qn("w:abstractNum")):
            aid = a.get(qn("w:abstractNumId"))
            lv = a.findall(qn("w:lvl"))[:4]
            fm = [l.find(qn("w:numFmt")).get(qn("w:val")) for l in lv]
            tx = [l.find(qn("w:lvlText")).get(qn("w:val")) for l in lv]
            if fm == want and tx == ["%1.", "%2.", "%3.", "%4."] and aid != self.h4_abs:
                self.num_abs = aid
                break
        if self.num_abs is None:
            self.num_abs = self.h4_abs

    @staticmethod
    def _first_rpr(p):
        for r in p.findall(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is not None:
                return copy.deepcopy(rpr)
        return None

    def _clear_body(self):
        for el in list(self.body):
            if el is self.sect or el is self.cover:
                continue
            self.body.remove(el)

    def _text_width(self):
        sz, mar = self.sect.find(qn("w:pgSz")), self.sect.find(qn("w:pgMar"))
        return int(sz.get(qn("w:w"))) - int(mar.get(qn("w:left"))) - int(mar.get(qn("w:right")))

    def _new_num(self, abstract_id, start=1):
        """Create a fresh numbering instance so each list restarts (at 1 unless told otherwise)."""
        nums = self.numbering.findall(qn("w:num"))
        new_id = max(int(n.get(qn("w:numId"))) for n in nums) + 1
        num = _el("num", numId=new_id)
        num.append(_el("abstractNumId", val=abstract_id))
        for lvl in range(4):
            ov = _el("lvlOverride", ilvl=lvl)
            ov.append(_el("startOverride", val=start if lvl == 0 else 1))
            num.append(ov)
        nums[-1].addnext(num)
        return str(new_id)

    # ---- text

    def _resolve(self, text):
        text = _clean(text)

        def sub(m):
            key = m.group(1).strip().lower()
            if key in self.refs:
                return self.refs[key]
            for k, v in self.refs.items():
                if k.startswith(key) or key in k:
                    return v
            return m.group(1).strip()
        return re.sub(r"\{\{\s*ref\s*:\s*(.+?)\s*\}\}", sub, text)

    def _run(self, text, base_rpr, bold=False, italic=False, underline=False, color=None, size=None):
        r = OxmlElement("w:r")
        rpr = copy.deepcopy(base_rpr) if base_rpr is not None else OxmlElement("w:rPr")
        if bold:
            _put(rpr, _el("b"), _RPR_ORDER)
            _put(rpr, _el("bCs"), _RPR_ORDER)
        if italic:
            _put(rpr, _el("i"), _RPR_ORDER)
            _put(rpr, _el("iCs"), _RPR_ORDER)
        if underline:
            _put(rpr, _el("u", val="single"), _RPR_ORDER)
        if color:
            _put(rpr, _el("color", val=color), _RPR_ORDER)
        if size:
            _put(rpr, _el("sz", val=size), _RPR_ORDER)
            _put(rpr, _el("szCs", val=size), _RPR_ORDER)
        if len(rpr):
            r.append(rpr)
        t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        r.append(t)
        return r

    def _fill(self, p, text, base_rpr, **fmt):
        """Replace the runs of paragraph p with formatted text."""
        for child in list(p):
            if child.tag != qn("w:pPr"):
                p.remove(child)
        force_b, force_i = fmt.pop("bold", False), fmt.pop("italic", False)
        for seg, b, i in _segments(self._resolve(text)):
            lines = seg.split("\n")
            for li, line in enumerate(lines):
                if li:
                    br = OxmlElement("w:r")
                    br.append(OxmlElement("w:br"))
                    p.append(br)
                if line:
                    p.append(self._run(line, base_rpr, bold=b or force_b, italic=i or force_i, **fmt))
        return p

    def _add(self, el):
        self.sect.addprevious(el)
        return el

    # ---- blocks

    def spacer(self):
        self._add(copy.deepcopy(self.p_spacer))

    def heading(self, text):
        p = copy.deepcopy(self.p_h2)
        self._add(self._fill(p, text, None))

    def paragraph(self, text, **fmt):
        p = copy.deepcopy(self.p_body)
        return self._add(self._fill(p, text, self.body_rpr, **fmt))

    def note(self, text):
        p = self.paragraph(text, italic=True, underline=True)
        _put(p.find(qn("w:pPr")), _el("ind", left=720), _PPR_ORDER)

    def subheading(self, text, level, num_id):
        p = copy.deepcopy(self.p_h4)
        rpr = self._first_rpr(p)
        p.find(".//" + qn("w:numId")).set(qn("w:val"), num_id)
        p.find(".//" + qn("w:ilvl")).set(qn("w:val"), str(level - 1))
        self._add(self._fill(p, text, rpr))

    def list_block(self, items, numbered, start=1):
        num_id = self._new_num(self.num_abs, start) if numbered else None
        for it in items:
            if isinstance(it, str):
                it = {"text": it}
            level = int(it.get("level", 0))
            as_bullet = (not numbered) or bool(it.get("bullet"))
            p = copy.deepcopy(self.p_bullet)
            ppr = p.find(qn("w:pPr"))
            _put(ppr, _el("pStyle", val="ListParagraph"), _PPR_ORDER)
            ppr.find(".//" + qn("w:numId")).set(qn("w:val"), self.bullet_num if as_bullet else num_id)
            ppr.find(".//" + qn("w:ilvl")).set(qn("w:val"), str(level))
            self._add(self._fill(p, it["text"], self.body_rpr))

    # ---- tables

    def _cell(self, proto, width, text, align=None, span=None, vmerge=None, tight=False, **fmt):
        tc = copy.deepcopy(proto)
        tcpr = tc.find(qn("w:tcPr"))
        tcw = tcpr.find(qn("w:tcW"))
        tcw.set(qn("w:w"), str(int(width)))
        tcw.set(qn("w:type"), "dxa")
        if span and span > 1:
            tcw.addnext(_el("gridSpan", val=span))
        if vmerge:
            anchor = tcpr.find(qn("w:gridSpan"))
            anchor = anchor if anchor is not None else tcw
            vm = _el("vMerge", val="restart") if vmerge == "restart" else _el("vMerge")
            anchor.addnext(vm)
        if tight:
            mar = tcpr.find(qn("w:tcMar"))
            if mar is not None:
                for side in ("left", "right"):
                    m = mar.find(qn("w:" + side))
                    if m is not None:
                        m.set(qn("w:w"), "15")
        if tcpr.find(qn("w:vAlign")) is None:
            # vAlign must come after tcMar / before hideMark
            anchor = None
            for name in ("tcMar", "noWrap", "shd", "tcBorders", "vMerge", "gridSpan", "tcW"):
                anchor = tcpr.find(qn("w:" + name))
                if anchor is not None:
                    break
            anchor.addnext(_el("vAlign", val="center"))
        paras = tc.findall(qn("w:p"))
        proto_p = paras[0]
        for extra in paras[1:]:
            tc.remove(extra)
        rpr = self._first_rpr(proto_p)
        if rpr is None:
            mark = proto_p.find(qn("w:pPr") + "/" + qn("w:rPr"))
            rpr = copy.deepcopy(mark) if mark is not None else None
        lines = _clean(text).split("\n") if text is not None else [""]
        for li, line in enumerate(lines):
            p = proto_p if li == 0 else copy.deepcopy(proto_p)
            if li:
                tc.append(p)
            self._fill(p, line, rpr, **fmt)
            if align:
                ppr = p.find(qn("w:pPr"))
                if ppr is None:
                    ppr = OxmlElement("w:pPr")
                    p.insert(0, ppr)
                _put(ppr, _el("jc", val={"centre": "center"}.get(align, align)), _PPR_ORDER)
        return tc

    def _table(self, widths):
        tbl = OxmlElement("w:tbl")
        pr = copy.deepcopy(self.tbl_pr)
        tw = pr.find(qn("w:tblW"))
        tw.set(qn("w:w"), str(sum(widths)))
        tw.set(qn("w:type"), "dxa")
        tbl.append(pr)
        grid = OxmlElement("w:tblGrid")
        for w in widths:
            grid.append(_el("gridCol", w=w))
        tbl.append(grid)
        return tbl

    def _row(self, proto, header=False):
        tr = OxmlElement("w:tr")
        trpr = proto.find(qn("w:trPr"))
        trpr = copy.deepcopy(trpr) if trpr is not None else OxmlElement("w:trPr")
        if trpr.find(qn("w:cantSplit")) is None:
            trpr.insert(0, _el("cantSplit"))
        if header and trpr.find(qn("w:tblHeader")) is None:
            trpr.append(_el("tblHeader"))
        tr.append(trpr)
        return tr

    def _col_widths(self, block):
        n = len(block["columns"])
        total = self.page_width
        if block.get("widths"):
            raw = [float(w) for w in block["widths"]]
        else:
            raw = []
            for c in range(n):
                cells = [block["columns"][c]] + [str(r[c]) for r in block["rows"]]
                longest = max(max((len(line) for line in _clean(x).split("\n")), default=0) for x in cells)
                raw.append(min(max(longest, 5), 55))
        s = sum(raw)
        widths = [int(total * w / s) for w in raw]
        # no column narrower than its longest header word
        mins = [max(420, 280 + 105 * max((len(wd) for wd in _clean(h).split()), default=1))
                for h in block["columns"]]
        if sum(mins) > total:
            mins = [420] * n
        for _ in range(n * 4):
            short = [i for i in range(n) if widths[i] < mins[i]]
            if not short:
                break
            i = short[0]
            need = mins[i] - widths[i]
            donor = max(range(n), key=lambda j: widths[j] - mins[j])
            give = min(need, widths[donor] - mins[donor])
            if give <= 0:
                break
            widths[i] += give
            widths[donor] -= give
        widths[widths.index(max(widths))] += total - sum(widths)
        return widths

    def table(self, block):
        widths = self._col_widths(block)
        aligns = block.get("align") or [None] * len(widths)
        tbl = self._table(widths)
        tr = self._row(self.tr_head, header=True)
        for w, text in zip(widths, block["columns"]):
            tr.append(self._cell(self.tc_head, w, text))
        tbl.append(tr)
        last = len(block["rows"]) - 1
        for ri, row in enumerate(block["rows"]):
            tr = self._row(self.tr_body)
            bold = bool(block.get("bold_last_row")) and ri == last
            for w, text, al in zip(widths, row, aligns):
                tr.append(self._cell(self.tc_body, w, "" if text is None else str(text),
                                     align=al if al in ("left", "center", "centre", "right") else "left",
                                     bold=bold))
            tbl.append(tr)
        self._add(tbl)

    def schedule(self, block):
        unit = block.get("unit", "week")
        acts = block["activities"]
        first = 0 if unit == "week" else 1
        last = max(a["end"] for a in acts)
        periods = list(range(first, last + 1))
        n = len(periods)
        w_num = 420
        w_period = max(250, min(520, (self.page_width - w_num - 2600) // n))
        w_act = self.page_width - w_num - w_period * n
        widths = [w_num, w_act] + [w_period] * n
        tbl = self._table(widths)
        small = "16"
        if unit == "week":
            # header row 1: month groups (M-1 = W0..W4, then four weeks each)
            groups, i = [], 0
            while i < n:
                size = 5 if i == 0 else 4
                size = min(size, n - i)
                groups.append(size)
                i += size
            tr = self._row(self.tr_head, header=True)
            tr.append(self._cell(self.tc_head, w_num, "#", vmerge="restart"))
            tr.append(self._cell(self.tc_head, w_act, "Activity", vmerge="restart"))
            for gi, size in enumerate(groups):
                tr.append(self._cell(self.tc_head, w_period * size, f"M-{gi + 1}", span=size, tight=True, size=small))
            tbl.append(tr)
            tr = self._row(self.tr_head, header=True)
            tr.append(self._cell(self.tc_head, w_num, "", vmerge="continue"))
            tr.append(self._cell(self.tc_head, w_act, "", vmerge="continue"))
            for wk in periods:
                tr.append(self._cell(self.tc_head, w_period, f"W{wk}", tight=True, size=small))
            tbl.append(tr)
        else:
            tr = self._row(self.tr_head, header=True)
            tr.append(self._cell(self.tc_head, w_num, "#"))
            tr.append(self._cell(self.tc_head, w_act, "Activity"))
            for m in periods:
                tr.append(self._cell(self.tc_head, w_period, f"M{m}", tight=True, size=small))
            tbl.append(tr)
        for ai, a in enumerate(acts, 1):
            tr = self._row(self.tr_body)
            tr.append(self._cell(self.tc_body, w_num, str(ai), align="center"))
            tr.append(self._cell(self.tc_body, w_act, a["name"], align="left"))
            for pd in periods:
                on = a["start"] <= pd <= a["end"]
                tr.append(self._cell(self.tc_body, w_period, "O" if on else "", align="center",
                                     tight=True, bold=True, color=self.accent))
            tbl.append(tr)
        self._add(tbl)

    # ---- cover / header

    def set_titles(self, title):
        title = _clean(title).replace("\n", " ").strip()
        roots = []
        if self.cover is not None:
            roots.append(self.cover)
        for rel in self.doc.part.rels.values():
            if rel.reltype.endswith("/header") and not rel.is_external:
                roots.append(rel.target_part.element)
        for root in roots:
            for box in list(root.iter(qn("w:txbxContent"))):
                for p in box.findall(qn("w:p")):
                    if not _text_of(p).strip():
                        continue
                    runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
                    if not runs:
                        continue
                    t0 = runs[0].find(qn("w:t"))
                    t0.text = title
                    t0.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    for r in runs[1:]:
                        p.remove(r)
        self.doc.core_properties.title = "SOW – " + title


def _letter(i):
    s = ""
    i += 1
    while i:
        i, rem = divmod(i - 1, 26)
        s = chr(65 + rem) + s
    return s


def render_sow(sow: dict, out_path, template_path=None) -> Path:
    """Validate the SOW dict and write it as a .docx in the house format."""
    errs = validate_sow(sow)
    if errs:
        raise SOWValidationError("\n".join(errs))
    b = _Builder(template_path or DEFAULT_TEMPLATE)
    b.set_titles(sow["cover_title"])
    sections = sow["sections"]
    for i, sec in enumerate(sections):
        h = sec["heading"].strip()
        b.refs[h.rstrip(":").lower()] = f"Section {_letter(i)} ({h.rstrip(':')})"

    b.heading(sow.get("document_title") or "Scope of Work (SOW)")
    b.spacer()
    for i, sec in enumerate(sections):
        b.heading(f"{_letter(i)}. {sec['heading'].strip()}")
        b.spacer()
        sub_num = None
        for blk in sec["blocks"]:
            t = blk["type"]
            if t == "paragraph":
                b.paragraph(blk["text"])
            elif t == "note":
                b.note(blk["text"])
            elif t == "subheading":
                if sub_num is None:
                    sub_num = b._new_num(b.h4_abs)
                b.subheading(blk["text"], int(blk.get("level", 1)), sub_num)
            elif t == "bullets":
                b.list_block(blk["items"], numbered=False)
            elif t == "numbered":
                b.list_block(blk["items"], numbered=True, start=int(blk.get("start", 1)))
            elif t == "table":
                b.table(blk)
            elif t == "schedule":
                b.schedule(blk)
            b.spacer()
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    b.doc.save(str(out_path))
    return out_path
