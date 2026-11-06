import io
import re
import base64
from flask import Flask, request, jsonify
from docx import Document
from docx.shared import Inches, Pt
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup, NavigableString

app = Flask(__name__)


# ================= markers =================
HL_OPEN_RE = re.compile(r"<<HL:([^>]*)>>")
HL_CLOSE = "<</HL>>"
HL_SPLIT_RE = re.compile(r"(<<HL:[^>]*>>|<</HL>>)")
HL_STRIP_RE = re.compile(r"<<HL:[^>]*>>|<</HL>>")

WD_HIGHLIGHT_HEX = {
    "YELLOW": "FFFF00", "BRIGHT_GREEN": "00FF00", "GREEN": "00FF00",
    "CYAN": "00FFFF", "TURQUOISE": "00FFFF", "MAGENTA": "FF00FF",
    "PINK": "FF00FF", "VIOLET": "FF00FF", "BLUE": "0000FF", "RED": "FF0000",
    "DARK_BLUE": "000080", "TEAL": "008080", "DARK_GREEN": "008000",
    "DARK_RED": "800000", "DARK_YELLOW": "808000", "GRAY_25": "C0C0C0",
    "GRAY_50": "808080", "BLACK": "000000",
}
CSS_NAMED_TO_HEX = {
    "yellow": "FFFF00", "lime": "00FF00", "aqua": "00FFFF", "cyan": "00FFFF",
    "magenta": "FF00FF", "fuchsia": "FF00FF", "blue": "0000FF", "red": "FF0000",
    "navy": "000080", "teal": "008080", "green": "008000", "olive": "808000",
    "purple": "800080", "maroon": "800000", "silver": "C0C0C0", "gray": "808080",
    "grey": "808080", "black": "000000", "white": "FFFFFF", "orange": "FFA500",
    "pink": "FFC0CB", "lightblue": "ADD8E6", "lightgreen": "90EE90",
}


def clean_line(s: str) -> str:
    return " ".join(s.replace("\t", " ").split()).strip()


def strip_hl(s: str) -> str:
    return HL_STRIP_RE.sub("", s)


def hex_to_wd_name(hex_no_hash: str):
    h = hex_no_hash.upper()
    for name, v in WD_HIGHLIGHT_HEX.items():
        if v == h and hasattr(WD_COLOR_INDEX, name):
            return name
    return None


def _css_color_to_marker_value(css: str):
    if not css:
        return None
    s = css.strip().lower()
    if s in ("transparent", "inherit", "initial", "none", "auto"):
        return None
    m = re.match(r"rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", s)
    if m:
        r, g, b = (int(x) for x in m.groups())
        h = f"{r:02X}{g:02X}{b:02X}"
        return hex_to_wd_name(h) or f"#{h}"
    if s.startswith("#"):
        h = s[1:].upper()
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        if re.fullmatch(r"[0-9A-F]{6}", h):
            return hex_to_wd_name(h) or f"#{h}"
    if s in CSS_NAMED_TO_HEX:
        h = CSS_NAMED_TO_HEX[s]
        return hex_to_wd_name(h) or f"#{h}"
    return None


# ================= font / highlight / spacing =================
def _apply_run_font(run, name: str, size: float):
    run.font.name = name
    run.font.size = Pt(size)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)


def _set_default_font(doc, name: str, size: float):
    style = doc.styles["Normal"]
    style.font.name = name
    style.font.size = Pt(size)
    rpr = style.element.get_or_add_rPr()
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)


def _set_spacing_element(pPr, before_pt, after_pt, line_mult):
    for old in pPr.findall(qn("w:spacing")):
        pPr.remove(old)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), str(int(round(before_pt * 20))))
    sp.set(qn("w:after"),  str(int(round(after_pt * 20))))
    sp.set(qn("w:line"),   str(int(round(line_mult * 240))))
    sp.set(qn("w:lineRule"), "auto")
    pPr.append(sp)


def _set_paragraph_spacing(paragraph, before_pt, after_pt, line_mult):
    pPr = paragraph._element.get_or_add_pPr()
    _set_spacing_element(pPr, before_pt, after_pt, line_mult)


def _set_normal_style_spacing(doc, before_pt, after_pt, line_mult):
    style = doc.styles["Normal"]
    pPr = style.element.get_or_add_pPr()
    _set_spacing_element(pPr, before_pt, after_pt, line_mult)


def _set_run_highlight(run, value: str):
    if not value.startswith("#"):
        enum_val = getattr(WD_COLOR_INDEX, value.upper(), None)
        if enum_val is not None:
            try:
                run.font.highlight_color = enum_val
                return
            except Exception:
                pass
    hex_val = value.lstrip("#").upper()
    if re.fullmatch(r"[0-9A-F]{6}", hex_val):
        rPr = run._element.get_or_add_rPr()
        for old in rPr.findall(qn("w:shd")):
            rPr.remove(old)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_val)
        rPr.append(shd)


def apply_highlights(paragraph, marked_text, font_name, font_size,
                     space_before=0.0, space_after=0.0, line_mult=1.0):
    _set_paragraph_spacing(paragraph, space_before, space_after, line_mult)
    current = None
    for chunk in HL_SPLIT_RE.split(marked_text):
        if not chunk:
            continue
        m = HL_OPEN_RE.fullmatch(chunk)
        if m:
            current = m.group(1)
        elif chunk == HL_CLOSE:
            current = None
        else:
            run = paragraph.add_run(chunk)
            _apply_run_font(run, font_name, font_size)
            if current:
                _set_run_highlight(run, current)


# ================= .docx highlight detection =================
def run_highlight_value(run):
    try:
        hc = run.font.highlight_color
        if hc is not None:
            name = getattr(hc, "name", None)
            if name and name not in ("AUTO", "INHERITED", "NONE"):
                return name
    except Exception:
        pass
    rPr = run._element.rPr
    if rPr is not None:
        shd = rPr.find(qn("w:shd"))
        if shd is not None:
            fill = (shd.get(qn("w:fill")) or "").upper()
            if fill and fill not in ("AUTO", "FFFFFF", "NONE"):
                wd = hex_to_wd_name(fill)
                return wd if wd else f"#{fill}"
    return None


def paragraph_marked(paragraph) -> str:
    out, current = [], None
    for run in paragraph.runs:
        if not run.text:
            continue
        val = run_highlight_value(run)
        if val != current:
            if current is not None:
                out.append(HL_CLOSE)
            if val is not None:
                out.append(f"<<HL:{val}>>")
            current = val
        out.append(run.text)
    if current is not None:
        out.append(HL_CLOSE)
    return "".join(out)


def cell_marked(cell) -> str:
    return " ".join(
        t for t in (clean_line(paragraph_marked(p)) for p in cell.paragraphs) if t
    )


# ================= extraction =================
def extract_docx(file_bytes: bytes):
    doc = Document(io.BytesIO(file_bytes))
    blocks, pending = [], []

    def flush():
        nonlocal pending
        if pending:
            blocks.append(("text", pending))
            pending = []

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            t = clean_line(paragraph_marked(Paragraph(child, doc)))
            if t:
                pending.append(t)
        elif child.tag == qn("w:tbl"):
            flush()
            rows = []
            for row in Table(child, doc).rows:
                cells = [cell_marked(c) for c in row.cells]
                if any(cells):
                    rows.append(cells)
            if rows:
                blocks.append(("table", rows))
    flush()
    return blocks


TEXT_TAGS = ["p", "li", "h1", "h2", "h3", "h4", "h5", "h6"]


def _collect(node, color, sink):
    if isinstance(node, NavigableString):
        txt = str(node)
        if txt:
            sink.append((txt, color))
        return
    new_color = color
    if getattr(node, "name", None) == "mark":
        new_color = "YELLOW"
    style = node.get("style") or ""
    m = re.search(r"background(?:-color)?\s*:\s*([^;]+)", style, re.I)
    if m:
        v = _css_color_to_marker_value(m.group(1))
        if v:
            new_color = v
    for child in node.children:
        _collect(child, new_color, sink)


def _segments_to_marked(segments):
    out, current = [], None
    for txt, val in segments:
        if val != current:
            if current is not None:
                out.append(HL_CLOSE)
            if val is not None:
                out.append(f"<<HL:{val}>>")
            current = val
        out.append(txt)
    if current is not None:
        out.append(HL_CLOSE)
    return "".join(out)


def extract_html(html: str):
    soup = BeautifulSoup(html, "html.parser")
    blocks, pending = [], []

    def flush():
        nonlocal pending
        if pending:
            blocks.append(("text", pending))
            pending = []

    for node in soup.find_all(TEXT_TAGS + ["table"]):
        if node.name == "table":
            if node.find_parent("table"):
                continue
            flush()
            rows = []
            for tr in node.find_all("tr"):
                cells = []
                for cell in tr.find_all(["td", "th"]):
                    segs = []
                    _collect(cell, None, segs)
                    cells.append(clean_line(_segments_to_marked(segs)))
                if any(cells):
                    rows.append(cells)
            if rows:
                blocks.append(("table", rows))
        else:
            if node.find_parent("table"):
                continue
            segs = []
            _collect(node, None, segs)
            t = clean_line(_segments_to_marked(segs))
            if t:
                pending.append(t)
    flush()
    return blocks


# ================= page setup =================
def _set_margins(section, top, bottom, left, right, gutter):
    section.top_margin = Inches(top)
    section.bottom_margin = Inches(bottom)
    section.left_margin = Inches(left)
    section.right_margin = Inches(right)
    sectPr = section._sectPr
    pgMar = sectPr.find(qn("w:pgMar"))
    if pgMar is None:
        pgMar = OxmlElement("w:pgMar")
        sectPr.append(pgMar)
    pgMar.set(qn("w:gutter"), str(int(gutter * 1440)))


def _set_gutter_position(doc, position: str):
    settings = doc.settings.element
    existing = settings.find(qn("w:gutterAtTop"))
    if position == "Top":
        if existing is None:
            settings.append(OxmlElement("w:gutterAtTop"))
    else:
        if existing is not None:
            settings.remove(existing)


def _set_columns(section, num_cols, spacing_inches, line_between):
    sectPr = section._sectPr
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sectPr.append(cols)
    cols.set(qn("w:num"), str(num_cols))
    cols.set(qn("w:equalWidth"), "1")
    cols.set(qn("w:space"), str(int(spacing_inches * 1440)))
    cols.set(qn("w:sep"), "1" if line_between else "0")


# ================= table styling =================
def _apply_table_style(table, style_name: str):
    try:
        table.style = style_name
    except Exception:
        try:
            table.style = "Table Grid"
        except Exception:
            pass


def _apply_table_autofit(table, mode: str):
    tblPr = table._element.tblPr
    for tag in ("w:tblW", "w:tblLayout"):
        for old in tblPr.findall(qn(tag)):
            tblPr.remove(old)

    if mode == "window":
        table.autofit = True
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:type"), "pct")
        tblW.set(qn("w:w"), "5000")
        tblPr.append(tblW)
    elif mode == "fixed":
        table.autofit = False
    else:
        table.autofit = True
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:type"), "auto")
        tblW.set(qn("w:w"), "0")
        tblPr.append(tblW)

    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed" if mode == "fixed" else "autofit")
    tblPr.append(tblLayout)


# ================= builders =================
def build_docx(blocks, sep: str, cfg: dict) -> bytes:
    doc = Document()
    _set_default_font(doc, cfg["font_name"], cfg["font_size"])
    _set_normal_style_spacing(
        doc, cfg["space_before"], cfg["space_after"], cfg["line_spacing"]
    )

    section = doc.sections[0]
    _set_margins(
        section,
        cfg["margin_top"], cfg["margin_bottom"],
        cfg["margin_left"], cfg["margin_right"],
        cfg["gutter"],
    )
    _set_gutter_position(doc, cfg["gutter_position"])
    _set_columns(section, cfg["num_cols"], cfg["col_spacing"], cfg["col_line"])

    def _para_factory(paragraph, marked_text):
        apply_highlights(
            paragraph, marked_text,
            cfg["font_name"], cfg["font_size"],
            cfg["space_before"], cfg["space_after"], cfg["line_spacing"],
        )

    for kind, data in blocks:
        if kind == "text":
            merged = sep.join(data)
            if merged:
                _para_factory(doc.add_paragraph(), merged)
        else:
            rows = data
            ncols = max(len(r) for r in rows)
            t = doc.add_table(rows=len(rows), cols=ncols)
            _apply_table_style(t, cfg["table_style"])
            _apply_table_autofit(t, cfg["autofit_mode"])
            for i, row in enumerate(rows):
                for j in range(ncols):
                    cell = t.cell(i, j)
                    cell.text = ""
                    _para_factory(cell.paragraphs[0], row[j] if j < len(row) else "")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_txt(blocks, sep: str) -> str:
    out, n = [], 0
    for kind, data in blocks:
        if kind == "text":
            out.append(strip_hl(sep.join(data)))
        else:
            n += 1
            out.append("")
            out.append(f"--- Table {n} ---")
            for row in data:
                out.append("\t".join(strip_hl(c) for c in row))
            out.append("")
    return "\n".join(out).strip()


# ================= API =================
@app.route("/", methods=["GET", "POST"])
def handle():
    if request.method == "GET":
        return jsonify({"status": "ok", "usage": "POST with 'file' or 'html' + settings"})

    def _f(name, default, cast=float):
        v = request.form.get(name)
        return cast(v) if v not in (None, "") else default

    cfg = {
        "font_name": request.form.get("font_name", "Aptos"),
        "font_size": _f("font_size", 5.0),
        "table_style": request.form.get("table_style", "Plain Table 1"),
        "autofit_mode": request.form.get("autofit_mode", "contents"),
        "space_before": _f("space_before", 0.0),
        "space_after": _f("space_after", 0.0),
        "line_spacing": _f("line_spacing", 1.0),
        "margin_top": _f("margin_top", 0.05),
        "margin_bottom": _f("margin_bottom", 0.05),
        "margin_left": _f("margin_left", 0.05),
        "margin_right": _f("margin_right", 0.05),
        "gutter": _f("gutter", 0.0),
        "gutter_position": request.form.get("gutter_position", "Left"),
        "num_cols": _f("num_cols", 3, int),
        "col_spacing": _f("col_spacing", 0.1),
        "col_line": request.form.get("col_line", "false").lower() == "true",
    }
    sep = request.form.get("sep", ", ")

    try:
        if "file" in request.files and request.files["file"].filename:
            blocks = extract_docx(request.files["file"].read())
        elif "html" in request.form and request.form["html"].strip():
            blocks = extract_html(request.form["html"])
        elif "text" in request.form and request.form["text"].strip():
            paras = [t for t in (clean_line(l) for l in request.form["text"].split("\n")) if t]
            blocks = [("text", paras)] if paras else []
        else:
            return jsonify({"error": "No input provided"}), 400
    except Exception as e:
        return jsonify({"error": f"Parse error: {e}"}), 500

    if not blocks:
        return jsonify({"error": "No content found in input"}), 400

    try:
        docx_bytes = build_docx(blocks, sep, cfg)
        txt = build_txt(blocks, sep)
    except Exception as e:
        return jsonify({"error": f"Build error: {e}"}), 500

    return jsonify({
        "docx_b64": base64.b64encode(docx_bytes).decode("ascii"),
        "txt": txt,
        "blocks": [{"kind": k, "data": d} for k, d in blocks],
        "sep": sep,
    })