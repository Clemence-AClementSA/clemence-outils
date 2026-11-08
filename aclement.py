import re, shutil, zipfile, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

VERT, BEIGE, BEIGE_CLAIR = "1D343D", "C9AA8A", "F0EDE5"
BULLET_NUMID = 5  # puces rondes du modèle Notes internes


def word_from_template(template_path):
    doc = Document(template_path)
    body = doc.element.body
    for el in list(body):
        if el.tag != qn("w:sectPr"):
            body.remove(el)
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer):
            for t in part._element.iter(qn("w:t")):
                if t.text and "ase " in t.text and "50" in t.text:
                    t.text = t.text.replace("Case Postale 50", "Rue de Vevey 178").replace("Case postale 50", "Rue de Vevey 178")
    return doc


def word_swap_logo(docx_path, base_word_path):
    logo = zipfile.ZipFile(base_word_path).read("word/media/image1.png")
    tmp = docx_path + ".tmp"; shutil.move(docx_path, tmp)
    zin = zipfile.ZipFile(tmp); zout = zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED)
    imgs = set()
    for n in zin.namelist():
        if re.match(r"word/_rels/(header|footer)\d+\.xml\.rels", n):
            imgs |= {"word/" + t for t in re.findall(r'Target="(media/[^"]+)"', zin.read(n).decode())}
    for n in zin.namelist():
        zout.writestr(n, logo if n in imgs else zin.read(n))
    zout.close(); zin.close(); os.remove(tmp)


def _font(r, font="Roboto", size=11, bold=None, color=None, italic=None):
    r.font.name = font; r.font.size = Pt(size)
    rpr = r._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), font)
    if bold is not None: r.bold = bold
    if italic is not None: r.italic = italic
    if color: r.font.color.rgb = RGBColor.from_string(color)


def rich(p, text, size=11, color=None, base_bold=False):
    """Texte avec **gras** inline."""
    text = text.replace("« ", "«\u00a0").replace(" »", "\u00a0»").replace(" :", "\u00a0:")
    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if not part: continue
        b = part.startswith("**") and part.endswith("**")
        r = p.add_run(part[2:-2] if b else part)
        _font(r, size=size, bold=(b or base_bold), color=color)
    return p


def para(doc, text="", size=11, color=None, bold=False, after=4, before=0, align=None, keep_next=False):
    p = doc.add_paragraph()
    if text: rich(p, text, size=size, color=color, base_bold=bold)
    pf = p.paragraph_format; pf.space_after = Pt(after); pf.space_before = Pt(before)
    if align is not None: p.alignment = align
    if keep_next: pf.keep_with_next = True
    return p


def blank(doc):
    return para(doc, "", after=0)


def title(doc, text, subtitle=None):
    para(doc, text, size=14, bold=True, color=VERT, after=2)
    if subtitle: para(doc, subtitle, size=11, color=BEIGE, after=8)


def _border(p, side="bottom", color=VERT, sz=6):
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    e = OxmlElement(f"w:{side}")
    e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:space"), "1"); e.set(qn("w:color"), color)
    bdr.append(e); pPr.append(bdr)


def section(doc, text):
    p = para(doc, text.upper(), size=11, bold=True, color=VERT, before=12, after=6, keep_next=True)
    _border(p)
    return p


def subsection(doc, text):
    return para(doc, text, size=11, bold=True, color=BEIGE, before=6, after=3, keep_next=True)


def bullet(doc, text, level=0, size=11):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    il = OxmlElement("w:ilvl"); il.set(qn("w:val"), str(level))
    ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(BULLET_NUMID))
    numPr.append(il); numPr.append(ni); pPr.insert(0, numPr)
    p.paragraph_format.space_after = Pt(2)
    rich(p, text, size=size)
    return p


def step(doc, n, text):
    """Étape numérotée : pastille n° beige + texte."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.9); p.paragraph_format.first_line_indent = Cm(-0.9)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(0.9))
    r = p.add_run(f"{n}.\t"); _font(r, bold=True, color=BEIGE)
    rich(p, text)
    return p


def _shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _cell_borders(table, color="BFBFBF", none=False):
    tbl = table._tbl; tblPr = tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "nil" if none else "single"); e.set(qn("w:sz"), "4"); e.set(qn("w:space"), "0"); e.set(qn("w:color"), color)
        b.append(e)
    tblPr.append(b)


def _fix_widths(t, widths_cm):
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    grid = t._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths_cm[i] * 567)))
    for row in t.rows:
        for i, c in enumerate(row.cells):
            c.width = Cm(widths_cm[i])


def _cell_text(cell, text, size=10, bold=False, color=None):
    cell.text = ""
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        rich(p, line, size=size, color=color, base_bold=bold)


def table(doc, headers, rows, widths_cm, size=9.5, header_fill=VERT, zebra=True, first_col_bold=False):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    _cell_borders(t)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.width = Cm(widths_cm[i])
        _shade(c, header_fill)
        _cell_text(c, h, size=size, bold=True, color="FFFFFF" if header_fill == VERT else VERT)
    # en-tête répétée
    trPr = t.rows[0]._tr.get_or_add_trPr(); th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].width = Cm(widths_cm[i])
            _cell_text(cells[i], str(v), size=size, bold=(first_col_bold and i == 0))
            if zebra and ri % 2 == 1: _shade(cells[i], BEIGE_CLAIR)
    _fix_widths(t, widths_cm)
    for row in t.rows:
        tr = row._tr; trPr = tr.get_or_add_trPr(); cs = OxmlElement("w:cantSplit"); trPr.append(cs)
        for c in row.cells:
            tcPr = c._tc.get_or_add_tcPr()
            mar = OxmlElement("w:tcMar")
            for side in ("top", "bottom"):
                e = OxmlElement(f"w:{side}"); e.set(qn("w:w"), "40"); e.set(qn("w:type"), "dxa"); mar.append(e)
            tcPr.append(mar)
    para(doc, "", after=2)
    return t


def info_block(doc, pairs, widths=(3.0, 5.5, 2.5, 6.0)):
    """pairs: liste de (libellé, valeur) ; 2 paires par ligne, sans bordures, Roboto 10."""
    rows = (len(pairs) + 1) // 2
    t = doc.add_table(rows=rows, cols=4); t.autofit = False
    _cell_borders(t, none=True)
    for i, (k, v) in enumerate(pairs):
        r, c = divmod(i, 2)
        a, b = t.rows[r].cells[c * 2], t.rows[r].cells[c * 2 + 1]
        a.width = Cm(widths[c * 2]); b.width = Cm(widths[c * 2 + 1])
        _cell_text(a, k, size=10, bold=True, color=VERT); _cell_text(b, v, size=10)
    _fix_widths(t, widths)
    para(doc, "", after=4)
    return t


def callout(doc, text, fill=BEIGE_CLAIR, size=10):
    t = doc.add_table(rows=1, cols=1); t.autofit = False
    _cell_borders(t, color=BEIGE)
    c = t.rows[0].cells[0]; c.width = Cm(17.0); _shade(c, fill)
    _cell_text(c, text, size=size, color=VERT)
    _fix_widths(t, [17.0])
    para(doc, "", after=4)
    return t


def page_break(doc):
    from docx.enum.text import WD_BREAK
    p = doc.add_paragraph(); p.add_run().add_break(WD_BREAK.PAGE)
