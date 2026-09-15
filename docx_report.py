"""docx_report.py — renders the same report context as a Word (.docx) file,
alongside (not instead of) the existing PDF from generate_report.py's
render_pdf(). Both consume the exact same `context` dict from
build_report_context(), so there is exactly one source of truth for the
numbers — this module only re-lays them out for Word.

Word has no flexbox/CSS, so layout is reconstructed with tables (for
side-by-side columns and shaded panels) rather than mirroring template.html's
markup directly. This is a deliberate structural match against template.html,
not a loose approximation:
- "Executive Snapshot" and "Performance Overview" sit in a narrow label
  column to the LEFT of their tables (matching .row-label/.row-content in the
  HTML — a sidebar layout, not a heading placed above the table)
- teal (#105652) for section labels, positive figures, and panel backgrounds
- red (#ff3131) for negative figures
- bordered tables for the Executive Snapshot / Performance Overview /
  Matched Listings, matching the PDF's table.plain / table.listings
- the teal Key Insights box as a shaded borderless table
- the Helpful Tip as a genuine circular shape (DrawingML ellipse with a
  text box inside), matching the PDF's border-radius: 50% .tip-circle —
  python-docx has no shape-drawing API, so this is hand-built OOXML parsed
  with lxml, not something python-docx's own OxmlElement() helper can
  construct (it doesn't know the DrawingML/WordprocessingShape namespaces).

Verified end-to-end against actual Word output, not just by reading the XML
back: Word is available via COM automation on the build machine, so every
change here has been round-tripped through a real docx -> Word -> exported
PDF cycle and visually compared against the reference PDF, the same way a
recipient would actually experience opening this file. The circle shape
specifically needed this — an early version used <a:noAutofit/> in the text
box, which looked fine with a short test string but let the real (much
longer) tip text run on as one line past the circle's edge entirely, only
caught by testing with the actual production text rather than a placeholder.
"""
import os
from xml.sax.saxutils import escape as _xml_escape

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Mm, Pt, RGBColor
from lxml import etree

TEAL = RGBColor(0x10, 0x56, 0x52)
RED = RGBColor(0xFF, 0x31, 0x31)
BLACK = RGBColor(0x00, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def _shade_cell(cell, hex_color: str):
    """Set a table cell's background fill. python-docx has no high-level API
    for this — it means dropping to the underlying XML directly."""
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def _no_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        borders.append(el)
    tbl_pr.append(borders)


def _set_full_borders(table, color_hex: str, sz: int):
    """Sets a solid box-and-grid border on every edge of a table (outer box
    plus the lines between cells) — matches table.listings' CSS border,
    which is on the table AND every cell. Explicit, rather than relying on
    Word's built-in "Table Grid" style: that style's own default border
    weight doesn't track whatever width we actually want, which is exactly
    why the Matched Listings table was rendering visibly thinner than the
    reference design's bolder grid."""
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), color_hex)
        borders.append(el)
    tbl_pr.append(borders)


def _clear_empty_leading_paragraph(cell):
    """A table cell always starts with exactly one empty paragraph (OOXML
    requires at least one) — this only matters when content is added as a
    nested table/object rather than through _para(), which already reuses
    that paragraph itself. Call this after adding a nested table to a cell,
    or the empty paragraph is left sitting above it as a stray blank line."""
    if cell.paragraphs and not cell.paragraphs[0].runs:
        cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)


def _set_run(run, size=10, bold=False, italic=False, color=BLACK, upper=False, font_name="Arial"):
    if upper:
        run.text = run.text.upper()
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    # Set directly on every run rather than relying on the Normal style
    # cascading down — a run with its own rPr (which every one of these has,
    # since size/bold/color are all direct formatting) should inherit the
    # style's font in Word's own rendering model, but setting it explicitly
    # here removes any dependency on that cascade actually happening the way
    # the spec says it should.
    # font_name defaults to Arial (the document-wide baseline) -- only the
    # Executive Snapshot/Performance Overview tables, the header-meta block,
    # and the office-name title pass a different one (Inter / Bebas Neue),
    # per an explicit, narrowly-scoped request. Word substitutes a fallback
    # if the actual font isn't installed on the machine opening the file,
    # same as any other missing-font case -- it won't error either way.
    run.font.name = font_name
    rFonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), font_name)
    return run


def _set_cell_margins(cell, top_mm=None, bottom_mm=None, left_mm=None, right_mm=None):
    """Sets cell padding (OOXML tcMar) in millimeters, matching template.html's
    CSS padding on table.plain/table.listings cells and the panel elements.
    python-docx has no high-level API for this either — table cells default
    to Word's own built-in margins, which don't match the CSS values at all."""
    mar = OxmlElement("w:tcMar")
    for tag, val_mm in (("top", top_mm), ("bottom", bottom_mm), ("left", left_mm), ("right", right_mm)):
        if val_mm is None:
            continue
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:w"), str(int(Mm(val_mm).twips)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    cell._tc.get_or_add_tcPr().append(mar)


def _para(doc_or_cell, text="", size=10, bold=False, italic=False, color=BLACK,
          align=None, upper=False, style=None, font_name="Arial"):
    """Add a styled paragraph to a Document or a table cell.

    A fresh table cell always has exactly one, empty paragraph already
    (OOXML requires it) — reusing it for the first piece of content, rather
    than always calling add_paragraph() and leaving that original one as a
    stray blank line, is correct there. But this reuse must be scoped to
    cells specifically (detected via the _tc attribute only cells have):
    applying the same trick to the top-level Document caused a real bug —
    once one _para(doc, ...) call left its paragraph empty, every later
    _para(doc, ...) call kept re-targeting that same original paragraph
    instead of appending a new one at the current end, silently detaching
    every table and paragraph added afterward from its intended position in
    the document (they'd still get created, just all shifted to the wrong
    place). A Document's paragraphs must always be freshly appended.
    """
    is_cell = hasattr(doc_or_cell, "_tc")
    existing = doc_or_cell.paragraphs if is_cell else None
    if is_cell and existing and len(existing) == 1 and not existing[0].runs and not style:
        p = existing[0]
    else:
        p = doc_or_cell.add_paragraph(style=style) if style else doc_or_cell.add_paragraph()
    if align is not None:
        p.alignment = align
    if text:
        _set_run(p.add_run(text), size=size, bold=bold, italic=italic, color=color, upper=upper,
                 font_name=font_name)
    return p


def _spacer(doc, pt=3):
    """A small gap between sections, matching template.html's tight
    margins (4-6mm) — NOT achieved via a run's font size, because an empty
    paragraph (no text, so no run at all) doesn't hold one: its height comes
    from the paragraph's own default line spacing, and with the document-wide
    1.4 line-spacing rule that made every "spacer" identically tall regardless
    of what size was passed in. Setting exact line spacing directly on this
    one paragraph is what actually controls its height.
    """
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(pt)
    return p


def _set_col_widths(table, widths_cm):
    """Sets explicit column widths AND forces fixed table layout.

    Word's default table layout ("autofit to contents") recalculates column
    widths from what's actually in each cell, silently overriding a width set
    on individual cells — this is exactly why a short label like "EXECUTIVE
    SNAPSHOT" was not staying confined to its narrow column: autofit was
    resizing it based on the label's own text length instead of respecting
    the column width we asked for. Fixed layout, plus rewriting the table's
    grid column widths (not just each cell's), is what makes Word actually
    honor them.

    Also pins the table's own overall preferred width (w:tblW) and zeroes
    its indent (w:tblInd) explicitly, rather than leaving Word to infer the
    table's width/position from its grid columns. Leaving those implicit was
    fine in Word itself, but Google Docs' docx renderer positioned at least
    one borderless table (the Key Insights panel) visibly further left than
    the header/other tables above it despite identical column-width math —
    an explicit width/indent removes that ambiguity for any renderer, Word
    included.

    w:tblW is set to the SUM OF THE ALREADY-ROUNDED per-column twips values,
    not a separately-rounded total (e.g. int(Cm(5.93).twips) three times
    doesn't necessarily sum to int(Cm(17.8).twips) -- each rounds/truncates
    independently, and for these three widths specifically it's off by a
    single twip). A twip is far too small to see on its own, but it means
    the table's own declared total width didn't exactly match the sum of
    its columns -- an internal inconsistency a renderer has to silently
    resolve somehow, and manually dragging the table's edge in Word/Docs
    (reported as "the left and right margins don't match" after doing so)
    is exactly the kind of interaction that can expose which side it
    resolves it on.
    """
    col_twips = [int(Cm(w).twips) for w in widths_cm]
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr.find(qn("w:tblLayout")) is None:
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tbl_pr.append(layout)

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(sum(col_twips)))

    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:type"), "dxa")
    tbl_ind.set(qn("w:w"), "0")

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for grid_col, w_twips in zip(grid.findall(qn("w:gridCol")), col_twips):
            grid_col.set(qn("w:w"), str(w_twips))

    # Setting each cell's own tcW from the same col_twips values (rather than
    # cell.width = Cm(w), which re-derives dxa from EMU independently) keeps
    # tblW/gridCol/tcW all traceable to one single rounded value per column,
    # instead of three separate roundings that could each land a twip apart.
    for row in table.rows:
        for cell, w_twips in zip(row.cells, col_twips):
            tcW = cell._tc.get_or_add_tcPr().find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                cell._tc.get_or_add_tcPr().append(tcW)
            tcW.set(qn("w:type"), "dxa")
            tcW.set(qn("w:w"), str(w_twips))


def _set_page_background(doc, hex_color: str):
    """Word's equivalent of template.html's .page { background: #efeee7 } --
    needs BOTH a <w:background> element on the document root AND
    <w:displayBackgroundShape/> in settings.xml; either alone is silently
    ignored. Confirmed present via an actual Word -> PDF export round-trip,
    not just by reading the XML back."""
    background = OxmlElement("w:background")
    background.set(qn("w:color"), hex_color)
    doc.element.insert(0, background)

    settings_el = doc.settings.element
    settings_el.insert(0, OxmlElement("w:displayBackgroundShape"))


def _set_cell_bottom_border(cell, color_hex: str, sz: int):
    """Sets only a bottom border on a cell — matches table.plain's CSS
    (border-bottom on header/body cells only, no vertical/top/left/right
    lines) rather than a full box grid like table.listings genuinely has."""
    borders = OxmlElement("w:tcBorders")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:color"), color_hex)
    borders.append(bottom)
    cell._tc.get_or_add_tcPr().append(borders)


def _add_circle_shape(paragraph, diameter_mm, fill_hex, title, body_text,
                       title_size_pt=7.5, body_size_pt=6.5):
    """Inserts a true circular shape (DrawingML ellipse) with centered text
    into a paragraph, matching template.html's .tip-circle (border-radius:
    50%). python-docx has no shape-drawing API at all, so this is hand-built
    OOXML parsed with lxml — the DrawingML/WordprocessingShape namespaces
    involved aren't ones python-docx's own OxmlElement() helper knows.

    <a:spAutoFit/> in bodyPr is the load-bearing setting here. The more
    "obvious"-looking <a:noAutofit/> also disables wrapping text to the
    shape's own width — the real (long) tip text just ran on as one line
    past the circle's edge with it, only caught by testing against the
    actual production text rather than a short placeholder. spAutoFit
    correctly wraps and centers the full text inside the fixed-size circle.
    """
    emu = int(Mm(diameter_mm).emu)
    title_sz = int(title_size_pt * 2)  # w:sz is in half-points
    body_sz = int(body_size_pt * 2)
    xml = f"""<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="{emu}" cy="{emu}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="1" name="TipCircle"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:cNvSpPr/>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu}" cy="{emu}"/></a:xfrm>
              <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>
              <a:ln><a:noFill/></a:ln>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p>
                  <!-- .tip-circle .tip-title margin-bottom: 1.5mm -->
                  <w:pPr><w:jc w:val="center"/><w:spacing w:after="85"/></w:pPr>
                  <w:r>
                    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="{title_sz}"/></w:rPr>
                    <w:t>{_xml_escape(title)}</w:t>
                  </w:r>
                </w:p>
                <w:p>
                  <w:pPr><w:jc w:val="center"/></w:pPr>
                  <w:r>
                    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:color w:val="FFFFFF"/><w:sz w:val="{body_sz}"/></w:rPr>
                    <w:t>{_xml_escape(body_text)}</w:t>
                  </w:r>
                </w:p>
              </w:txbxContent>
            </wps:txbx>
            <!-- .tip-circle padding: 3mm, i.e. 108000 EMU per side -->
            <wps:bodyPr wrap="square" lIns="108000" tIns="108000" rIns="108000" bIns="108000" anchor="ctr">
              <a:spAutoFit/>
            </wps:bodyPr>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>"""
    r_element = etree.fromstring(xml.encode("utf-8"))
    paragraph._p.append(r_element)


def _labeled_row(doc, label_text: str):
    """Matches template.html's .labeled-row: a narrow label column (.row-label,
    30mm/~3cm in the CSS) beside a wider content column (.row-content), with
    a 6mm gap between them (CSS `gap: 6mm` on the flex row) — a sidebar
    layout, not a heading placed above the content. The middle column here is
    that gap, given a table's own columns otherwise sit flush against each
    other with nothing between them. Returns the content cell (now the third
    column) for the caller to build into."""
    wrap = doc.add_table(rows=1, cols=3)
    _no_borders(wrap)
    _set_col_widths(wrap, [3.0, 0.6, 14.2])
    label_cell, content_cell = wrap.cell(0, 0), wrap.cell(0, 2)
    # .row-label { padding-top: 1.5mm } -- nudges the label down to align
    # with the table header baseline instead of its own top edge. This only
    # has an effect if the cell is actually top-aligned in the first place —
    # without an explicit vertical_alignment, at least one renderer seen in
    # testing centered the label vertically across the *entire* row's height
    # (which spans the whole exec/perf table below it, not just its header),
    # leaving the label looking like it floats independently of the header
    # line it's meant to sit beside.
    label_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    content_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _set_cell_margins(label_cell, top_mm=1.5)
    _para(label_cell, label_text, size=9, bold=True, color=TEAL, upper=True, font_name="Inter")
    return content_cell


def render_docx(context: dict, output_path: str):
    doc = Document()
    _set_page_background(doc, "EFEEE7")

    # template.html's font-family is Arial first — Word's own default (which
    # varies by version/locale, e.g. Calibri or Aptos) is never guaranteed to
    # match, so every run needs this set explicitly rather than relying on
    # whatever the installed Word's own default template happens to be.
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)  # html, body { font-size: 10.5pt } baseline
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    # Word's own default template usually adds ~8-10pt space after every
    # paragraph — template.html's margins are much tighter (4-6mm gaps
    # between sections), and that mismatch compounded across a whole page's
    # worth of paragraphs was pushing content onto an extra page. Spacer
    # paragraphs below now control gaps explicitly via font size instead.
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    # template.html's line-height: 1.4 on html/body -- Word's own default
    # single-spacing is tighter than that for most fonts.
    normal.paragraph_format.line_spacing = 1.4

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)

    # -- running footer, appears on every page, matching template.html's .footer --
    footer_table = section.footer.add_table(rows=1, cols=2, width=Cm(17.8))
    _no_borders(footer_table)
    _set_col_widths(footer_table, [12.0, 5.8])
    _para(footer_table.cell(0, 0), context["footer_contact"], size=7.5)
    _para(footer_table.cell(0, 1), context["logo_text"], size=10, bold=True,
          align=WD_ALIGN_PARAGRAPH.RIGHT)

    # ================= PAGE 1: SUMMARY =================

    # .header { gap: 6mm } -- a middle spacer column, same approach as
    # _labeled_row, since a table's columns otherwise sit flush together.
    header_table = doc.add_table(rows=1, cols=3)
    _no_borders(header_table)
    _set_col_widths(header_table, [10.4, 0.6, 6.8])
    left = header_table.cell(0, 0)
    # The document-wide 1.4 line-spacing (matching template.html's line-height:
    # 1.4 on body text) compounds very differently on a single 24pt line than
    # it does on 9-10pt text -- Word's own single-line-spacing metric for a
    # large bold font is already taller than Chromium's, so multiplying that
    # by 1.4 on top left a visibly larger gap around the title specifically
    # than the PDF shows. Single-spacing just this one paragraph (not the
    # whole document) removes that compounding without touching anything else.
    # The reference design sets the office-name title in Bebas Neue, an
    # already-heavy display font in its own right -- applying Word's bold on
    # top of it (as we did for the Arial fallback) over-thickens it, which is
    # why bold is explicitly off here specifically.
    office_name_p = _para(left, context["office_name"], size=24, bold=False, upper=True,
                           font_name="Bebas Neue")
    office_name_p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    _para(left, "Quiet List Exchange Activity Report", size=11, bold=True, upper=True, color=TEAL)
    right = header_table.cell(0, 2)
    for label, value in (
        (context["reporting_period_label"], context["office_name"]),
        ("REPORTING PERIOD", context["period_range_label"]),
        ("PREPARED FOR", context["prepared_for"]),
        ("DATE", context["report_date_label"]),
    ):
        p = _para(right, align=WD_ALIGN_PARAGRAPH.RIGHT)
        _set_run(p.add_run(f"{label}: "), size=9, color=TEAL, font_name="Inter")
        _set_run(p.add_run(value), size=9, bold=True, color=BLACK, font_name="Inter")

    # .header { margin-bottom: 5mm }
    _spacer(doc, pt=Mm(5).pt)
    exec_content = _labeled_row(doc, "Executive Snapshot")
    exec_table = exec_content.add_table(rows=1, cols=3)
    exec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = exec_table.rows[0].cells
    for i, label in enumerate(("Metric", "Result", "Change vs. Previous Period")):
        _para(hdr[i], label, size=7, bold=True, upper=True,
              align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER,
              font_name="Inter")
        _set_cell_bottom_border(hdr[i], "105652", 12)
        _set_cell_margins(hdr[i], top_mm=1.3, bottom_mm=1.3, left_mm=3, right_mm=3)
    for row in context["exec_snapshot"]:
        cells = exec_table.add_row().cells
        _para(cells[0], row["label"], size=7.5, bold=True, upper=True, font_name="Inter")
        _para(cells[1], row["result"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER, font_name="Inter")
        _para(cells[2], row["change"], size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              color=TEAL if row["positive"] else RED, font_name="Inter")
        for cell in cells:
            _set_cell_bottom_border(cell, "B9C4BF", 6)
            _set_cell_margins(cell, top_mm=1.6, bottom_mm=1.6, left_mm=3, right_mm=3)
    # Result only ever holds a short number/percentage -- giving Metric most
    # of the row's width (so its longer uppercase labels stay comfortably on
    # one line at a readable size) and Result very little is both a closer
    # match to how much room each column's actual content needs, and what
    # was asked for directly: a narrower Result column.
    _set_col_widths(exec_table, [6.5, 3.0, 4.7])
    _clear_empty_leading_paragraph(exec_content)

    # table.plain's own margin-bottom (4mm, inside row-content) plus
    # .labeled-row's margin-bottom (4mm, after the whole row) stack to 8mm
    # of real gap in the HTML before the next labeled row begins.
    _spacer(doc, pt=Mm(8).pt)
    perf_content = _labeled_row(doc, "Performance Overview")
    # .two-col { gap: 6mm } -- middle spacer column, same approach as above.
    perf_wrap = perf_content.add_table(rows=1, cols=3)
    _no_borders(perf_wrap)
    _set_col_widths(perf_wrap, [9.4, 0.6, 4.2])
    perf_cell, tip_cell = perf_wrap.cell(0, 0), perf_wrap.cell(0, 2)

    perf_table = perf_cell.add_table(rows=1, cols=3)
    hdr = perf_table.rows[0].cells
    for i, label in enumerate(("Property Type", "Listings", "Matches")):
        _para(hdr[i], label, size=7, bold=True, upper=True,
              align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER,
              font_name="Inter")
        _set_cell_bottom_border(hdr[i], "105652", 12)
        _set_cell_margins(hdr[i], top_mm=1.3, bottom_mm=1.3, left_mm=3, right_mm=3)
    for row in context["performance_overview"]:
        cells = perf_table.add_row().cells
        _para(cells[0], row["property_type"], size=7.5, bold=True, upper=True, font_name="Inter")
        _para(cells[1], row["listings"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER, font_name="Inter")
        _para(cells[2], row["matches"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER, font_name="Inter")
        for cell in cells:
            _set_cell_bottom_border(cell, "B9C4BF", 6)
            _set_cell_margins(cell, top_mm=1.6, bottom_mm=1.6, left_mm=3, right_mm=3)
    # Same reasoning as the Executive Snapshot table: Listings/Matches only
    # ever hold a short number, so Property Type gets most of the width.
    _set_col_widths(perf_table, [4.7, 2.0, 2.7])
    _clear_empty_leading_paragraph(perf_cell)

    tip_cell.vertical_alignment = 1  # center
    tip_para = _para(tip_cell, align=WD_ALIGN_PARAGRAPH.CENTER)
    # 38mm rather than the CSS's exact 32mm -- a deliberate, larger buffer
    # than before. The circle's text is a hand-built DrawingML shape+textbox
    # (see _add_circle_shape's docstring); real Word renders it correctly at
    # 34mm, but Google Docs' docx renderer does not reliably keep the text
    # box visually anchored inside the shape at all -- its text appeared
    # below the circle entirely, not just overflowing its edge. A bigger
    # shape with a smaller font (below) narrows that gap, though it may not
    # fully close it if Google Docs' shape support is the real limit rather
    # than available space.
    circle_diameter_mm = 38
    _add_circle_shape(tip_para, diameter_mm=circle_diameter_mm, fill_hex="105652",
                       title="Helpful Tip", body_text=context["helpful_tip"],
                       title_size_pt=7, body_size_pt=6)
    # Word's row-height auto-calculation doesn't reliably count an inline
    # drawing's height toward the row it sits in the same way it counts
    # ordinary text -- the PDF export path recalculates layout and looked
    # correct, but Word's own on-screen view left this row sized to
    # perf_cell's much shorter nested table, so the 34mm circle overflowed
    # upward into the row above rather than the row expanding to fit it.
    # Forcing an explicit minimum height removes that ambiguity outright.
    perf_wrap.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    perf_wrap.rows[0].height = Mm(circle_diameter_mm + 4)
    _clear_empty_leading_paragraph(perf_content)

    # .two-col's own margin-bottom (4mm) plus .labeled-row's (4mm) stack to
    # 8mm, same reasoning as the Executive Snapshot gap above.
    _spacer(doc, pt=Mm(8).pt)
    insights = doc.add_table(rows=1, cols=3)
    _no_borders(insights)
    # Full container width -- spans the same total width as the row-label +
    # content area above it (17.8cm), same as the header/tables. The actual
    # bug reported earlier wasn't the width, it was that this table rendered
    # with an unwanted left offset in Google Docs despite that; the
    # explicit width + zeroed indent in _set_col_widths is what actually
    # fixes that positioning, independent of how wide the table itself is.
    _set_col_widths(insights, [5.93, 5.93, 5.94])
    for c in range(3):
        _shade_cell(insights.cell(0, c), "105652")
    col1, col2, col3 = insights.cell(0, 0), insights.cell(0, 1), insights.cell(0, 2)
    # .insights-panel { padding: 4mm 7mm; gap: 8mm } -- the 7mm is the panel's
    # own outer edge, the 8mm gap between columns splits ~4mm to each side of
    # the inner edges; top/bottom (4mm) applies uniformly to all three.
    _set_cell_margins(col1, top_mm=4, bottom_mm=4, left_mm=7, right_mm=4)
    _set_cell_margins(col2, top_mm=4, bottom_mm=4, left_mm=4, right_mm=4)
    _set_cell_margins(col3, top_mm=4, bottom_mm=4, left_mm=4, right_mm=7)

    # .insights-title { text-align: center; margin-bottom: 2.5mm } -- applies
    # to all three columns' titles, including col1's real "Key Insights" (not
    # just the invisible placeholders in col2/col3 below), so it's centered
    # like the other two, not left-aligned.
    title_gap = Pt(Mm(2.5).pt)
    p = _para(col1, "Key Insights", size=9, bold=True, color=WHITE, upper=True,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_after = title_gap

    # .insights-col { font-size: 8pt } -- nudged down to 7.5pt here only,
    # a deliberate docx-specific compromise: the extra vertical gaps just
    # added above (title/block margins) are accurate to the CSS but, unlike
    # the PDF's own text, push Word's page count from 2 to 3. Shrinking this
    # one block's font a half-point recovers most of that without cutting
    # any of the actual commentary content below.
    INSIGHT_BODY_SIZE = 7.5

    def _insight_block(cell, label, *lines, align=None, gap_before_last=None):
        # .insight-block { margin-bottom: 3mm } -- one call per block; the
        # optional gap_before_last reproduces a single line's own
        # margin-top (e.g. the featured-listing caption) between two lines
        # that would otherwise sit back-to-back.
        _para(cell, label, size=INSIGHT_BODY_SIZE, bold=True, italic=True, color=WHITE, align=align)
        for i, line in enumerate(lines):
            lp = _para(cell, line, size=INSIGHT_BODY_SIZE, color=WHITE, align=align)
            if i > 0 and gap_before_last is not None:
                lp.paragraph_format.space_before = gap_before_last
        cell.paragraphs[-1].paragraph_format.space_after = Pt(Mm(3).pt)

    _insight_block(col1, "Featured Listing", context["featured_listing_address"],
                    context["featured_listing_caption"],
                    gap_before_last=Pt(Mm(1.5).pt))
    _insight_block(col1, "Top Performing Suburb", context["top_suburb_display"])

    # col2/col3 have no real title of their own -- template.html reserves an
    # invisible placeholder line (visibility:hidden) so all three columns'
    # actual content still starts at the same height. A run coloured to
    # match the panel's own teal background does the same job here: present
    # for layout, invisible against the fill.
    for c in (col2, col3):
        ph = _para(c, " ", size=9, bold=True, color=TEAL, align=WD_ALIGN_PARAGRAPH.CENTER)
        ph.paragraph_format.space_after = title_gap

    _insight_block(col2, "Most active budget range.", context["budget_range_display"],
                    align=WD_ALIGN_PARAGRAPH.CENTER)
    _insight_block(col2, "Most active dwelling type.", context["dwelling_type_display"],
                    align=WD_ALIGN_PARAGRAPH.CENTER)

    _para(col3, "Most Active Buyer's Agencies", size=INSIGHT_BODY_SIZE, bold=True, italic=True,
          color=WHITE, align=WD_ALIGN_PARAGRAPH.RIGHT)
    for i, name in enumerate(context["top_agencies"], 1):
        _para(col3, f"{i}. {name}", size=INSIGHT_BODY_SIZE, color=WHITE,
              align=WD_ALIGN_PARAGRAPH.RIGHT)
    col3.paragraphs[-1].paragraph_format.space_after = Pt(Mm(3).pt)
    _para(col3, "Operative", size=INSIGHT_BODY_SIZE, bold=True, italic=True, color=WHITE,
          align=WD_ALIGN_PARAGRAPH.RIGHT)
    for i, name in enumerate(context["top_operatives"], 1):
        _para(col3, f"{i}. {name}", size=INSIGHT_BODY_SIZE, color=WHITE,
              align=WD_ALIGN_PARAGRAPH.RIGHT)
    col3.paragraphs[-1].paragraph_format.space_after = Pt(Mm(3).pt)

    # .insights-panel { margin-bottom: 4mm }
    _spacer(doc, pt=Mm(4).pt)
    _para(doc, "Commentary:", size=9, bold=True, color=TEAL, upper=True)
    for bullet in context["commentary"]:
        # .commentary li { font-size: 8.5pt; margin-bottom: 2mm } -- nudged
        # down to 8pt here, the other half of the docx-specific compromise
        # described above: keeps the full commentary text intact rather than
        # cutting it, while still recovering the vertical room the CSS-
        # accurate section gaps above now take up.
        p = _para(doc, bullet, size=7.5, style="List Bullet")
        p.paragraph_format.space_after = Pt(Mm(2).pt)

    # ================= PAGES 2..N: MATCHED LISTINGS =================
    for i, page in enumerate(context["listing_pages"]):
        if i == 0:
            # add_page_break() creates a whole new paragraph just to hold the
            # break character -- even with its line spacing collapsed to
            # near-zero, that paragraph is still a distinct object needing
            # its own (however tiny) sliver of room, and the Summary page
            # above is packed right to its very last drop with nothing left
            # to give. Attaching the break to a run on the last existing
            # paragraph (the final commentary bullet) instead avoids adding
            # a paragraph at all, so there's nothing left that needs room.
            doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
        else:
            # Later listing pages break from a full table page, not a
            # packed summary page, so a dedicated page-break paragraph here
            # is the normal, safe case.
            doc.add_page_break()
        heading = doc.add_paragraph()
        _set_run(heading.add_run(f"{context['office_name']} "), size=20, bold=True, upper=True)
        _set_run(heading.add_run("x"), size=20, bold=True, italic=True, color=TEAL, upper=True)
        _set_run(heading.add_run(" Quiet List."), size=20, bold=True, upper=True)
        # .listings-heading { margin-bottom: 6mm }
        heading.paragraph_format.space_after = Pt(Mm(6).pt)

        # template.html renders "Matched Listings:" and the period on two
        # lines of ONE paragraph via <br>, with a single margin-bottom: 6mm
        # after the whole two-line block -- a real line break within one
        # paragraph, not two separate paragraphs, matches that directly.
        period = doc.add_paragraph()
        _set_run(period.add_run("Matched Listings:"), size=9, bold=True, color=TEAL, upper=True)
        period.add_run().add_break()
        _set_run(period.add_run(context["matched_listings_period_label"]), size=9, bold=True, color=TEAL)
        period.paragraph_format.space_after = Pt(Mm(6).pt)

        rows = max(len(page["left"]), len(page["right"]))
        listing_table = doc.add_table(rows=rows, cols=2)
        # table.listings { border: 2.5px solid #000000 } -- explicit, not
        # Word's built-in "Table Grid" style, whose own default weight was
        # rendering visibly thinner than intended.
        _set_full_borders(listing_table, "000000", 20)
        _set_col_widths(listing_table, [8.9, 8.9])
        for i in range(rows):
            left_val = page["left"][i] if i < len(page["left"]) else ""
            right_val = page["right"][i] if i < len(page["right"]) else ""
            for c, val in ((0, left_val), (1, right_val)):
                cell = listing_table.cell(i, c)
                _para(cell, val, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
                _set_cell_margins(cell, top_mm=3, bottom_mm=3, left_mm=4, right_mm=4)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
