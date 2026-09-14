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
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
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


def _clear_empty_leading_paragraph(cell):
    """A table cell always starts with exactly one empty paragraph (OOXML
    requires at least one) — this only matters when content is added as a
    nested table/object rather than through _para(), which already reuses
    that paragraph itself. Call this after adding a nested table to a cell,
    or the empty paragraph is left sitting above it as a stray blank line."""
    if cell.paragraphs and not cell.paragraphs[0].runs:
        cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)


def _set_run(run, size=10, bold=False, italic=False, color=BLACK, upper=False):
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
    run.font.name = "Arial"
    rFonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "Arial")
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
          align=None, upper=False, style=None):
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
        _set_run(p.add_run(text), size=size, bold=bold, italic=italic, color=color, upper=upper)
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
    """
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr.find(qn("w:tblLayout")) is None:
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tbl_pr.append(layout)

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for grid_col, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
            grid_col.set(qn("w:w"), str(int(Cm(w).twips)))

    for row in table.rows:
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)


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
                       title_size_pt=7, body_size_pt=6):
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
                  <w:pPr><w:jc w:val="center"/></w:pPr>
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
            <wps:bodyPr wrap="square" lIns="91440" tIns="91440" rIns="91440" bIns="91440" anchor="ctr">
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
    30mm/~3cm in the CSS) beside a wider content column (.row-content) — a
    sidebar layout, not a heading placed above the content. Returns the
    content cell for the caller to build into."""
    wrap = doc.add_table(rows=1, cols=2)
    _no_borders(wrap)
    _set_col_widths(wrap, [3.0, 14.8])
    label_cell, content_cell = wrap.cell(0, 0), wrap.cell(0, 1)
    _para(label_cell, label_text, size=9, bold=True, color=TEAL, upper=True)
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

    header_table = doc.add_table(rows=1, cols=2)
    _no_borders(header_table)
    _set_col_widths(header_table, [11.0, 6.8])
    left = header_table.cell(0, 0)
    _para(left, context["office_name"], size=24, bold=True, upper=True)
    _para(left, "Quiet List Exchange Activity Report", size=11, bold=True, upper=True)
    right = header_table.cell(0, 1)
    for label, value in (
        (context["reporting_period_label"], context["office_name"]),
        ("REPORTING PERIOD", context["period_range_label"]),
        ("PREPARED FOR", context["prepared_for"]),
        ("DATE", context["report_date_label"]),
    ):
        p = _para(right, align=WD_ALIGN_PARAGRAPH.RIGHT)
        _set_run(p.add_run(f"{label}: "), size=9, color=TEAL)
        _set_run(p.add_run(value), size=9, bold=True, color=BLACK)

    _spacer(doc, pt=3)
    exec_content = _labeled_row(doc, "Executive Snapshot")
    exec_table = exec_content.add_table(rows=1, cols=3)
    exec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = exec_table.rows[0].cells
    for i, label in enumerate(("Metric", "Result", "Change vs. Previous Period")):
        _para(hdr[i], label, size=8, bold=True, upper=True,
              align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_bottom_border(hdr[i], "105652", 12)
        _set_cell_margins(hdr[i], top_mm=1.3, bottom_mm=1.3, left_mm=3, right_mm=3)
    for row in context["exec_snapshot"]:
        cells = exec_table.add_row().cells
        _para(cells[0], row["label"], size=8.5, bold=True, upper=True)
        _para(cells[1], row["result"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _para(cells[2], row["change"], size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              color=TEAL if row["positive"] else RED)
        for cell in cells:
            _set_cell_bottom_border(cell, "B9C4BF", 6)
            _set_cell_margins(cell, top_mm=1.6, bottom_mm=1.6, left_mm=3, right_mm=3)
    _clear_empty_leading_paragraph(exec_content)

    _spacer(doc, pt=3)
    perf_content = _labeled_row(doc, "Performance Overview")
    perf_wrap = perf_content.add_table(rows=1, cols=2)
    _no_borders(perf_wrap)
    _set_col_widths(perf_wrap, [9.0, 5.8])
    perf_cell, tip_cell = perf_wrap.cell(0, 0), perf_wrap.cell(0, 1)

    perf_table = perf_cell.add_table(rows=1, cols=3)
    hdr = perf_table.rows[0].cells
    for i, label in enumerate(("Property Type", "Listings", "Matches")):
        _para(hdr[i], label, size=8, bold=True, upper=True,
              align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_bottom_border(hdr[i], "105652", 12)
        _set_cell_margins(hdr[i], top_mm=1.3, bottom_mm=1.3, left_mm=3, right_mm=3)
    for row in context["performance_overview"]:
        cells = perf_table.add_row().cells
        _para(cells[0], row["property_type"], size=8.5, bold=True, upper=True)
        _para(cells[1], row["listings"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _para(cells[2], row["matches"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        for cell in cells:
            _set_cell_bottom_border(cell, "B9C4BF", 6)
            _set_cell_margins(cell, top_mm=1.6, bottom_mm=1.6, left_mm=3, right_mm=3)
    _clear_empty_leading_paragraph(perf_cell)

    tip_cell.vertical_alignment = 1  # center
    tip_para = _para(tip_cell, align=WD_ALIGN_PARAGRAPH.CENTER)
    _add_circle_shape(tip_para, diameter_mm=34, fill_hex="105652",
                       title="Helpful Tip", body_text=context["helpful_tip"])
    _clear_empty_leading_paragraph(perf_content)

    _spacer(doc, pt=3)
    insights = doc.add_table(rows=1, cols=3)
    _no_borders(insights)
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

    _para(col1, "Key Insights", size=9, bold=True, color=WHITE, upper=True)
    _para(col1, "Featured Listing", size=8, bold=True, italic=True, color=WHITE)
    _para(col1, context["featured_listing_address"], size=8, color=WHITE)
    _para(col1, context["featured_listing_caption"], size=8, color=WHITE)
    _para(col1, "Top Performing Suburb", size=8, bold=True, italic=True, color=WHITE)
    _para(col1, context["top_suburb_display"], size=8, color=WHITE)

    _para(col2, "Most active budget range.", size=8, bold=True, italic=True, color=WHITE,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    _para(col2, context["budget_range_display"], size=8, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    _para(col2, "Most active dwelling type.", size=8, bold=True, italic=True, color=WHITE,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    _para(col2, context["dwelling_type_display"], size=8, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)

    _para(col3, "Most Active Buyer's Agencies", size=8, bold=True, italic=True, color=WHITE,
          align=WD_ALIGN_PARAGRAPH.RIGHT)
    for i, name in enumerate(context["top_agencies"], 1):
        _para(col3, f"{i}. {name}", size=8, color=WHITE, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _para(col3, "Operative", size=8, bold=True, italic=True, color=WHITE,
          align=WD_ALIGN_PARAGRAPH.RIGHT)
    for i, name in enumerate(context["top_operatives"], 1):
        _para(col3, f"{i}. {name}", size=8, color=WHITE, align=WD_ALIGN_PARAGRAPH.RIGHT)

    _spacer(doc, pt=3)
    _para(doc, "Commentary:", size=9, bold=True, color=TEAL, upper=True)
    for bullet in context["commentary"]:
        _para(doc, bullet, size=8.5, style="List Bullet")

    # ================= PAGES 2..N: MATCHED LISTINGS =================
    for page in context["listing_pages"]:
        doc.add_page_break()
        heading = doc.add_paragraph()
        _set_run(heading.add_run(f"{context['office_name']} "), size=20, bold=True, upper=True)
        _set_run(heading.add_run("x"), size=20, bold=True, italic=True, color=TEAL, upper=True)
        _set_run(heading.add_run(" Quiet List."), size=20, bold=True, upper=True)

        p = doc.add_paragraph()
        _set_run(p.add_run("Matched Listings:"), size=9, bold=True, color=TEAL, upper=True)
        p2 = doc.add_paragraph()
        _set_run(p2.add_run(context["matched_listings_period_label"]), size=9, bold=True, color=TEAL)

        rows = max(len(page["left"]), len(page["right"]))
        listing_table = doc.add_table(rows=rows, cols=2)
        listing_table.style = "Table Grid"
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
