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
- a shaded "panel" table standing in for the PDF's rounded Helpful Tip
  circle and the teal Key Insights box — Word has no easy equivalent to a
  CSS border-radius circle, so that one specific shape is a deliberate
  simplification, not an oversight.

One thing to flag for whoever opens the result in Word: page/table background
shading is a Word feature that is not always included by default when
printing or exporting to PDF (Word: File > Options > Display > "Print
background colors and images" must be on) — the teal panels use cell
shading rather than a page background specifically to avoid depending on
that setting, but it's worth knowing if colors ever seem to vanish on export.
"""
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

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
    return run


def _para(doc_or_cell, text="", size=10, bold=False, italic=False, color=BLACK,
          align=None, upper=False, style=None):
    """Add a styled paragraph to a Document or a table cell.

    A fresh Document has zero paragraphs; a fresh table cell always has
    exactly one, empty one (OOXML requires it). Reusing that existing empty
    paragraph for the first piece of content — rather than always calling
    add_paragraph(), which appends a new one after it — avoids leaving a
    stray blank line above every single cell's content. Once that paragraph
    has a run in it, later calls correctly fall through to add_paragraph()
    for subsequent lines in the same cell.
    """
    existing = getattr(doc_or_cell, "paragraphs", None)
    if existing and len(existing) == 1 and not existing[0].runs and not style:
        p = existing[0]
    else:
        p = doc_or_cell.add_paragraph(style=style) if style else doc_or_cell.add_paragraph()
    if align is not None:
        p.alignment = align
    if text:
        _set_run(p.add_run(text), size=size, bold=bold, italic=italic, color=color, upper=upper)
    return p


def _set_col_widths(table, widths_cm):
    for row in table.rows:
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)


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
    _para(footer_table.cell(0, 0), context["footer_contact"], size=7)
    _para(footer_table.cell(0, 1), context["logo_text"], size=9, bold=True,
          align=WD_ALIGN_PARAGRAPH.RIGHT)

    # ================= PAGE 1: SUMMARY =================

    header_table = doc.add_table(rows=1, cols=2)
    _no_borders(header_table)
    _set_col_widths(header_table, [11.0, 6.8])
    left = header_table.cell(0, 0)
    _para(left, context["office_name"], size=24, bold=True, upper=True)
    _para(left, "Quiet List Exchange Activity Report", size=11, bold=True)
    right = header_table.cell(0, 1)
    for label, value in (
        (context["reporting_period_label"], context["office_name"]),
        ("REPORTING PERIOD", context["period_range_label"]),
        ("PREPARED FOR", context["prepared_for"]),
        ("DATE", context["report_date_label"]),
    ):
        p = _para(right, align=WD_ALIGN_PARAGRAPH.RIGHT)
        _set_run(p.add_run(f"{label}: "), size=8, color=TEAL)
        _set_run(p.add_run(value), size=8, bold=True, color=BLACK)

    doc.add_paragraph()
    exec_content = _labeled_row(doc, "Executive Snapshot")
    exec_table = exec_content.add_table(rows=1, cols=3)
    exec_table.style = "Table Grid"
    exec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = exec_table.rows[0].cells
    for i, label in enumerate(("Metric", "Result", "Change vs. Previous Period")):
        _para(hdr[i], label, size=8, bold=True, upper=True,
              align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    for row in context["exec_snapshot"]:
        cells = exec_table.add_row().cells
        _para(cells[0], row["label"], size=8.5, bold=True, upper=True)
        _para(cells[1], row["result"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _para(cells[2], row["change"], size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              color=TEAL if row["positive"] else RED)
    _clear_empty_leading_paragraph(exec_content)

    doc.add_paragraph()
    perf_content = _labeled_row(doc, "Performance Overview")
    perf_wrap = perf_content.add_table(rows=1, cols=2)
    _no_borders(perf_wrap)
    _set_col_widths(perf_wrap, [9.0, 5.8])
    perf_cell, tip_cell = perf_wrap.cell(0, 0), perf_wrap.cell(0, 1)

    perf_table = perf_cell.add_table(rows=1, cols=3)
    perf_table.style = "Table Grid"
    hdr = perf_table.rows[0].cells
    for i, label in enumerate(("Property Type", "Listings", "Matches")):
        _para(hdr[i], label, size=8, bold=True, upper=True,
              align=WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER)
    for row in context["performance_overview"]:
        cells = perf_table.add_row().cells
        _para(cells[0], row["property_type"], size=8.5, bold=True, upper=True)
        _para(cells[1], row["listings"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _para(cells[2], row["matches"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    _clear_empty_leading_paragraph(perf_cell)

    _shade_cell(tip_cell, "105652")
    tip_cell.vertical_alignment = 1  # center
    _para(tip_cell, "Helpful Tip", size=8, bold=True, color=WHITE, upper=True,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    _para(tip_cell, context["helpful_tip"], size=7, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    _clear_empty_leading_paragraph(perf_content)

    doc.add_paragraph()
    insights = doc.add_table(rows=1, cols=3)
    _no_borders(insights)
    for c in range(3):
        _shade_cell(insights.cell(0, c), "105652")
    col1, col2, col3 = insights.cell(0, 0), insights.cell(0, 1), insights.cell(0, 2)

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

    doc.add_paragraph()
    _para(doc, "Commentary:", size=9, bold=True, color=TEAL, upper=True)
    for bullet in context["commentary"]:
        _para(doc, bullet, size=8.5, style="List Bullet")

    # ================= PAGES 2..N: MATCHED LISTINGS =================
    for page in context["listing_pages"]:
        doc.add_page_break()
        heading = doc.add_paragraph()
        _set_run(heading.add_run(f"{context['office_name']} "), size=18, bold=True, upper=True)
        _set_run(heading.add_run("x"), size=18, bold=True, italic=True, color=TEAL)
        _set_run(heading.add_run(" Quiet List."), size=18, bold=True, upper=True)

        p = doc.add_paragraph()
        _set_run(p.add_run("Matched Listings:"), size=9, bold=True, color=TEAL, upper=True)
        p2 = doc.add_paragraph()
        _set_run(p2.add_run(context["matched_listings_period_label"]), size=9, bold=True, color=TEAL)

        rows = max(len(page["left"]), len(page["right"]))
        listing_table = doc.add_table(rows=rows, cols=2)
        listing_table.style = "Table Grid"
        for i in range(rows):
            left_val = page["left"][i] if i < len(page["left"]) else ""
            right_val = page["right"][i] if i < len(page["right"]) else ""
            _para(listing_table.cell(i, 0), left_val, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
            _para(listing_table.cell(i, 1), right_val, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
