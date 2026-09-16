"""check_docx_layout.py -- structural audit of a generated report .docx.

Screenshots and Word's own rendering are not enough to catch the class of bug
this checks for: Word silently tolerates a schema-invalid property order and an
overfull table, so a file can look correct in Word and still be wrong. Google
Docs' importer is strict and discards out-of-sequence properties, which is how
the same file ends up visibly misaligned there. This reads the raw OOXML and
asserts the things a renderer is entitled to get right OR wrong at its own
discretion are simply never ambiguous in the first place.

    python tools/check_docx_layout.py reports/Some_Office_20260916.docx
"""
import sys
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# xsd:sequence orders from ECMA-376 for the container elements this renderer
# builds by hand. A child appearing out of this order is schema-invalid.
SEQUENCES = {
    "tblPr": ("tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
              "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
              "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
              "tblCellMar", "tblLook", "tblCaption", "tblDescription"),
    "tcPr": ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
             "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
             "hideMark"),
    "tcMar": ("top", "start", "left", "bottom", "end", "right"),
    "tblCellMar": ("top", "start", "left", "bottom", "end", "right"),
    "tcBorders": ("top", "start", "left", "bottom", "end", "right",
                  "insideH", "insideV", "tl2br", "tr2bl"),
    "tblBorders": ("top", "start", "left", "bottom", "end", "right",
                   "insideH", "insideV"),
}

DEFAULT_CELL_MARGIN_TWIPS = 108  # Word's built-in left/right default


def ln(el):
    return etree.QName(el).localname


def check_order(root, problems):
    for name, order in SEQUENCES.items():
        for el in root.iter(W + name):
            ranks, seen = [], []
            for child in el:
                cn = ln(child)
                if cn not in order:
                    continue
                ranks.append(order.index(cn))
                seen.append(cn)
            if ranks != sorted(ranks):
                problems.append(
                    f"<w:{name}> children out of schema order: {seen} "
                    f"(expected the order {[o for o in order if o in seen]})")


def cell_h_margin(tc):
    """Horizontal padding actually in force for a cell: its own tcMar if it has
    one, else its table's tblCellMar, else Word's 108-twip default."""
    def pair(mar):
        if mar is None:
            return None
        out = {}
        for side in ("left", "start", "right", "end"):
            el = mar.find(W + side)
            if el is not None:
                out["left" if side in ("left", "start") else "right"] = int(el.get(W + "w"))
        return out if len(out) == 2 else None

    own = pair(tc.find(f"{W}tcPr/{W}tcMar"))
    if own:
        return own["left"], own["right"]
    tbl = tc.getparent().getparent()
    tbl_level = pair(tbl.find(f"{W}tblPr/{W}tblCellMar"))
    if tbl_level:
        return tbl_level["left"], tbl_level["right"]
    return DEFAULT_CELL_MARGIN_TWIPS, DEFAULT_CELL_MARGIN_TWIPS


def grid_total(tbl):
    return sum(int(gc.get(W + "w")) for gc in tbl.findall(f"{W}tblGrid/{W}gridCol"))


def check_widths(root, text_width, problems):
    for tbl in root.iter(W + "tbl"):
        total = grid_total(tbl)
        declared = tbl.find(f"{W}tblPr/{W}tblW")
        ind = tbl.find(f"{W}tblPr/{W}tblInd")
        name = f"table at line {tbl.sourceline}"

        if declared is None or declared.get(W + "type") != "dxa":
            problems.append(f"{name}: no explicit dxa w:tblW")
        elif int(declared.get(W + "w")) != total:
            problems.append(f"{name}: w:tblW {declared.get(W+'w')} != sum of gridCol {total}")

        parent_tc = tbl.getparent() if ln(tbl.getparent()) == "tc" else None
        if parent_tc is None:
            # A top-level table's VISIBLE box lands at
            #   margin + tblInd - <first column's left padding>
            # in both Word and Google Docs, so tblInd must equal that padding
            # for the box to sit on the margin. Nested tables get no such
            # compensation and must keep tblInd=0.
            first_tc = tbl.find(f"{W}tr/{W}tc")
            pad = cell_h_margin(first_tc)[0] if first_tc is not None else 0
            if ind is None or int(ind.get(W + "w")) != pad:
                problems.append(
                    f"{name}: top-level w:tblInd is "
                    f"{ind.get(W+'w') if ind is not None else 'absent'}, expected {pad} "
                    f"(its first column's left padding) — box would sit "
                    f"{pad - (int(ind.get(W+'w')) if ind is not None else 0)} twips "
                    f"left of the page margin")
            if total != text_width:
                problems.append(
                    f"{name}: top-level width {total} twips != page text width {text_width}")
        else:
            tcW = parent_tc.find(f"{W}tcPr/{W}tcW")
            if tcW is None:
                problems.append(f"{name}: nested in a cell with no explicit tcW")
                continue
            if ind is not None and int(ind.get(W + "w")) != 0:
                problems.append(f"{name}: nested table with non-zero w:tblInd "
                                f"{ind.get(W+'w')} (Word applies no indent "
                                f"compensation inside a cell)")
            left, right = cell_h_margin(parent_tc)
            avail = int(tcW.get(W + "w")) - left - right
            if total > avail:
                problems.append(
                    f"{name}: nested width {total} twips OVERFLOWS its cell's "
                    f"content area {avail} (cell {tcW.get(W+'w')} minus padding "
                    f"{left}+{right}) by {total - avail} twips")


def main(path):
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read("word/document.xml"))
    sect = root.find(f".//{W}sectPr/{W}pgSz")
    mar = root.find(f".//{W}sectPr/{W}pgMar")
    text_width = int(sect.get(W + "w")) - int(mar.get(W + "left")) - int(mar.get(W + "right"))

    problems = []
    check_order(root, problems)
    check_widths(root, text_width, problems)

    # the footer is a separate part with its own table
    for part in z.namelist():
        if part.startswith("word/footer") and part.endswith(".xml"):
            froot = etree.fromstring(z.read(part))
            check_order(froot, problems)
            for tbl in froot.iter(W + "tbl"):
                if grid_total(tbl) != text_width:
                    problems.append(
                        f"{part}: footer table width {grid_total(tbl)} != {text_width}")

    print(f"page text width: {text_width} twips")
    if problems:
        print(f"\nFAIL -- {len(problems)} problem(s):")
        for p in dict.fromkeys(problems):
            print(f"  * {p}")
        return 1
    print("PASS -- property order valid, every table width exact, no overflow")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
