from __future__ import annotations

from lxml import etree  # ty:ignore[unresolved-import]
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.table import Table, _Cell
from pptx.util import Inches, Pt


def get_measurement_str(
    measurement_data: dict[tuple[str, str], tuple[float, float]] | None,
    displacement: str,
    section_id: str,
) -> tuple[str, str]:
    """Look up a measurement pair from the measurement data dictionary.

    Given a mapping of ``(displacement, section_id)`` keys to ``(force, length)``
    value pairs, return the string representation of the matching entry.

    Args:
        measurement_data: Nested dictionary keyed by ``(displacement, section_id)``
            tuples with ``(force, length)`` float pairs.  May be ``None``.
        displacement: The displacement level identifier string (first key component).
        section_id: The section identifier string (second key component).

    Returns:
        A 2-tuple of ``(force_str, length_str)``.  If the key is not found or
        *measurement_data* is ``None``, both elements are empty strings.
    """
    if measurement_data is None:
        return ("", "")
    pair = measurement_data.get((displacement, section_id))
    if pair is None:
        return ("", "")
    return (f"{pair[0]:.2f}", f"{pair[1]:.2f}")


def create_styled_table(
    slide: Slide,
    rows: int,
    cols: int,
    left: Inches,
    top: Inches,
    width: Inches,
    height: Inches,
) -> Table:
    """Create a new table shape on a slide and apply the default border style.

    Adds an ``rows × cols`` table at the specified position and dimensions,
    then calls :func:`apply_table_style` to give every cell a thin black
    border.

    Args:
        slide: The target slide to place the table on.
        rows: Number of rows in the table.
        cols: Number of columns in the table.
        left: Distance from the left edge of the slide (in EMU / ``Inches``).
        top: Distance from the top edge of the slide (in EMU / ``Inches``).
        width: Total width of the table (in EMU / ``Inches``).
        height: Total height of the table (in EMU / ``Inches``).

    Returns:
        The newly created |Table| object with styling applied.
    """
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    apply_table_style(table)
    return table


def set_merged_cell(
    table: Table,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    text: str,
    bold: bool = False,
    font_size: float = 10,
    alignment: PP_ALIGN = PP_ALIGN.CENTER,
) -> None:
    """Merge a rectangular region of cells and populate it with formatted text.

    Merges the cell at ``(r1, c1)`` with the cell at ``(r2, c2)`` (inclusive),
    then delegates to :func:`set_cell_text` to write *text* into the top-left
    cell of the merged range with the requested formatting.

    Args:
        table: The table containing the cells to merge.
        r1: Start row index (0‑based).
        c1: Start column index (0‑based).
        r2: End row index (0‑based, inclusive).
        c2: End column index (0‑based, inclusive).
        text: The cell text content.
        bold: Whether the text should be bold (default ``False``).
        font_size: Font size in points (default ``10``).
        alignment: Paragraph alignment (default ``PP_ALIGN.CENTER``).

    Returns:
        ``None``.
    """
    table.cell(r1, c1).merge(table.cell(r2, c2))
    set_cell_text(table.cell(r1, c1), text, bold=bold, font_size=font_size, alignment=alignment)


def apply_table_style(table: Table) -> None:
    """Apply a consistent thin black border to every cell in *table*.

    Operates on the underlying XML (``a:tbl`` element):

    1. Removes any existing ``tableStyleId`` and fill children from the
       table-level properties (``a:tblPr``).
    2. For every cell, removes existing fill and border elements from the
       cell-level properties (``a:tcPr``).
    3. Adds new ``a:lnB`` / ``a:lnT`` / ``a:lnL`` / ``a:lnR`` border
       sub-elements, each 1 pt wide with a solid black (``000000``) fill.

    Args:
        table: The |Table| whose cells should be styled.

    Returns:
        ``None``.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = etree.SubElement(tbl, qn("a:tblPr"))

    tableStyleId = tblPr.find(qn("a:tableStyleId"))
    if tableStyleId is not None:
        tblPr.remove(tableStyleId)

    fill_tags = ["a:solidFill", "a:gradFill", "a:pattFill", "a:noFill", "a:grpFill"]
    for fill_tag in fill_tags:
        for el in tblPr.findall(qn(fill_tag)):
            tblPr.remove(el)

    line_width = Pt(1)
    border_tags = ["a:lnB", "a:lnT", "a:lnL", "a:lnR"]

    for cell in table.iter_cells():
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()

        for fill_tag in fill_tags:
            for el in tcPr.findall(qn(fill_tag)):
                tcPr.remove(el)

        for tag in border_tags:
            for old in tcPr.findall(qn(tag)):
                tcPr.remove(old)

        for tag in border_tags:
            ln = etree.SubElement(tcPr, qn(tag))
            ln.attrib["w"] = str(line_width)
            sf = etree.SubElement(ln, qn("a:solidFill"))
            srgb = etree.SubElement(sf, qn("a:srgbClr"))
            srgb.attrib["val"] = "000000"


def set_cell_text(
    cell: _Cell,
    text: str,
    bold: bool = False,
    font_size: float = 10,
    alignment: PP_ALIGN = PP_ALIGN.CENTER,
) -> None:
    """Set the text content and formatting of a table cell.

    Clears any existing content, then writes *text* into the cell's first
    paragraph with the specified font and alignment.  The vertical anchor is
    set to middle.

    Args:
        cell: The ``_Cell`` object to modify.
        text: The text string to display.
        bold: Whether the text should be bold (default ``False``).
        font_size: Font size in points (default ``10``).
        alignment: Paragraph alignment constant (default ``PP_ALIGN.CENTER``).

    Returns:
        ``None``.
    """
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.name = "微软雅黑"
    run.font.size = Pt(font_size)
    run.font.bold = bold


def column_left_inch(table: Table, col_idx: int) -> float:
    """Calculate the left edge offset of a table column in inches.

    Sums the widths of all columns preceding *col_idx* and returns the total
    in inches.

    Args:
        table: The |Table| containing the column.
        col_idx: The 0‑based index of the target column.

    Returns:
        The cumulative width of columns ``[0, col_idx)`` expressed as a
        ``float`` number of inches.
    """
    total = 0.0
    for c in range(col_idx):
        total += table.columns[c].width / Inches(1)
    return total


def row_top_inch(table: Table, row_idx: int) -> float:
    """Calculate the top edge offset of a table row in inches.

    Sums the heights of all rows preceding *row_idx* and returns the total
    in inches.

    Args:
        table: The |Table| containing the row.
        row_idx: The 0‑based index of the target row.

    Returns:
        The cumulative height of rows ``[0, row_idx)`` expressed as a
        ``float`` number of inches.
    """
    total = 0.0
    for r in range(row_idx):
        total += table.rows[r].height / Inches(1)
    return total
