from pptx.table import Table, _Cell
from pptx.slide import Slide
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml.html import etree


def get_measurement_str(
    measurement_data: dict[tuple[str, str], tuple[float, float]] | None,
    displacement: str,
    section_id: str,
) -> tuple[str, str]:
    if measurement_data is None:
        return ("", "")
    pair = measurement_data.get((displacement, section_id))
    if pair is None:
        return ("", "")
    return (str(pair[0]), str(pair[1]))


def create_styled_table(
    slide: Slide,
    rows: int,
    cols: int,
    left: Inches,
    top: Inches,
    width: Inches,
    height: Inches,
) -> Table:
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
    table.cell(r1, c1).merge(table.cell(r2, c2))
    set_cell_text(
        table.cell(r1, c1), text, bold=bold, font_size=font_size, alignment=alignment
    )


def apply_table_style(table: Table) -> None:
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
    total = 0.0
    for c in range(col_idx):
        total += table.columns[c].width / Inches(1)
    return total


def row_top_inch(table: Table, row_idx: int) -> float:
    total = 0.0
    for r in range(row_idx):
        total += table.rows[r].height / Inches(1)
    return total
