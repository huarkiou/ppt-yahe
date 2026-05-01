from pptx.table import Table, _Cell
from pptx.slide import Slide
from pathlib import Path
from pptx import Presentation, presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml.html import etree
from PIL import Image

# PowerPoint 默认幻灯片尺寸（单位：英寸）
DEFAULT_SLIDE_WIDTH_INCHES = 10.0
DEFAULT_SLIDE_HEIGHT_INCHES = 7.5


def main():
    IMAGE_DIR = r"testdata/images"
    OUTPUT_PPT = r"testdata/image_matrix.pptx"

    PARAM_A = ["low", "mid", "high"]
    PARAM_B = ["exp1", "exp2", "exp3", "exp4"]

    SUPPLEMENT = {
        ("low", "exp1"): ("0.12", "±0.01"),
        ("low", "exp2"): ("0.34", "±0.02"),
        ("low", "exp3"): ("0.14", "±0.03"),
        ("low", "exp4"): ("0.33", "±0.04"),
        ("mid", "exp1"): ("0.56", "±0.03"),
        ("high", "exp4"): ("0.78", "±0.01"),
    }

    prs = Presentation()
    prs.slide_width = Inches(DEFAULT_SLIDE_WIDTH_INCHES)
    prs.slide_height = Inches(DEFAULT_SLIDE_HEIGHT_INCHES)

    add_summary_table_slide(
        prs,
        param_a_values=PARAM_A,
        param_b_values=PARAM_B,
        supplement_data=SUPPLEMENT,
    )

    add_image_table_slide(
        prs,
        image_dir=IMAGE_DIR,
        param_a_values=PARAM_A,
        param_b_values=PARAM_B,
        filename_template="{a}_{b}.png",
        top_left_label="卧槽牛逼",
        supplement_data=SUPPLEMENT,
        supp_row_ratio=0.30,
    )

    prs.save(OUTPUT_PPT)
    print(f"✅ PPT 已保存: {OUTPUT_PPT}")


def _apply_table_style(table: Table):
    """去掉表格样式和所有填充，设置黑色细框线"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = etree.SubElement(tbl, qn("a:tblPr"))

    # 1) 干掉表格样式 ID（这是蓝白碗的根源）
    tableStyleId = tblPr.find(qn("a:tableStyleId"))
    if tableStyleId is not None:
        tblPr.remove(tableStyleId)

    # 2) 清除表格级别的所有填充
    fill_tags = ["a:solidFill", "a:gradFill", "a:pattFill", "a:noFill", "a:grpFill"]
    for fill_tag in fill_tags:
        for el in tblPr.findall(qn(fill_tag)):
            tblPr.remove(el)

    # 3) 每个单元格：清填充 + 设细黑框线
    line_width = Pt(1)

    ln_tags = ["a:lnB", "a:lnT", "a:lnL", "a:lnR"]
    for cell in table.iter_cells():
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()

        # 清除所有类型填充
        for fill_tag in fill_tags:
            for el in tcPr.findall(qn(fill_tag)):
                tcPr.remove(el)

        # 移除旧边框
        for tag in ln_tags:
            for old in tcPr.findall(qn(tag)):
                tcPr.remove(old)

        # 添加黑色细边框
        for tag in ln_tags:
            ln = etree.SubElement(tcPr, qn(tag))
            ln.attrib["w"] = str(line_width)
            sf = etree.SubElement(ln, qn("a:solidFill"))
            srgb = etree.SubElement(sf, qn("a:srgbClr"))
            srgb.attrib["val"] = "000000"


def _set_cell_text(
    cell: _Cell,
    text: str,
    bold: bool = False,
    font_size: float = 10,
    alignment: PP_ALIGN = PP_ALIGN.CENTER,
):
    """设置表格单元格文本、字体及对齐方式"""
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold


def add_image_table_slide(
    prs: presentation.Presentation,
    image_dir: str,
    param_a_values: list[str],
    param_b_values: list[str],
    filename_template: str = "{a}_{b}.png",
    top_left_label: str = "",
    header_col_width: float = 1.2,
    header_row_height: float = 0.5,
    supplement_data: dict | None = None,
    supp_row_ratio: float = 0.30,
):
    image_dir: Path = Path(image_dir)
    if supplement_data is None:
        supplement_data = {}

    slide: Slide = prs.slides.add_slide(prs.slide_layouts[6])

    slide_width: float = prs.slide_width / Inches(1)
    slide_height: float = prs.slide_height / Inches(1)  # ty:ignore[unsupported-operator]

    margin_lr = 0.0
    margin_tb = 0.3
    usable_width = slide_width - 2 * margin_lr
    usable_height = slide_height - 2 * margin_tb

    n_rows = len(param_a_values)
    n_cols = len(param_b_values)
    if n_rows == 0 or n_cols == 0:
        raise ValueError("参数列表不能为空")

    data_area_width = usable_width - header_col_width
    data_area_height = usable_height - header_row_height

    total_data_cols = n_cols * 2  # 每个 b 占 2 小列
    total_data_rows = n_rows * 3  # 每个 a 占 3 小行（图片 + 标题 + 数据）

    sub_cell_width = data_area_width / total_data_cols
    sub_cell_height = data_area_height / total_data_rows

    three_row_height = 3 * sub_cell_height
    supp_total_height = three_row_height * supp_row_ratio
    img_row_height = three_row_height - supp_total_height
    title_row_height = supp_total_height / 2  # 标题行
    data_row_height = supp_total_height / 2  # 数据行

    img_area_width = sub_cell_width * 2
    img_area_height = img_row_height

    padding = 0.08
    square_size = min(img_area_width, img_area_height) - 2 * padding

    # ---------------------------- 创建表格 ----------------------------
    total_cols = 1 + total_data_cols
    total_rows = 1 + total_data_rows

    table_shape = slide.shapes.add_table(
        total_rows,
        total_cols,
        Inches(margin_lr),
        Inches(margin_tb),
        Inches(usable_width),
        Inches(usable_height),
    )
    table = table_shape.table
    _apply_table_style(table)

    # 列宽
    table.columns[0].width = Inches(header_col_width)
    for c in range(1, total_cols):
        table.columns[c].width = Inches(sub_cell_width)

    # 行高
    table.rows[0].height = Inches(header_row_height)
    for i in range(n_rows):
        base = 1 + i * 3
        table.rows[base].height = Inches(img_row_height)  # 图片行
        table.rows[base + 1].height = Inches(title_row_height)  # 标题行
        table.rows[base + 2].height = Inches(data_row_height)  # 数据行

    # ---------------------------- 标题 ----------------------------
    _set_cell_text(table.cell(0, 0), top_left_label, bold=True, font_size=11)

    for j in range(n_cols):
        col_a = 1 + j * 2
        col_b = col_a + 1
        table.cell(0, col_a).merge(table.cell(0, col_b))
        _set_cell_text(table.cell(0, col_a), param_b_values[j], bold=True, font_size=11)

    for i in range(n_rows):
        row_a = 1 + i * 3
        row_b = row_a + 2
        table.cell(row_a, 0).merge(table.cell(row_b, 0))
        _set_cell_text(table.cell(row_a, 0), param_a_values[i], bold=True, font_size=11)

    # ---------------------------- 合并单元格 & 填写内容 ----------------------------
    for i, a_val in enumerate(param_a_values):
        for j, b_val in enumerate(param_b_values):
            img_row = 1 + i * 3
            title_row = img_row + 1
            data_row = img_row + 2
            col_a = 1 + j * 2
            col_b = col_a + 1

            # 图片行：两列合并
            table.cell(img_row, col_a).merge(table.cell(img_row, col_b))

            # 标题行：两列独立，填固定标题
            _set_cell_text(table.cell(title_row, col_a), "力值", bold=True, font_size=8)
            _set_cell_text(table.cell(title_row, col_b), "长度", bold=True, font_size=8)

            # 数据行：两列独立，填补充数据
            info = supplement_data.get((a_val, b_val), ("", ""))
            left_text, right_text = info
            _set_cell_text(table.cell(data_row, col_a), left_text, font_size=8)
            _set_cell_text(table.cell(data_row, col_b), right_text, font_size=8)

    # ---------------------------- 放置图片 ----------------------------
    data_origin_left = margin_lr + header_col_width
    data_origin_top = margin_tb + header_row_height

    for i, a_val in enumerate(param_a_values):
        for j, b_val in enumerate(param_b_values):
            filename = filename_template.format(a=a_val, b=b_val)
            img_path = image_dir / filename
            if not img_path.exists():
                print(f"⚠️ 图片不存在，跳过 — {img_path}")
                continue

            with Image.open(img_path) as im:
                orig_w, orig_h = im.size
            aspect = orig_w / orig_h

            if aspect >= 1:
                disp_w = square_size
                disp_h = square_size / aspect
            else:
                disp_h = square_size
                disp_w = square_size * aspect

            offset_x = (img_area_width - disp_w) / 2
            offset_y = (img_area_height - disp_h) / 2

            block_left = j * img_area_width
            block_top = i * three_row_height

            left = Inches(data_origin_left + block_left + offset_x)
            top = Inches(data_origin_top + block_top + offset_y)

            slide.shapes.add_picture(
                str(img_path),
                left=left,
                top=top,
                width=Inches(disp_w),
            )


def add_summary_table_slide(
    prs: presentation.Presentation,
    param_a_values: list[str],
    param_b_values: list[str],
    supplement_data: dict,  # {(a, b): (力值, 长度)}
    header_col_width: float = 0.9,
    param_col_width: float = 1.0,
    criteria_col_width: float = 1.2,
    header_row_height: float = 0.5,
    data_row_height: float = 0.35,
    title: str = "",
) -> None:
    """
    追加一页汇总表，表格宽度占满幻灯片。
    行：汇总表(合并) | 列标题 | 力值(n行) | 长度(n行) | 结论 | 备注
    列：类别 | 参数 | 评判标准 | B1 | B2 | ...
    依赖：_set_cell_text(table.cell, text, bold, font_size) 已在外部定义。
    """

    n_rows = len(param_a_values)
    n_cols = len(param_b_values)
    total_cols = 3 + n_cols
    total_rows = 2 + n_rows * 2 + 2  # 汇总表 + 列标题 + 力值 + 长度 + 结论 + 备注

    slide: Slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_width_emu = prs.slide_width  # 直接用 EMU，不引入魔数
    slide_height_emu = prs.slide_height  # noqa: F841

    # ---- 幻灯片级可选标题 ----
    if title:
        title_box = slide.shapes.add_textbox(
            Inches(0.3), Inches(0.15), slide_width_emu - Inches(0.6), Inches(0.4)
        )
        tf = title_box.text_frame
        tf.text = title
        p = tf.paragraphs[0]
        p.font.size = Pt(16)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER

    # ---- 表格占满幻灯片宽度 ----
    # 前3列固定英寸宽度，余下给数据列均分
    fixed = header_col_width + param_col_width + criteria_col_width
    # 用 Inches 算出总表宽后得出数据列英寸宽
    total_table_inches = slide_width_emu / Inches(
        1
    )  # 通过 Inches(1) 得到一英寸对应的 EMU
    data_col_width = max(0, (total_table_inches - fixed) / n_cols) if n_cols else 0

    table_top = Inches(0.7) if title else Inches(0.4)
    table_height = Inches(
        data_row_height + header_row_height + (total_rows - 2) * data_row_height
    )

    table_shape = slide.shapes.add_table(
        total_rows,
        total_cols,
        Inches(0),
        table_top,
        slide_width_emu,  # 占满幻灯片宽度
        table_height,
    )
    table = table_shape.table
    _apply_table_style(table)

    # 列宽
    table.columns[0].width = Inches(header_col_width)
    table.columns[1].width = Inches(param_col_width)
    table.columns[2].width = Inches(criteria_col_width)
    for c in range(3, total_cols):
        table.columns[c].width = Inches(data_col_width)

    # 行高
    table.rows[0].height = Inches(data_row_height * 1.2)  # 汇总表行稍高
    table.rows[1].height = Inches(header_row_height)
    for r in range(2, total_rows):
        table.rows[r].height = Inches(data_row_height)

    # ===================== 汇总表标题行（第0行） =====================
    table.cell(0, 0).merge(table.cell(0, total_cols - 1))
    _set_cell_text(table.cell(0, 0), "汇总表", bold=True, font_size=14)

    # ===================== 列标题行（第1行） =====================
    _set_cell_text(table.cell(1, 0), "类别", bold=True, font_size=11)
    _set_cell_text(table.cell(1, 1), "参数", bold=True, font_size=11)
    _set_cell_text(table.cell(1, 2), "评判标准", bold=True, font_size=11)
    for j, b_val in enumerate(param_b_values):
        _set_cell_text(table.cell(1, 3 + j), b_val, bold=True, font_size=11)

    # ===================== 力值区域 (行 2 .. 1+n_rows) =====================
    force_start, force_end = 2, 1 + n_rows
    _set_cell_text(table.cell(force_start, 0), "力值", bold=True, font_size=11)
    if n_rows > 1:
        table.cell(force_start, 0).merge(table.cell(force_end, 0))
    for i, a_val in enumerate(param_a_values):
        row_idx = force_start + i
        _set_cell_text(table.cell(row_idx, 1), a_val, bold=True, font_size=11)
        _set_cell_text(table.cell(row_idx, 2), "", font_size=10)
        for j, b_val in enumerate(param_b_values):
            info = supplement_data.get((a_val, b_val), ("", ""))
            _set_cell_text(table.cell(row_idx, 3 + j), info[0], font_size=10)

    # ===================== 长度区域 (行 2+n_rows .. 1+n_rows*2) =====================
    length_start, length_end = 2 + n_rows, 1 + n_rows * 2
    _set_cell_text(table.cell(length_start, 0), "长度", bold=True, font_size=11)
    if n_rows > 1:
        table.cell(length_start, 0).merge(table.cell(length_end, 0))
    for i, a_val in enumerate(param_a_values):
        row_idx = length_start + i
        _set_cell_text(table.cell(row_idx, 1), a_val, bold=True, font_size=11)
        _set_cell_text(table.cell(row_idx, 2), "", font_size=10)
        for j, b_val in enumerate(param_b_values):
            info = supplement_data.get((a_val, b_val), ("", ""))
            _set_cell_text(table.cell(row_idx, 3 + j), info[1], font_size=10)

    # ===================== 结论 =====================
    conclusion_row = 2 + n_rows * 2
    _set_cell_text(table.cell(conclusion_row, 0), "结论", bold=True, font_size=11)
    table.cell(conclusion_row, 1).merge(table.cell(conclusion_row, total_cols - 1))
    _set_cell_text(table.cell(conclusion_row, 1), "", font_size=10)

    # ===================== 备注 =====================
    remark_row = conclusion_row + 1
    _set_cell_text(table.cell(remark_row, 0), "备注", bold=True, font_size=11)
    table.cell(remark_row, 1).merge(table.cell(remark_row, total_cols - 1))
    _set_cell_text(table.cell(remark_row, 1), "", font_size=10)


if __name__ == "__main__":
    main()
