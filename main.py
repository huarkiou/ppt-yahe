from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree  # ty:ignore[unresolved-import]
from PIL import Image

# PowerPoint 默认幻灯片尺寸（单位：英寸）
DEFAULT_SLIDE_WIDTH_INCHES = 10.0
DEFAULT_SLIDE_HEIGHT_INCHES = 7.5


def _apply_table_style(table):
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
    for fill_tag in [
        "a:solidFill",
        "a:gradFill",
        "a:pattFill",
        "a:noFill",
        "a:grpFill",
    ]:
        for el in tblPr.findall(qn(fill_tag)):
            tblPr.remove(el)

    # 3) 每个单元格：清填充 + 设细黑框线
    line_width = Pt(1)
    fill_tags = ["a:solidFill", "a:gradFill", "a:pattFill", "a:noFill", "a:grpFill"]

    for cell in table.iter_cells():
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()

        # 清除所有类型填充
        for fill_tag in fill_tags:
            for el in tcPr.findall(qn(fill_tag)):
                tcPr.remove(el)

        # 移除旧边框
        for tag in ["a:lnB", "a:lnT", "a:lnL", "a:lnR"]:
            for old in tcPr.findall(qn(tag)):
                tcPr.remove(old)

        # 添加黑色细边框
        for tag in ["a:lnB", "a:lnT", "a:lnL", "a:lnR"]:
            ln = etree.SubElement(tcPr, qn(tag))
            ln.attrib["w"] = str(line_width)
            sf = etree.SubElement(ln, qn("a:solidFill"))
            srgb = etree.SubElement(sf, qn("a:srgbClr"))
            srgb.attrib["val"] = "000000"


def _set_cell_text(
    cell,
    text: str,
    bold: bool = False,
    font_size: float = 10,
    alignment=PP_ALIGN.CENTER,
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


def generate_image_matrix_ppt(
    image_dir: str,
    output_ppt: str,
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

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    slide_width = DEFAULT_SLIDE_WIDTH_INCHES
    slide_height = DEFAULT_SLIDE_HEIGHT_INCHES

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

    prs.save(output_ppt)
    print(f"✅ PPT 已保存: {output_ppt}")


# ========================= 使用示例 =========================
if __name__ == "__main__":
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

    generate_image_matrix_ppt(
        image_dir=IMAGE_DIR,
        output_ppt=OUTPUT_PPT,
        param_a_values=PARAM_A,
        param_b_values=PARAM_B,
        filename_template="{a}_{b}.png",
        top_left_label="卧槽牛逼",
        supplement_data=SUPPLEMENT,
        supp_row_ratio=0.30,
    )
