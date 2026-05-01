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
    """
    生成一页 PPT，主体为一个表格：
      - 每个 (a, b) 对应一个 2 行 × 2 列 的子区域（但图片行两列合并）：
          · 图片行：图片居中（跨已合并的 2 列）
          · 补充行：左右两个独立小格，用于补充信息
      - 所有图片的外接正方形包围盒大小统一。

    Args:
        image_dir:           图片所在目录
        output_ppt:          输出 PPT 文件路径（.pptx）
        param_a_values:      参数 a 取值列表（决定行组数）
        param_b_values:      参数 b 取值列表（决定列组数）
        filename_template:   文件名模板，如 "{a}_{b}.png"
        top_left_label:      表格左上角单元格文本（默认空）
        header_col_width:    行标题列宽度（英寸）
        header_row_height:   列标题行高度（英寸）
        supplement_data:     补充信息字典，key 为 (a_val, b_val) 元组，
                             value 为 (左格文本, 右格文本) 元组；无数据则留空
        supp_row_ratio:      补充行占每组 2 行总高的比例（0~1），默认 0.30
    """
    image_dir: Path = Path(image_dir)
    if supplement_data is None:
        supplement_data = {}

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式

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

    # 数据区尺寸
    data_area_width = usable_width - header_col_width
    data_area_height = usable_height - header_row_height

    total_data_cols = n_cols * 2  # 每个 b 占 2 小列
    total_data_rows = n_rows * 2  # 每个 a 占 2 小行

    sub_cell_width = data_area_width / total_data_cols
    sub_cell_height = data_area_height / total_data_rows

    two_row_height = 2 * sub_cell_height
    supp_row_height = two_row_height * supp_row_ratio
    img_row_height = two_row_height - supp_row_height

    # 图片区域尺寸（跨 2 小列 × 图片行高）
    img_area_width = sub_cell_width * 2
    img_area_height = img_row_height

    padding = 0.08
    square_size = min(img_area_width, img_area_height) - 2 * padding

    # ---------------------------- 创建表格 ----------------------------
    total_cols = 1 + total_data_cols  # 1 行标题列 + 数据列
    total_rows = 1 + total_data_rows  # 1 列标题行 + 数据行

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
        table.rows[1 + i * 2].height = Inches(img_row_height)  # 图片行
        table.rows[1 + i * 2 + 1].height = Inches(supp_row_height)  # 补充行

    # ---------------------------- 标题 ----------------------------
    _set_cell_text(table.cell(0, 0), top_left_label, bold=True, font_size=11)

    # 列标题：每 2 列合并
    for j in range(n_cols):
        col_a = 1 + j * 2
        col_b = col_a + 1
        table.cell(0, col_a).merge(table.cell(0, col_b))
        _set_cell_text(table.cell(0, col_a), param_b_values[j], bold=True, font_size=11)

    # 行标题：每 2 行合并（图片行 + 补充行为一组）
    for i in range(n_rows):
        row_a = 1 + i * 2
        row_b = row_a + 1
        table.cell(row_a, 0).merge(table.cell(row_b, 0))
        _set_cell_text(table.cell(row_a, 0), param_a_values[i], bold=True, font_size=11)

    # ---------------------------- 合并图片行单元格 & 填写补充信息 ----------------------------
    for i, a_val in enumerate(param_a_values):
        for j, b_val in enumerate(param_b_values):
            img_row = 1 + i * 2
            supp_row = img_row + 1
            col_a = 1 + j * 2
            col_b = col_a + 1

            # 图片行：两列合并为一个单元格
            table.cell(img_row, col_a).merge(table.cell(img_row, col_b))

            # 补充行：左右独立
            info = supplement_data.get((a_val, b_val), ("", ""))
            left_text, right_text = info
            _set_cell_text(table.cell(supp_row, col_a), left_text, font_size=8)
            _set_cell_text(table.cell(supp_row, col_b), right_text, font_size=8)

    # ---------------------------- 放置图片 ----------------------------
    # 关键修正：数据区域起点
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

            # 长边撑满 square_size
            if aspect >= 1:
                disp_w = square_size
                disp_h = square_size / aspect
            else:
                disp_h = square_size
                disp_w = square_size * aspect

            # 图片在图片区域（img_area_width × img_area_height）内居中
            offset_x = (img_area_width - disp_w) / 2
            offset_y = (img_area_height - disp_h) / 2

            # 图片区域左上角在数据区中的位置
            block_left = j * img_area_width
            block_top = i * two_row_height

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
