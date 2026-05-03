from pathlib import Path

from pptx import presentation
from pptx.slide import Slide
from pptx.util import Inches, Pt
from PIL import Image

from ppt_yahe.table_utils import (
    column_left_inch,
    create_styled_table,
    get_measurement_str,
    row_top_inch,
    set_cell_text,
    set_merged_cell,
)


def add_image_slide(
    prs: presentation.Presentation,
    image_dir: str | Path,
    displacement_levels: list[str],
    section_ids: list[str],
    filename_template: str = "{displacement}_{section}",
    top_left_label: str = "",
    header_col_width: float = 1.2,
    header_row_height: float = 0.5,
    measurement_data: dict[tuple[str, str], tuple[float, float]] | None = None,
    supplement_row_ratio: float = 0.30,
    image_fill_uniform: bool = False,
) -> None:
    image_dir = Path(image_dir)
    if measurement_data is None:
        measurement_data = {}

    slide: Slide = prs.slides.add_slide(prs.slide_layouts[6])

    slide_width: float = prs.slide_width / Inches(1)
    slide_height: float = prs.slide_height / Inches(1)

    margin_lr = 0.0
    margin_tb = 0.5
    usable_width = slide_width - 2 * margin_lr
    usable_height = slide_height - 2 * margin_tb

    n_rows = len(displacement_levels)
    n_cols = len(section_ids)
    if n_rows == 0 or n_cols == 0:
        raise ValueError("参数列表不能为空")

    data_area_width = usable_width - header_col_width
    data_area_height = usable_height - header_row_height

    total_data_cols = n_cols * 2
    total_data_rows = n_rows * 3

    sub_cell_width = data_area_width / total_data_cols
    sub_cell_height = data_area_height / total_data_rows

    three_row_height = 3 * sub_cell_height
    supplement_total_height = three_row_height * supplement_row_ratio
    image_row_height = three_row_height - supplement_total_height
    title_row_height = supplement_total_height / 2
    data_row_height = supplement_total_height / 2

    image_area_width = sub_cell_width * 2
    image_area_height = image_row_height

    padding = Pt(1).inches
    available_width = image_area_width - 2 * padding
    available_height = image_area_height - 2 * padding
    square_size = min(available_width, available_height)

    total_cols = 1 + total_data_cols
    total_rows = 1 + total_data_rows

    table = create_styled_table(
        slide,
        total_rows,
        total_cols,
        Inches(margin_lr),
        Inches(margin_tb),
        Inches(usable_width),
        Inches(usable_height),
    )

    table_left = margin_lr
    table_top = margin_tb

    table.columns[0].width = Inches(header_col_width)
    for c in range(1, total_cols):
        table.columns[c].width = Inches(sub_cell_width)

    table.rows[0].height = Inches(header_row_height)
    for i in range(n_rows):
        base = 1 + i * 3
        table.rows[base].height = Inches(image_row_height)
        table.rows[base + 1].height = Inches(title_row_height)
        table.rows[base + 2].height = Inches(data_row_height)

    set_cell_text(table.cell(0, 0), top_left_label, bold=True, font_size=12)

    for j, section_id in enumerate(section_ids):
        col_a = 1 + j * 2
        col_b = col_a + 1
        set_merged_cell(
            table, 0, col_a, 0, col_b, section_id, bold=True, font_size=12
        )

    for i, displacement in enumerate(displacement_levels):
        row_a = 1 + i * 3
        row_b = row_a + 2
        set_merged_cell(
            table, row_a, 0, row_b, 0, displacement, bold=True, font_size=12
        )

    for i, displacement in enumerate(displacement_levels):
        for j, section_id in enumerate(section_ids):
            image_row = 1 + i * 3
            title_row = image_row + 1
            data_row = image_row + 2
            col_a = 1 + j * 2
            col_b = col_a + 1

            table.cell(image_row, col_a).merge(table.cell(image_row, col_b))

            set_cell_text(
                table.cell(title_row, col_a), "力值", bold=True, font_size=10
            )
            set_cell_text(
                table.cell(title_row, col_b), "长度", bold=True, font_size=10
            )

            left_text, right_text = get_measurement_str(
                measurement_data, displacement, section_id
            )
            set_cell_text(table.cell(data_row, col_a), left_text, font_size=10)
            set_cell_text(table.cell(data_row, col_b), right_text, font_size=10)

    for i, displacement in enumerate(displacement_levels):
        for j, section_id in enumerate(section_ids):
            filename = filename_template.format(
                displacement=displacement, section=section_id
            )
            img_path = image_dir / filename
            for suffix in (".png", ".jpg"):
                candidate = img_path.with_suffix(suffix)
                if candidate.exists():
                    img_path = candidate
                    break
            else:
                print(f"[WARN] 跳过不存在的图片: {img_path}.[png|jpg]")
                continue

            with Image.open(img_path) as im:
                original_width, original_height = im.size
            aspect = original_width / original_height

            if image_fill_uniform:
                if aspect >= 1:
                    display_width = square_size
                    display_height = square_size / aspect
                else:
                    display_height = square_size
                    display_width = square_size * aspect
            else:
                if available_width / available_height >= aspect:
                    display_height = available_height
                    display_width = available_height * aspect
                else:
                    display_width = available_width
                    display_height = available_width / aspect

            image_row = 1 + i * 3
            col_a = 1 + j * 2

            cell_left = table_left + column_left_inch(table, col_a)
            cell_top = table_top + row_top_inch(table, image_row)
            cell_w = (
                table.columns[col_a].width + table.columns[col_a + 1].width
            ) / Inches(1)
            cell_h = table.rows[image_row].height / Inches(1)

            offset_x = (cell_w - display_width) / 2
            offset_y = (cell_h - display_height) / 2

            left = Inches(cell_left + offset_x)
            top = Inches(cell_top + offset_y)

            slide.shapes.add_picture(
                str(img_path),
                left=left,
                top=top,
                width=Inches(display_width),
            )
