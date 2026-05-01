# requirements: pip install python-pptx pillow

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

# PowerPoint 默认幻灯片尺寸（单位：英寸）
DEFAULT_SLIDE_WIDTH_INCHES = 10.0
DEFAULT_SLIDE_HEIGHT_INCHES = 7.5


def generate_image_matrix_ppt(
    image_dir: str,
    output_ppt: str,
    param_a_values: list[str],
    param_b_values: list[str],
    filename_template: str = "{a}_{b}.png",
):
    """
    生成一页 PPT，将图片按参数 a（行）、b（列）排列成矩阵。
    每张图片的外接正方形（包围盒）大小统一，图片在各自单元格内居中。

    Args:
        image_dir: 图片所在目录
        output_ppt: 输出 PPT 文件路径（.pptx）
        param_a_values: 参数 a 的所有取值（决定行数）
        param_b_values: 参数 b 的所有取值（决定列数）
        filename_template: 文件名模板，如 "{a}_{b}.png"
    """
    image_dir: Path = Path(image_dir)

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式

    slide_width = DEFAULT_SLIDE_WIDTH_INCHES
    slide_height = DEFAULT_SLIDE_HEIGHT_INCHES

    # 页面边距（英寸）
    margin_lr = 0.5
    margin_tb = 0.75
    usable_width = slide_width - 2 * margin_lr
    usable_height = slide_height - 2 * margin_tb

    n_rows = len(param_a_values)
    n_cols = len(param_b_values)

    if n_rows == 0 or n_cols == 0:
        raise ValueError("参数列表不能为空")

    # 每个单元格的尺寸
    cell_width = usable_width / n_cols
    cell_height = usable_height / n_rows

    # 正方形包围盒边长 —— 取单元格宽、高中较小者，留少量内边距
    padding = 0.08  # 包围盒与单元格边缘的间距（英寸）
    square_size = min(cell_width, cell_height) - 2 * padding

    for i, a_val in enumerate(param_a_values):
        for j, b_val in enumerate(param_b_values):
            filename = filename_template.format(a=a_val, b=b_val)
            img_path = image_dir / filename

            if not img_path.exists():
                print(f"⚠️ 警告: 图片不存在，跳过 — {img_path}")
                continue

            # 读取图片原始宽高
            with Image.open(img_path) as im:
                orig_w, orig_h = im.size
            aspect = orig_w / orig_h

            # 根据宽高比计算实际显示尺寸（长边撑满 square_size）
            if aspect >= 1:  # 宽图或正方形
                disp_w = square_size
                disp_h = square_size / aspect
            else:  # 高图
                disp_h = square_size
                disp_w = square_size * aspect

            # 在单元格内居中
            offset_x = (cell_width - disp_w) / 2
            offset_y = (cell_height - disp_h) / 2

            left = Inches(margin_lr + j * cell_width + offset_x)
            top = Inches(margin_tb + i * cell_height + offset_y)

            slide.shapes.add_picture(
                str(img_path),
                left=left,
                top=top,
                width=Inches(disp_w),
            )

    prs.save(output_ppt)
    print(f"✅ PPT 已保存: {output_ppt}")


# === 使用示例 ===
if __name__ == "__main__":
    IMAGE_DIR = r"testdata/images"
    OUTPUT_PPT = r"testdata/image_matrix.pptx"

    PARAM_A = ["low", "mid", "high"]  # 行
    PARAM_B = ["exp1", "exp2", "exp3", "exp4"]  # 列

    generate_image_matrix_ppt(
        image_dir=IMAGE_DIR,
        output_ppt=OUTPUT_PPT,
        param_a_values=PARAM_A,
        param_b_values=PARAM_B,
        filename_template="{a}_{b}.png",
    )
