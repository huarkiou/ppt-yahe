from pptx import Presentation

from ppt_yahe.image_slide import add_image_slide
from ppt_yahe.summary_slide import add_summary_slide

IMAGE_DIR = r"testdata/images"
OUTPUT_PPT = r"testdata/image_matrix.pptx"

DISPLACEMENT_LEVELS = ["low", "mid", "high"]
SECTION_IDS = ["exp1", "exp2", "exp3", "exp4", "exp6", "exp7"]

MEASUREMENT_DATA = {
    ("low", "exp1"): (0.12, 0.01),
    ("low", "exp2"): (0.34, 0.02),
    ("low", "exp3"): (0.14, 0.03),
    ("low", "exp4"): (0.33, 0.04),
    ("mid", "exp1"): (0.56, 0.03),
    ("high", "exp4"): (0.78, 0.01),
}


def main() -> None:
    prs = Presentation()

    add_summary_slide(
        prs,
        displacement_levels=DISPLACEMENT_LEVELS,
        section_ids=SECTION_IDS,
        measurement_data=MEASUREMENT_DATA,
    )

    add_image_slide(
        prs,
        image_dir=IMAGE_DIR,
        displacement_levels=DISPLACEMENT_LEVELS,
        section_ids=SECTION_IDS,
        filename_template="{displacement}_{section}.png",
        top_left_label="",
        measurement_data=MEASUREMENT_DATA,
        supplement_row_ratio=0.30,
    )

    prs.save(OUTPUT_PPT)
    print(f"[OK] PPT 已保存: {OUTPUT_PPT}")


if __name__ == "__main__":
    main()
