from __future__ import annotations

import argparse
import logging
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation

from ppt_yahe.image_slide import add_image_slide
from ppt_yahe.summary_slide import add_summary_slide

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


@dataclass
class Dataset:
    """A group of measurement data producing one summary slide and one image slide.

    Each dataset is a self-contained group of (displacement_levels, section_ids,
    measurement_data). Multiple datasets produce multiple slide pairs in the
    same PPTX.

    Attributes:
        displacement_levels: Labels for displacement levels (rows in the table).
        section_ids: Labels for experiment sections (columns in the table).
        measurement_data: Mapping of (displacement, section_id) → (force, length).
        image_dir: Directory containing experiment images for this dataset.
        filename_template: Template for image filenames. Supports {displacement}
            and {section} placeholders.
    """

    title: str
    displacement_levels: list[str]
    section_ids: list[str]
    measurement_data: dict[tuple[str, str], tuple[float, float]]
    image_dir: str | Path = ""
    filename_template: str = "{displacement}_{section}"


# Add new Dataset() entries here for additional data groups.
# Each dataset will produce its own pair of slides (summary + image matrix).
DATASETS: list[Dataset] = [
    Dataset(
        title="数据组1",
        displacement_levels=["low", "mid", "high"],
        section_ids=["exp1", "exp2", "exp3", "exp4", "exp6", "exp7"],
        measurement_data={
            ("low", "exp1"): (0.12, 0.01),
            ("low", "exp2"): (0.34, 0.02),
            ("low", "exp3"): (0.14, 0.03),
            ("low", "exp4"): (0.33, 0.04),
            ("mid", "exp1"): (0.56, 0.03),
            ("high", "exp4"): (0.78, 0.01),
        },
        image_dir="testdata/images",
    ),
]


def main() -> None:
    """Generate a PowerPoint presentation from one or more measurement datasets.

    Each dataset in DATASETS produces two slides: a summary table with comparison
    chart, and an image matrix grid. Default values match the original
    configuration for backward compatibility.
    """
    parser = argparse.ArgumentParser(description="Generate PPTX from experimental measurement data")
    parser.add_argument(
        "--image-dir",
        default=None,
        help="Override image directory for all datasets",
    )
    parser.add_argument(
        "--output",
        default="testdata/image_matrix.pptx",
        help="Output PPTX file path",
    )
    args = parser.parse_args()

    prs = Presentation()

    for dataset in DATASETS:
        image_dir = args.image_dir if args.image_dir else dataset.image_dir

        add_summary_slide(
            prs,
            title=dataset.title,
            displacement_levels=dataset.displacement_levels,
            section_ids=dataset.section_ids,
            measurement_data=dataset.measurement_data,
        )

        add_image_slide(
            prs,
            image_dir=image_dir,
            displacement_levels=dataset.displacement_levels,
            section_ids=dataset.section_ids,
            filename_template=dataset.filename_template,
            top_left_label=dataset.title,
            measurement_data=dataset.measurement_data,
            supplement_row_ratio=0.30,
        )

    prs.save(args.output)
    logger.info("PPT saved: %s", args.output)


if __name__ == "__main__":
    main()
