from __future__ import annotations

import argparse
import logging
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation

from ppt_yahe.builder import build_image_slide, build_summary_slide

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
            ("low", "exp6"): (0.21, 0.02),
            ("low", "exp7"): (0.28, 0.03),
            ("mid", "exp1"): (0.56, 0.03),
            ("mid", "exp2"): (0.48, 0.02),
            ("mid", "exp3"): (0.42, 0.03),
            ("mid", "exp4"): (0.51, 0.04),
            ("mid", "exp6"): (0.45, 0.02),
            ("mid", "exp7"): (0.39, 0.03),
            ("high", "exp1"): (0.71, 0.02),
            ("high", "exp2"): (0.65, 0.03),
            ("high", "exp3"): (0.69, 0.04),
            ("high", "exp4"): (0.78, 0.01),
            ("high", "exp6"): (0.74, 0.02),
            ("high", "exp7"): (0.82, 0.03),
        },
        image_dir="testdata/images",
    ),
    Dataset(
        title="数据组2",
        displacement_levels=["low", "mid", "high"],
        section_ids=["exp3", "exp6", "exp1", "exp7", "exp4", "exp2"],
        measurement_data={
            ("low", "exp1"): (0.22, 1.52),
            ("low", "exp2"): (0.18, 1.68),
            ("low", "exp3"): (0.25, 1.41),
            ("low", "exp4"): (0.20, 1.59),
            ("low", "exp6"): (0.28, 1.33),
            ("low", "exp7"): (0.15, 1.77),
            ("mid", "exp1"): (0.41, 2.03),
            ("mid", "exp2"): (0.37, 2.15),
            ("mid", "exp3"): (0.44, 1.96),
            ("mid", "exp4"): (0.39, 2.09),
            ("mid", "exp6"): (0.49, 1.83),
            ("mid", "exp7"): (0.35, 2.26),
            ("high", "exp1"): (0.62, 2.58),
            ("high", "exp2"): (0.58, 2.67),
            ("high", "exp3"): (0.65, 2.49),
            ("high", "exp4"): (0.60, 2.61),
            ("high", "exp6"): (0.70, 2.33),
            ("high", "exp7"): (0.55, 2.75),
        },
        image_dir="testdata/images",
    ),
    Dataset(
        title="数据组3",
        displacement_levels=["low", "mid", "high"],
        section_ids=["exp7", "exp4", "exp2", "exp6", "exp1", "exp3"],
        measurement_data={
            ("low", "exp1"): (0.09, 3.12),
            ("low", "exp2"): (0.11, 3.04),
            ("low", "exp3"): (0.08, 3.18),
            ("low", "exp4"): (0.10, 3.08),
            ("low", "exp6"): (0.13, 2.95),
            ("low", "exp7"): (0.07, 3.22),
            ("mid", "exp1"): (0.31, 2.44),
            ("mid", "exp2"): (0.35, 2.38),
            ("mid", "exp3"): (0.29, 2.51),
            ("mid", "exp4"): (0.33, 2.41),
            ("mid", "exp6"): (0.38, 2.29),
            ("mid", "exp7"): (0.27, 2.57),
            ("high", "exp1"): (0.52, 1.89),
            ("high", "exp2"): (0.56, 1.82),
            ("high", "exp3"): (0.49, 1.95),
            ("high", "exp4"): (0.54, 1.86),
            ("high", "exp6"): (0.59, 1.74),
            ("high", "exp7"): (0.47, 2.01),
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

        build_summary_slide(
            prs,
            title=dataset.title,
            displacement_levels=dataset.displacement_levels,
            section_ids=dataset.section_ids,
            measurement_data=dataset.measurement_data,
        )

        build_image_slide(
            prs,
            image_dir=image_dir,
            displacement_levels=dataset.displacement_levels,
            section_ids=dataset.section_ids,
            filename_template=dataset.filename_template,
            top_left_label=dataset.title,
            measurement_data=dataset.measurement_data,
        )

    prs.save(args.output)
    logger.info("PPT saved: %s", args.output)


if __name__ == "__main__":
    main()
