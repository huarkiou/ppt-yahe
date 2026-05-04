"""Slide builders with layout-config-based parameter grouping."""

from ppt_yahe.builder.image import build_image_slide
from ppt_yahe.builder.summary import build_summary_slide
from ppt_yahe.builder.types import CellDimensions, ImageLayoutConfig, SummaryLayoutConfig

__all__ = [
    "CellDimensions",
    "ImageLayoutConfig",
    "SummaryLayoutConfig",
    "build_image_slide",
    "build_summary_slide",
]
