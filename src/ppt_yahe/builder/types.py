"""Layout configuration dataclasses for slide builders.

Groups visual and spatial parameters into typed objects, eliminating
long parameter lists in slide-generation functions.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class SummaryLayoutConfig:
    """Visual and layout parameters for the summary slide.

    All dimension values are in inches; font sizes are in points.
    """

    # --- column widths ---
    header_col_width: float = 0.9
    param_col_width: float = 1.0
    criteria_col_width: float = 1.2

    # --- row heights ---
    header_row_height: float = 0.4
    data_row_height: float = 0.3

    # --- font sizes ---
    font_size_title: float = 14
    font_size_header: float = 12
    font_size_chart: float = 10

    # --- chart layout ---
    chart_top_padding: float = 0.0
    chart_bottom_padding: float = 0.2
    chart_max_height: float = 3.5
    chart_left_indent: float = 1.6

    # --- table positioning ---
    table_top_offset: float = 0.4


@dataclass
class ImageLayoutConfig:
    """Visual and layout parameters for the image-matrix slide.

    All dimension values are in inches; font sizes are in points;
    ``image_padding_pt`` is in points.
    """

    # --- margins ---
    margin_lr: float = 0.0
    margin_tb: float = 0.5

    # --- cell dimensions ---
    header_col_width: float = 1.2
    header_row_height: float = 0.5

    # --- row layout ---
    supplement_row_ratio: float = 0.30

    # --- font sizes ---
    font_size_header: float = 12
    font_size_data: float = 10

    # --- image ---
    image_padding_pt: float = 1.0


@dataclass
class CellDimensions:
    """Computed cell dimensions for the image-matrix table.

    All values in inches unless otherwise noted.  ``padding`` is in inches
    (converted from points during computation).
    """

    # --- area ---
    data_area_width: float
    data_area_height: float

    # --- column / row counts ---
    total_data_cols: int
    total_data_rows: int
    total_cols: int
    total_rows: int

    # --- cell sizes ---
    sub_cell_width: float
    sub_cell_height: float

    # --- row heights ---
    three_row_height: float
    supplement_total_height: float
    image_row_height: float
    title_row_height: float
    data_row_height: float

    # --- image placement ---
    image_area_width: float
    image_area_height: float
    padding: float
    available_width: float
    available_height: float
    square_size: float
