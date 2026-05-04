# ppt-yahe

Generate PowerPoint presentations from experimental measurement data.

## Overview

ppt-yahe creates formatted PPTX slides with:
- **Summary slide** — measurement data table with force/length values per displacement level and section, plus a comparison bar chart
- **Image matrix slide** — grid layout embedding experiment images with measurement labels

## Requirements

- Python >= 3.13
- python-pptx >= 1.0
- Pillow >= 12.2

## Installation

```bash
uv sync
```

## Usage

```bash
uv run python main.py
```

Configure image directory, displacement levels, section IDs, and measurement data in `src/ppt_yahe/__main__.py`.
