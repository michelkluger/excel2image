# image2excel

[![CI](https://github.com/michelkluger/excel2image/actions/workflows/ci.yml/badge.svg)](https://github.com/michelkluger/excel2image/actions/workflows/ci.yml)
[![Python 3.12+](https://img.shields.io/badge/python-3.12%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
[![Typed](https://img.shields.io/badge/type--checked-ty-blue.svg)](https://github.com/astral-sh/ty)

Convert any image into an Excel spreadsheet where **each cell is a pixel**, colored to recreate the original image.

## Installation

```bash
pip install image2excel
```

Or with [uv](https://docs.astral.sh/uv/):

```bash
uv add image2excel
```

## Quick Start

### Python API

```python
from image2excel import im2xlsx

# Basic usage - creates photo.xlsx alongside the source image
im2xlsx("photo.png")

# Keep original size (no resizing)
im2xlsx("photo.png", resize=False)

# Resize while preserving aspect ratio
im2xlsx("photo.png", keep_aspect=True)
```

### Command Line

```bash
# Convert an image
image2excel photo.png

# Skip resizing
image2excel photo.png --no-resize

# Preserve aspect ratio
image2excel photo.png --keep-aspect
```

## How It Works

1. Opens the image and converts it to RGB
2. Optionally resizes to fit within 260x300 pixels (configurable)
3. Maps each pixel to an Excel cell with a matching background color
4. Sets the zoom to 10% so you can see the full picture

The output `.xlsx` file is saved next to the source image.

## API Reference

### `im2xlsx(file, *, resize=True, keep_aspect=False) -> Path`

| Parameter     | Type         | Default | Description                                |
| ------------- | ------------ | ------- | ------------------------------------------ |
| `file`        | `str \| Path` | -       | Path to the source image                   |
| `resize`      | `bool`       | `True`  | Shrink images larger than 260x300          |
| `keep_aspect` | `bool`       | `False` | Preserve aspect ratio when resizing        |

**Returns:** `Path` to the generated `.xlsx` file.

## Development

```bash
# Clone and install with dev dependencies
git clone https://github.com/michelkluger/excel2image.git
cd excel2image
uv sync --dev

# Run tests
uv run pytest

# Lint and format
uv run ruff check .
uv run ruff format .

# Type check
uv run ty check src/
```

## License

[MIT](LICENSE)
