"""Command-line interface for image2excel."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from image2excel.converter import im2xlsx


def main(argv: list[str] | None = None) -> None:
    """Entry point for the ``image2excel`` CLI."""
    parser = argparse.ArgumentParser(
        prog="image2excel",
        description="Convert an image to an Excel spreadsheet where each cell is a colored pixel.",
    )
    parser.add_argument(
        "file",
        type=Path,
        help="path to the source image",
    )
    parser.add_argument(
        "--no-resize",
        action="store_true",
        default=False,
        help="skip automatic resizing of large images",
    )
    parser.add_argument(
        "--keep-aspect",
        action="store_true",
        default=False,
        help="preserve the aspect ratio when resizing",
    )

    args = parser.parse_args(argv)

    try:
        output = im2xlsx(args.file, resize=not args.no_resize, keep_aspect=args.keep_aspect)
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Created {output}")


if __name__ == "__main__":
    main()
