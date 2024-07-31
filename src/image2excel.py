import os
import warnings

import pyarrow as pa
from matplotlib.image import imread
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image


def default_path(file_name: str) -> str:
    return os.path.join(os.getcwd(), file_name)


def change_zoom(excel_file: str, zoom: int = 10) -> None:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(excel_file)
        for ws in wb.worksheets:
            ws.sheet_view.zoomScale = zoom
        wb.save(excel_file)


def resize_picture(file: str, max_pixels: int = 300 * 300) -> str:
    with Image.open(file) as image:
        current_pixels = image.width * image.height

        if current_pixels > max_pixels:
            scale = (max_pixels / current_pixels) ** 0.5
            new_size = (int(image.width * scale), int(image.height * scale))

            image = image.resize(new_size, Image.LANCZOS)

            file_name, extension = os.path.splitext(file)
            new_file = f"{file_name}_resized{extension}"
            image.save(new_file)
            return new_file
    return file


def rgb_to_hex(r, g, b):
    return "{:02x}{:02x}{:02x}".format(int(r), int(g), int(b))


def im2xlsx(file: str, resize: bool = True, max_pixels: int = 300 * 300) -> None:
    file = default_path(file) if os.path.split(file)[0] == "" else file

    if resize:
        file = resize_picture(file, max_pixels)

    try:
        img = imread(file)
    except IOError:
        raise FileNotFoundError(f"Unable to read image file: {file}")

    if img.ndim == 3 and img.shape[2] == 4:  # Handle RGBA images
        img = img[:, :, :3]  # Remove alpha channel

    excel_file = os.path.splitext(file)[0] + ".xlsx"

    # Convert image data to a PyArrow Table
    height, width, _ = img.shape
    flattened_img = img.reshape(-1, 3)
    table = pa.Table.from_arrays(
        [pa.array(flattened_img[:, i]) for i in range(3)], names=["r", "g", "b"]
    )

    # Create a new workbook and select the active sheet
    wb = Workbook()
    ws = wb.active

    # Set column width and row height to make cells square
    for col in range(1, width + 1):
        column_letter = get_column_letter(col)
        ws.column_dimensions[column_letter].width = 2.5
    for row in range(1, height + 1):
        ws.row_dimensions[row].height = 15

    # Process the data in chunks
    chunk_size = 10000  # Adjust this based on your memory constraints
    for i in range(0, len(table), chunk_size):
        chunk = table.slice(i, chunk_size)

        # Convert RGB values to hex colors
        hex_colors = [
            rgb_to_hex(r, g, b)
            for r, g, b in zip(
                chunk["r"].to_pylist(), chunk["g"].to_pylist(), chunk["b"].to_pylist()
            )
        ]

        # Fill cells with colors
        for j, color in enumerate(hex_colors):
            row = (i + j) // width + 1
            col = (i + j) % width + 1
            cell = ws.cell(row=row, column=col)
            cell.fill = PatternFill(
                start_color=color, end_color=color, fill_type="solid"
            )

    # Save the workbook
    wb.save(excel_file)

    change_zoom(excel_file)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert image to Excel file with colored cells."
    )
    parser.add_argument("file", help="Path to the image file")
    parser.add_argument(
        "--no-resize",
        action="store_false",
        dest="resize",
        help="Don't resize the image",
    )
    parser.add_argument(
        "--max-pixels",
        type=int,
        default=300 * 300,
        help="Maximum number of pixels in the resized image",
    )

    args = parser.parse_args()

    im2xlsx(args.file, resize=args.resize, max_pixels=args.max_pixels)
    im2xlsx(args.file, resize=args.resize, max_pixels=args.max_pixels)
