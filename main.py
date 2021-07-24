import math
import sys
from pathlib import Path

import numpy as np
import xlsxwriter
from PIL import Image

MAX_NUM_COLORS = 65490  # The maximum amount of decorators allowed in one XLSX file.


def validate_CLI():
    """
    Validates command line inputs for a valid filepath, file extension, and cell width.
    """
    errors = []

    if len(sys.argv) != 3:
        print("Incorrect number of command line arguments.")
        exit(1)

    file_path = Path(sys.argv[1])

    # If no error produced, these functions return None.
    errors.append(validate_image_file_path(file_path))
    errors.append(validate_image_path_extension(file_path))
    errors.append(validate_cell_size())

    errors_filtered = [i for i in errors if i]  # Remove all none values
    if len(errors_filtered) != 0:
        print(*errors_filtered, sep='\n')
        exit(1)

    print("Command line inputs validated successfully...")


def validate_image_file_path(file_path: Path) -> None or FileNotFoundError:
    """
    Validates that the file path given points to file.

    :param file_path: Path to the file containing an image.
    :return: None if successful; FileNotFoundError returned if file does not exist on this path.
    """
    try:
        file_path.resolve(strict=True)  # Confirm file exists
    except FileNotFoundError as err:
        return err


def validate_image_path_extension(file_path: Path) -> None or str:
    """
    Ensures that the file path's suffix is any one of a list of valid file extensions.

    :param: file_path: Path to the file containing an image.
    :return: None if valid image format; str containing error message if not.
    """
    valid_file_types = ['.bmp', '.jpeg', '.jpg', '.png']

    if file_path.suffix not in valid_file_types:
        return (f"Given path does contain a valid ending. It ends in {file_path.suffix}. "
                f"It must end in one of: {*valid_file_types,}")


def validate_cell_size() -> None or str:
    """
    Ensures that the given cell width is an integer within the bounds of 0 and 100.

    :return: None if correct size; otherwise str containing error message.
    """
    try:
        num = int(sys.argv[2])
        var = 0 < num < 100
        if not var:
            raise ValueError
    except ValueError:
        return "Cell dimension must be convertible to an integer, and have a value between 0 and 100."


def load_image_from_file(filename: str) -> Image:
    """
    Loads image from given file and converts it to RGB format.

    :param filename: Path to file containing image to be converted
    :return: loaded PIL image
    """
    try:
        picture = Image.open(filename)
        picture.load()
        picture.convert('RGB')
        return picture
    except Exception as err:
        raise err


def adjust_image_for_xslx_compatibility(pil_img: Image) -> np.array:
    """
    If the image contains more colours than Excel allows in one workbook,
    then the size of the image is reduced incrementally
    until the number of unique colors is less than this maximum amount.

    :param pil_img: Image object (PIL).
    :return: rgb_array: Numpy array containing each individual pixel's RGB value.
    """
    max_width = float('inf')
    unique_colours = np.empty(MAX_NUM_COLORS)
    resized = False
    print("Validating image color profile...")
    while len(unique_colours) >= MAX_NUM_COLORS:  # This runs at least once, as the initial array will be too large.
        rgb_array = convert_pil_img_to_rgb_array(pil_img)
        _, unique_colours = np.unique(rgb_array.reshape(-1, 3), axis=0, return_counts=1)

        if len(unique_colours) >= MAX_NUM_COLORS:
            img_w, img_h = pil_img.size
            if img_w <= max_width:
                max_width = img_w
            pil_img = resize_img(pil_img, img_w, img_h)
            resized = True

    if resized:
        print(f"Image adjusted in size to ensure xslx compatibility. "
              f"New dimensions: {pil_img.size[0]} x {pil_img.size[1]} px with {len(unique_colours)} colours...")
    else:
        print("Image valid...")

    return rgb_array


def resize_img(img: Image, img_w: int, img_h: int) -> Image:
    """
    Resizes the image by a reduction factor of 20%.

    :param img: Original PIL image.
    :param img_w: Original image width.
    :param img_h: Original image height.
    :return: res: PIL image, with dimensions reduced by 20%.
    """
    new_width = math.floor(img_w * 0.8)  # Rounding down is best for our application.
    new_height = math.floor(img_h * 0.8)
    try:
        res = img.resize((new_width, new_height))
        return res
    except ValueError:
        raise ValueError


def convert_pil_img_to_rgb_array(pil_img: Image) -> Image:
    """
    Convert PIL image to a 3D array of size (image_width * image_width * 3)
    where each cell contains the a single pixel's RGB value (RRR, GGG, BBB).

    :param pil_img: Original PIL Image.
    :return: Image: PIL Image converted to NumPy array.
    """
    return np.asarray(pil_img, dtype="uint32")


def convert_rgb_array_to_hex_array(rgb_arr: np.array) -> np.array:
    """
    Maps the function `rgb_array_to_hex_string` over every item in the np array.
    Speed is increased by converting from a flattening the array, mapping the function,
    and then converting back to a 3d array.
    Taken from: https://stackoverflow.com/questions/22424096/apply-functions-to-3d-numpy-array

    :param rgb_arr: NumPy array containing RGB values.
    :return: reshaped_arr: NumPy array containing hex values.
    """
    x, y, z = rgb_arr.shape

    reshaped_arr = rgb_arr.reshape(x * y, z)
    reshaped_arr = np.apply_along_axis(rgb_array_to_hex_string, 1, reshaped_arr)
    reshaped_arr = reshaped_arr.reshape(x, y, 1)
    return reshaped_arr


def rgb_array_to_hex_string(cell: list) -> str:
    """
    Converts RGB array to hex string.

    :param cell: RGB array of a single cell.
    :return: Hex string.
    """
    return "#{:02x}{:02x}{:02x}".format(cell[0], cell[1], cell[2])


def make_excel_file(filename: str) -> (xlsxwriter.Workbook.worksheet_class, xlsxwriter.Workbook):
    wb = xlsxwriter.Workbook(f'output/{filename}.xlsx')  # Create new workbook at this location
    ws = wb.add_worksheet()
    return ws, wb


def write_cells(arr: np.array, ws: xlsxwriter.Workbook.worksheet_class, wb: xlsxwriter.Workbook, cell_size: int):
    """
    Writes array of pixels to spreadsheet, incrementally formatting each cell as it goes through.

    :param arr: NumPy array containing hex values for pixel of the final image to be written.
    :param ws: xlsxwriter worksheet
    :param wb: xlsx workbook
    :param cell_size: desired cell height in pixels
    """
    row_num = 0
    col_num = 0
    for row in arr:
        ws.set_row_pixels(row_num, cell_size)  # Set row height
        for col in row:
            cell_format = wb.add_format()  # Add format to cell
            cell_format.set_bg_color(col[0])
            ws.write(row_num, col_num, '', cell_format)
            col_num += 1
        col_num = 0  # Reset to 0 at end of line
        row_num += 1

    num_cols_occupied = arr.shape[1] - 1
    worksheet.set_column_pixels(0, num_cols_occupied, cell_size)  # Resize used columns to desired width
    print("Worksheet filled successfully; please wait...")


if __name__ == '__main__':
    validate_CLI()
    image_file_path = sys.argv[1]
    cell_dimensions = int(sys.argv[2])
    image_file_stem = Path(image_file_path).stem
    loaded_img = load_image_from_file(image_file_path)
    final_rgb_array = adjust_image_for_xslx_compatibility(loaded_img)
    final_hex_array = convert_rgb_array_to_hex_array(final_rgb_array)
    worksheet, workbook = make_excel_file(image_file_stem)
    write_cells(final_hex_array, worksheet, workbook, cell_dimensions)

    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as ex:
        raise ex

    loaded_img.close()
    print("Workbook successfully saved (*￣▽￣)b")
