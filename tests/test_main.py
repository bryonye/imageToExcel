import io
import random
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch

import PIL
import numpy as np
import xlsxwriter

from main import (
    validate_image_file_path,
    validate_image_path_extension,
    load_image_from_file,
    validate_cell_size,
    adjust_image_for_xslx_compatibility,
    resize_img,
    convert_pil_img_to_rgb_array,
    convert_rgb_array_to_hex_array,
    rgb_array_to_hex_string,
    make_excel_file
)


class TestCLIValidation(TestCase):
    def test_validate_image_file_path_valid(self):
        # Pointing to a file
        with tempfile.NamedTemporaryFile() as dirpath:
            file_path = Path(dirpath.name)
            self.assertIsNone(validate_image_file_path(file_path))

    def test_validate_image_path_path_invalid(self):
        # Not pointing to a file
        file_path = Path(f'idontexist{str(random.random())}')
        ret = validate_image_file_path(file_path)
        self.assertIsNotNone(ret)

    def test_validate_image_path_extension_valid_options(self):
        # Valid extension
        valid_file_types = ['.bmp', '.jpeg', '.jpg', '.png']
        for ext in valid_file_types:
            with tempfile.NamedTemporaryFile(suffix=ext) as dirpath:
                file_path = Path(dirpath.name)
                self.assertIsNone(validate_image_path_extension(file_path))

    def test_validate_image_path_extension_invalid_options(self):
        # Invalid extension
        invalid_file_types = ['.txt', '.pdf', '.gif', '.doc', '.mov', '.mp3']
        for ext in invalid_file_types:
            with tempfile.NamedTemporaryFile(suffix=ext) as dirpath:
                file_path = Path(dirpath.name)
                self.assertIsNotNone(validate_image_path_extension(file_path))

    def test_validate_image_path_extension_multiple(self):
        # More than one extension
        valid_file_types = ['.bmp', '.jpeg', '.jpg', '.png']
        invalid_file_types = ['.txt', '.pdf', '.gif', '.doc', '.mov', '.mp3']

        for valid_ext in valid_file_types:
            for invalid_ext in invalid_file_types:
                with tempfile.NamedTemporaryFile(suffix=f"{valid_ext}{invalid_ext}") as dirpath:
                    file_path = Path(dirpath.name)
                    self.assertIsNotNone(validate_image_path_extension(file_path))

    def test_validate_image_path_extension_none(self):
        # No extension
        with tempfile.NamedTemporaryFile(suffix='') as dirpath:
            file_path = Path(dirpath.name)
            self.assertIsNotNone(validate_image_path_extension(file_path))

    def test_validate_cell_size(self):
        # Passing strings
        testargs_string = [
            ["prog", "file", "a sentence"],  # String
            ["prog", "file", "a longer one"],  # String
            ["prog", "file", "-1"],  # Out of bounds
            ["prog", "file", "0"],  # Out of bounds
            ["prog", "file", "101"],  # Out of bounds
            ["prog", "file", "100"],  # Out of bounds
            ["prog", "file", "6.5"],  # Float
            ["prog", "file", "13.453"],  # Float
            ["prog", "file", "[1, 2, 3, 4, 5]"]  # Array of ints
        ]

        for argset in testargs_string:
            with unittest.mock.patch('sys.argv', argset):
                self.assertIsNotNone(validate_cell_size())

    def test_validate_CLI(self):
        pass


class TestImageFunctions(TestCase):
    def test_load_image(self):
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as dirpath:
            img = PIL.Image.new('RGB', (60, 60), color='white')
            img.save(dirpath)
            res = load_image_from_file(dirpath.name)
            self.assertIsInstance(res, PIL.PngImagePlugin.PngImageFile)

        # Pass in file which cannot be opened due to unacceptable extension
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as dirpath:
            with self.assertRaises(PIL.UnidentifiedImageError):
                load_image_from_file(dirpath.name)

    def test_adjust_image_for_xslx_compatibility(self):
        # Input image which doesn't require optimisation
        img = PIL.Image.new('RGB', (60, 60), color='white')
        buf = io.StringIO()
        with redirect_stdout(buf):
            res = adjust_image_for_xslx_compatibility(img)
        self.assertEqual('Validating image color profile...\nImage valid...\n', buf.getvalue())
        self.assertIsInstance(res, np.ndarray)
        self.assertEqual(res.ndim, 3)

        # Input image which requires optimisation
        img = PIL.Image.open("res/flower.jpg")
        buf = io.StringIO()
        with redirect_stdout(buf):
            res = adjust_image_for_xslx_compatibility(img)
        self.assertEqual('Validating image color profile...\n'
                         'Image adjusted in size to ensure xslx compatibility. '
                         'New dimensions: 419 x 202 px with 54165 colours...\n', buf.getvalue())
        self.assertIsInstance(res, np.ndarray)
        self.assertEqual(res.ndim, 3)

    def test_resize_img(self):
        # Input image too small to resize
        img = PIL.Image.new('RGB', (1, 1), color='white')
        with self.assertRaises(ValueError):
            resize_img(img, 1, 1)

        # Input valid image
        img = PIL.Image.new('RGB', (25, 25), color='white')
        res = resize_img(img, 25, 25)
        self.assertEqual(res.size, (20, 20))

    def test_convert_pil_img_to_rgb_array(self):
        # All-white image, check every for 255, 255, 255
        img = PIL.Image.new('RGB', (5, 5), color='white')
        arr = convert_pil_img_to_rgb_array(img)
        for row in arr:
            self.assertTrue((row == 255).all())

        # All-black, same as above
        img = PIL.Image.new('RGB', (5, 5), color='black')
        arr = convert_pil_img_to_rgb_array(img)
        for row in arr:
            self.assertTrue((row == 0).all())

    def test_convert_rgb_array_to_hex_array(self):
        img = PIL.Image.new('RGB', (5, 5), color='white')
        rgb_arr = convert_pil_img_to_rgb_array(img)
        hex_arr = convert_rgb_array_to_hex_array(rgb_arr)

        for row in hex_arr:
            self.assertTrue((row == '#ffffff').all())
        self.assertEqual(hex_arr.ndim, 3)

        img = PIL.Image.new('RGB', (5, 5), color='black')
        rgb_arr = convert_pil_img_to_rgb_array(img)
        hex_arr = convert_rgb_array_to_hex_array(rgb_arr)
        for row in hex_arr:
            self.assertTrue((row == '#000000').all())
        self.assertEqual(hex_arr.ndim, 3)

    def test_rgb_array_to_hex_string(self):
        # Array of known values check
        tests = [
            [[107, 168, 50], '#6ba832'],
            [[73, 83, 222], '#4953de'],
            [[255, 0, 212], '#ff00d4']
        ]

        for suite in tests:
            res = rgb_array_to_hex_string(suite[0])
            self.assertEqual(res, suite[1])


class TestXLSXFunctions(TestCase):
    def test_make_excel_file(self):
        worksheet, workbook = make_excel_file('file')
        self.assertIsInstance(workbook, xlsxwriter.Workbook)
        self.assertIsInstance(worksheet, xlsxwriter.Workbook.worksheet_class)

    def test_write_cells(self):
        pass
