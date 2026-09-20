"""Regression checks that do not need WPS, LibreOffice, or a display."""

import unittest
from pathlib import Path

from et_converter import ConverterApp, file_group


class FakeApp:
    def __init__(self):
        self.callback = None
        self.result = None

    def after(self, _delay, callback):
        self.callback = callback

    def _complete(self, success, detail):
        self.result = (success, detail)


class WorkerErrorTests(unittest.TestCase):
    def test_wps_file_types_are_classified(self):
        self.assertEqual(file_group(Path("book.et")), "spreadsheet")
        self.assertEqual(file_group(Path("document.wps")), "document")
        self.assertEqual(file_group(Path("slides.dps")), "presentation")
        self.assertEqual(file_group(Path("document.ofd")), "fixed_layout")

    def test_conversion_error_is_reported_after_worker_returns(self):
        """The deferred Tk callback must retain the exception text."""
        app = FakeApp()
        ConverterApp._convert_worker(app, Path("missing.et"), Path("result.xlsx"))

        self.assertIsNotNone(app.callback)
        app.callback()
        self.assertFalse(app.result[0])
        self.assertIn("备用方案也失败", app.result[1])


if __name__ == "__main__":
    unittest.main()
