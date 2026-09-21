# -*- coding: utf-8 -*-
"""quick_plot 输入解析的无 Origin 单元测试。"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from quick_plot import read_table, resolve_cols  # noqa: E402


class ReadTableTests(unittest.TestCase):
    def test_csv_header_numbers_blanks_and_ragged_rows(self):
        with tempfile.TemporaryDirectory() as td:
            csv_path = Path(td) / "curve.csv"
            csv_path.write_text(
                "Time,Signal A,Signal B\n"
                "0,1,2\n"
                "\n"
                "0.5,3\n",
                encoding="utf-8-sig",
            )
            header, rows = read_table(csv_path, None)
        self.assertEqual(header, ["Time", "Signal A", "Signal B"])
        self.assertEqual(rows, [[0.0, 1.0, 2.0], [0.5, 3.0, ""]])

    def test_numeric_first_row_is_data_not_header(self):
        with tempfile.TemporaryDirectory() as td:
            csv_path = Path(td) / "numeric.csv"
            csv_path.write_text("0,1\n1,2\n", encoding="utf-8")
            header, rows = read_table(csv_path, None)
        self.assertIsNone(header)
        self.assertEqual(rows, [[0.0, 1.0], [1.0, 2.0]])


class ResolveColumnsTests(unittest.TestCase):
    def test_columns_can_mix_numbers_and_header_names(self):
        self.assertEqual(resolve_cols(["Time", "A", "B"], "2,B", "Y"), [2, 3])

    def test_invalid_header_name_returns_actionable_error(self):
        with self.assertRaisesRegex(SystemExit, "Y 列无效: Missing"):
            resolve_cols(["Time", "A"], "Missing", "Y")


if __name__ == "__main__":
    unittest.main(verbosity=2)
