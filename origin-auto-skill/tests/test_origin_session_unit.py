# -*- coding: utf-8 -*-
"""不启动 Origin 的回退层单元测试。

这些测试专门覆盖 COM 交互前后的确定性逻辑；真实 Origin 链路仍由
``scripts/test_session.py`` 验证。
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import origin_session as mod  # noqa: E402
from origin_session import OriginSession  # noqa: E402


class NamingTests(unittest.TestCase):
    def test_page_name_sanitization_matches_origin_rules(self):
        cases = {
            "Data_0.6": "Data06",
            "Force_Displacement_0.6": "ForceDisplace",
            "My Book": "MyBook",
            "中文名": "A",
            "123456789012345": "1234567890123",
        }
        for raw, expected in cases.items():
            with self.subTest(raw=raw):
                self.assertEqual(mod.sanitize_page_name(raw), expected)

    def test_name_key_ignores_case_and_punctuation(self):
        self.assertEqual(mod._name_key("Data_0.6"), mod._name_key("data06"))

    def test_unique_name_is_computed_after_sanitization(self):
        sess = OriginSession()
        sess.all_pages = lambda: [("ForceDisplace", "Worksheet")]
        self.assertEqual(sess._unique_name("Force_Displacement_0.6"), "ForceDisplac2")

    def test_resolve_page_accepts_unsanitized_alias_and_kind(self):
        sess = OriginSession()
        sess.all_pages = lambda: [
            ("Data06", "Worksheet"),
            ("Data062", "Graph"),
        ]
        self.assertEqual(sess._resolve_page("Data_0.6", "worksheet"), "Data06")


class PlotExpressionTests(unittest.TestCase):
    def test_column_letters_cross_boundaries(self):
        cases = {1: "A", 26: "Z", 27: "AA", 52: "AZ", 53: "BA", 702: "ZZ"}
        for index, expected in cases.items():
            with self.subTest(index=index):
                self.assertEqual(mod._col_letter(index), expected)

    def test_runs_split_non_adjacent_y_columns(self):
        self.assertEqual(
            OriginSession._runs([5, 2, 3, 3, 8]),
            [[2, 3], [5], [8]],
        )

    def test_iy_uses_origin_multi_y_range_without_nested_parentheses(self):
        self.assertEqual(OriginSession._iy(1, [2]), "(1,2)")
        self.assertEqual(OriginSession._iy(1, [2, 3, 4]), "(1,2:4)")

    def test_unknown_plot_type_is_rejected_before_com_calls(self):
        sess = OriginSession()
        with self.assertRaisesRegex(ValueError, "未知 plot_type"):
            sess.plot_create("Book", 1, [2], plot_type="heatmap")


class DataPutTests(unittest.TestCase):
    def test_empty_data_is_a_noop(self):
        self.assertEqual(OriginSession().data_put("Book", []), (0, 0))

    def test_block_write_uses_zero_based_com_offsets_and_pads_rows(self):
        sess = OriginSession()
        scripts = []
        captured = {}
        sess._resolve_page = lambda name, kind=None: "Data06"
        sess.ex = lambda script: scripts.append(script) or True

        def put_block(book, rect, row, col):
            captured.update(book=book, rect=rect, row=row, col=col)
            return True

        sess._put_block = put_block
        result = sess.data_put(
            "Data_0.6",
            [[1, 2], [3]],
            header=["Time", "Signal"],
            start_row=2,
            start_col=2,
        )
        self.assertEqual(result, (2, 2))
        self.assertEqual(captured, {
            "book": "Data06",
            "rect": [(1, 2), (3, "")],
            "row": 1,
            "col": 1,
        })
        self.assertIn("wks.ncols < 3", scripts[0])
        self.assertIn('hhB[L]$ = "Time"', scripts[0])

    def test_chunked_write_advances_start_row(self):
        sess = OriginSession()
        calls = []
        sess._put_once = lambda book, rect, row, col, timeout=300: (
            calls.append((len(rect), row, col)) or True
        )
        rows = [[i] for i in range(mod.PUT_CHUNK_ROWS + 1)]
        self.assertTrue(sess._put_chunked("Book", rows, 4, 2))
        self.assertEqual(calls, [(mod.PUT_CHUNK_ROWS, 4, 2), (1, 4 + mod.PUT_CHUNK_ROWS, 2)])

    def test_oversized_data_never_falls_back_to_cell_by_cell(self):
        sess = OriginSession()
        sess._resolve_page = lambda name, kind=None: name
        sess.ex = lambda script: True
        sess._put_block = lambda *args: False
        sess._put_chunked = lambda *args: False
        rows = [[0, 1] for _ in range(mod.CELL_FALLBACK_MAX_CELLS // 2 + 1)]
        with self.assertRaisesRegex(RuntimeError, "超过逐格兜底上限"):
            sess.data_put("Book", rows)


class AxisAndOutputSafetyTests(unittest.TestCase):
    def test_axis_limits_automatically_lock_rescale(self):
        sess = OriginSession()
        emitted = []
        sess._resolve_page = lambda name, kind=None: "Graph1"
        sess.ex = lambda script: emitted.append(script) or True
        self.assertTrue(sess.axis_set("Graph1", "y", vmin=0, vmax=30))
        self.assertIn("layer.y.from = 0.0;", emitted[0])
        self.assertIn("layer.y.to = 30.0;", emitted[0])
        self.assertIn("layer.y.rescale = 0;", emitted[0])

    def test_unsupported_export_format_fails_before_com(self):
        with tempfile.TemporaryDirectory() as td:
            with self.assertRaisesRegex(ValueError, "不支持的导出格式"):
                OriginSession().graph_export(str(Path(td) / "plot.svg"))

    def test_empty_project_cannot_overwrite_large_existing_project(self):
        sess = OriginSession()
        sess.all_pages = lambda: []
        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "result.opju"
            target.write_bytes(b"x" * 100_001)
            with self.assertRaisesRegex(RuntimeError, "拒绝用空工程覆盖"):
                sess.project_save(str(target))
            self.assertEqual(target.stat().st_size, 100_001)
            self.assertFalse(Path(str(target) + ".bak").exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
