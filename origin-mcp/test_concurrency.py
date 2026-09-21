# -*- coding: utf-8 -*-
"""Origin COM 单飞与超时语义的无 Origin 回归测试。"""

from __future__ import annotations

import sys
import threading
import time
import types
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "origin-mcp"))
sys.path.insert(0, str(ROOT / "origin-auto-skill" / "scripts"))

import origin_mcp_server as mcp_server  # noqa: E402
import origin_session as fallback  # noqa: E402


WORKERS = (mcp_server.ComWorker, fallback.ComWorker)


class WorkerTimeoutTests(unittest.TestCase):
    def setUp(self):
        fake_pythoncom = types.SimpleNamespace(
            CoInitialize=lambda: None,
            CoUninitialize=lambda: None,
        )
        self.pythoncom_patch = patch.dict(sys.modules, {"pythoncom": fake_pythoncom})
        self.pythoncom_patch.start()

    def tearDown(self):
        self.pythoncom_patch.stop()

    def test_queued_timeout_cancels_task_before_side_effect(self):
        for worker_class in WORKERS:
            with self.subTest(worker=worker_class.__module__):
                worker = worker_class()
                started = threading.Event()
                release = threading.Event()
                side_effects = []

                def blocking_call():
                    started.set()
                    release.wait(1)
                    return "first"

                first_result = {}
                first = threading.Thread(
                    target=lambda: first_result.setdefault(
                        "value", worker.submit(blocking_call, timeout=1)
                    )
                )
                first.start()
                self.assertTrue(started.wait(1))

                with self.assertRaisesRegex(TimeoutError, "已取消且不会执行"):
                    worker.submit(lambda: side_effects.append("ran"), timeout=0.02)

                release.set()
                first.join(1)
                self.assertEqual(first_result.get("value"), "first")
                time.sleep(0.02)
                self.assertEqual(side_effects, [])
                worker.stop()

    def test_started_task_returns_real_result_instead_of_false_timeout(self):
        for worker_class in WORKERS:
            with self.subTest(worker=worker_class.__module__):
                worker = worker_class()
                started = threading.Event()
                release = threading.Event()
                result = {}

                def slow_call():
                    started.set()
                    release.wait(1)
                    return "done"

                caller = threading.Thread(
                    target=lambda: result.setdefault(
                        "value", worker.submit(slow_call, timeout=0.02)
                    )
                )
                caller.start()
                self.assertTrue(started.wait(1))
                time.sleep(0.04)
                self.assertTrue(caller.is_alive(), "已开始的调用不应伪装成超时失败")
                release.set()
                caller.join(1)
                self.assertEqual(result.get("value"), "done")
                worker.stop()


class ToolSingleFlightTests(unittest.TestCase):
    def test_concurrent_tool_call_is_rejected_without_execution(self):
        entered = threading.Event()
        release = threading.Event()
        calls = []

        @mcp_server._origin_single_flight
        def stateful_tool(value):
            calls.append(value)
            entered.set()
            release.wait(1)
            return {"ok": True, "value": value}

        first_result = {}
        first = threading.Thread(
            target=lambda: first_result.setdefault("value", stateful_tool("first"))
        )
        first.start()
        self.assertTrue(entered.wait(1))

        second = stateful_tool("second")
        self.assertFalse(second["ok"])
        self.assertTrue(second["busy"])
        self.assertTrue(second["not_executed"])
        self.assertEqual(calls, ["first"])

        release.set()
        first.join(1)
        self.assertEqual(first_result["value"], {"ok": True, "value": "first"})


if __name__ == "__main__":
    unittest.main(verbosity=2)