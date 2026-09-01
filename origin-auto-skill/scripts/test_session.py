# -*- coding: utf-8 -*-
"""
test_session.py — origin-auto-skill 回退脚本的自检（真实 Origin）。

验证 origin_session.py 与 origin-mcp 服务端同源的那批修复：
  窗口名净化 / 未净化名自动解析 / 大表写入 / 绘图回读实际名 /
  COM 轴回读 / series_style 空图层兜底 / 空工程覆盖保护

用法：
    python scripts/test_session.py
结束码 0 = 全部通过。
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from origin_session import OriginSession, sanitize_page_name  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "test_output"
OUT.mkdir(exist_ok=True)

RESULTS: list[tuple[str, bool, str]] = []


def step(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, bool(ok), detail))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))


def main() -> int:
    print("=== origin-auto-skill 回退脚本自检 ===")

    step("净化规则与实测一致",
         sanitize_page_name("Data_0.6") == "Data06"
         and sanitize_page_name("Force_Displacement_0.6") == "ForceDisplace"
         and sanitize_page_name("My Book") == "MyBook",
         "Data_0.6->Data06, Force_Displacement_0.6->ForceDisplace")

    sess = OriginSession()
    try:
        sess.connect(visible=True)
        step("连接 Origin", True)

        # 清场：避免上一次运行的残留窗口导致断言误失败
        sess.project_new()
        time.sleep(0.3)

        # -- 窗口名净化 ------------------------------------------------
        b = sess.workbook_new("Data_0.6")
        step("workbook_new 返回实际净化名", b == "Data06", f"book={b}")

        r, c = sess.data_put("Data_0.6", [[1, 2], [3, 4]])
        step("用未净化名写入（自动解析）", r == 2 and c == 2, f"{r}x{c}")

        # -- 大表 ------------------------------------------------------
        big = [[float(i * 20 + k) for k in range(20)] for i in range(40000)]
        t0 = time.time()
        r, c = sess.data_put(b, big)
        dt = time.time() - t0
        step("40000x20 写入", r == 40000 and c == 20 and dt < 60,
             f"{r}x{c} 耗时 {dt:.2f}s")

        back = sess.read_worksheet(b)
        step("回读一致", len(back) == 40000
             and float(back[39999][19]) == 39999 * 20 + 19,
             f"{len(back)} 行, 末值={back[-1][19] if back else None}")

        # -- 绘图 ------------------------------------------------------
        graphs = sess.plot_create(b, x_col=1, y_cols=[2, 4],
                                  plot_type="line", graph_name="Fig_0.6")
        step("plot_create 返回实际净化名", graphs and graphs[0] == "Fig06",
             f"graphs={graphs}")
        g = graphs[0]

        # -- 轴 --------------------------------------------------------
        sess.axis_set(g, "y", title="Force (N)", vmin=0, vmax=30)
        ax = sess.axis_get(g, "y")
        step("COM 通道读回轴范围（含自动锁轴）",
             ax.get("ok") and ax.get("from") == 0.0 and ax.get("to") == 30.0,
             f"from={ax.get('from')} to={ax.get('to')}")

        # -- 异常兜底 --------------------------------------------------
        try:
            sess.series_style("NoSuchGraph", 1, color="red")
            step("series_style 缺失图时不抛 NoneType 异常", False, "未抛异常")
        except RuntimeError as exc:
            step("series_style 缺失图时抛可读 RuntimeError",
                 "找不到图形窗口或图层" in str(exc), str(exc)[:60])
        except AttributeError as exc:
            step("series_style 缺失图时抛可读 RuntimeError", False,
                 f"仍是 AttributeError: {exc}")

        step("graph_frame 可用", sess.graph_frame(g, boxed=True))

        # -- 输出 ------------------------------------------------------
        png = sess.graph_export(str(OUT / "session_test.png"), graph=g)
        step("导出 PNG", png.exists() and png.stat().st_size > 1000,
             f"{png.stat().st_size} bytes")

        opju = OUT / "session_test.opju"
        p = sess.project_save(str(opju))
        size1 = p.stat().st_size
        step("保存工程", p.exists() and size1 > 100_000, f"{size1} bytes")

        # -- 空工程覆盖保护 --------------------------------------------
        if size1 > 100_000:
            sess.ex("newbook;")   # 不置空，仅确认正常重保存
            sess.project_save(str(opju))
            step("正常重复保存后体积未异常缩水",
                 opju.stat().st_size >= size1 * 0.5,
                 f"{opju.stat().st_size} bytes")
    except Exception as exc:  # noqa: BLE001
        step(f"异常: {type(exc).__name__}: {exc}", False)
    finally:
        sess.close()

    print("\n" + "=" * 56)
    bad = [n for n, ok, _ in RESULTS if not ok]
    print(f"通过 {len(RESULTS) - len(bad)}/{len(RESULTS)}")
    for n in bad:
        print("  失败:", n)
    return 0 if not bad else 1


if __name__ == "__main__":
    raise SystemExit(main())
