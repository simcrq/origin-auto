# -*- coding: utf-8 -*-
"""
test_fixes_e2e.py — 针对 6 个实测问题的回归测试（真实 Origin，端到端）。

对应问题编号：
  #1 窗口名净化不一致      #2 data_put 回退慢
  #3 绘图验证 / series_style #4 轴状态回读
  #5 参考线高度 / 图例      #6 空工程覆盖保护

用法：
    python scripts/test_fixes_e2e.py
结束码 0 = 全部通过。
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from origin_mcp_server import session  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "test_output"
OUT.mkdir(exist_ok=True)

RESULTS: list[tuple[str, bool, str]] = []


def step(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, bool(ok), detail))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))


def main() -> int:
    print("=== Origin MCP 修复回归测试 ===")

    # ------------------------------------------------------------- #1
    print("\n--- #1 窗口名净化 ---")
    r = session.connect(visible=True)
    step("连接 Origin", r.get("ok"), f"version={r.get('version_labtalk')}")
    session.project_new()

    r1 = session.workbook_new(name="Data_0.6")
    b1 = r1.get("book")
    step("workbook_new('Data_0.6') 返回实际净化名",
         r1.get("ok") and b1 == "Data06", f"book={b1} warning={r1.get('warning')}")

    t0 = time.time()
    r1b = session.workbook_new(name="Data_0.6")
    dt = time.time() - t0
    step("重名不挂死（旧实现会阻塞 180~366s）",
         r1b.get("ok") and dt < 20, f"第二次 book={r1b.get('book')} 耗时 {dt:.1f}s")

    r2 = session.workbook_new(name="Force_Displacement_0.6")
    step("workbook_new 长名截断到 13 字符",
         r2.get("book") == "ForceDisplace", f"book={r2.get('book')}")

    # 用【未净化】的名字调用，应该能自动解析到实际窗口
    r3 = session.page_activate("Data_0.6")
    step("用未净化名 Data_0.6 激活（自动解析）",
         r3.get("ok"), f"activated={r3.get('activated')}")
    r4 = session.data_put("Data_0.6", [[1, 2], [3, 4]])
    step("用未净化名 data_put（自动解析）", r4.get("ok"),
         f"method={r4.get('method')}")

    # ------------------------------------------------------------- #2
    print("\n--- #2 大表写入 ---")
    book = r1.get("book", "Data06")
    rows = [[float(i * 20 + c) for c in range(20)] for i in range(68804)]
    t0 = time.time()
    rp = session.data_put(book, rows)
    dt = time.time() - t0
    step("68804x20 写入", rp.get("ok"),
         f"method={rp.get('method')} 耗时 {dt:.2f}s")
    step("大表未走逐格兜底", rp.get("method") != "cell-by-cell",
         f"method={rp.get('method')}")

    rchk = session.worksheet_get_data(book)
    if rchk.get("ok") and rchk.get("data"):
        d = rchk["data"]
        ok_data = (len(d) == 68804 and float(d[0][0]) == 0.0
                   and float(d[68803][19]) == 68803 * 20 + 19)
        step("写入数据回读一致（首/末单元格）", ok_data,
             f"{len(d)} 行, 首={d[0][0]}, 末={d[68803][19]}")
    else:
        step("写入数据回读一致（首/末单元格）", False, str(rchk.get("error")))

    # ------------------------------------------------------------- #3
    print("\n--- #3 绘图验证 / series_style ---")
    # 注意：Force_Displacement_0.6 的净化名 ForceDisplace 已被上面的工作簿占用，
    # 因此这里用另一个名字验证净化；跨类型重名避让另设一项断言。
    pairs = [[1, 2], [3, 4], [5, 6]]
    t0 = time.time()
    rplot = session.plot_create(book, pairs=pairs, graph_name="FD_0.6")
    g = rplot.get("graph")
    step("plot_create(pairs) 成功且返回实际净化名",
         rplot.get("ok") and g == "FD06",
         f"graph={g} curves={rplot.get('curves_detected')} "
         f"耗时 {time.time() - t0:.1f}s")
    # 工作簿与图形共用窗口名空间，重名必须自动避让而不是挂死
    t0 = time.time()
    rcoll = session.plot_create(book, pairs=[[9, 10]],
                                graph_name="Force_Displacement_0.6")
    dt = time.time() - t0
    step("图形名与已有工作簿冲突时自动避让且不挂死",
         rcoll.get("ok") and rcoll.get("graph") != "ForceDisplace" and dt < 20,
         f"graph={rcoll.get('graph')} 耗时 {dt:.1f}s")
    step("曲线数与 pairs 数一致",
         rplot.get("curves_detected") == len(pairs),
         f"curves={rplot.get('curves_detected')} 期望={len(pairs)}")

    rbad = session.series_style("NoSuchGraph", color="red")
    step("series_style 对不存在的图不再抛 NoneType 异常",
         rbad.get("ok") is False and "找不到" in str(rbad.get("error")),
         f"error={rbad.get('error')}")
    rok = session.series_style(g, y_col=2, color="red", line_width_pt=1.5)
    step("series_style(y_col=) 正常图可用", rok.get("ok"), f"applied={rok.get('applied')}")

    # ------------------------------------------------------------- #4
    print("\n--- #4 轴状态回读 ---")
    rax = session.axis_set(g, "y", vmin=0, vmax=25)
    act = rax.get("actual") or {}
    step("axis_set(vmin/vmax) 回读校验通过",
         rax.get("verified") and act.get("from") == 0.0 and act.get("to") == 25.0,
         f"actual={act} locked={rax.get('locked')}")
    step("axis_set 默认自动锁轴(rescale=0)", rax.get("locked") is True)

    rget = session.axis_get(g, "y")
    step("axis_get 走 COM 通道读回", rget.get("ok")
         and rget.get("from") == 0.0 and rget.get("to") == 25.0,
         f"from={rget.get('from')} to={rget.get('to')} inc={rget.get('increment')}")

    # 活动窗口是工作簿时，裸 LabTalk 读到的是工作表图层的值（实测 0/1/8 之类）。
    # 用 y.to 对比：图上是 25，工作簿图层读到的是完全无关的值。
    session.page_activate(book)
    raw = session.labtalk_evaluate("layer.y.to")
    withwin = session.labtalk_evaluate("layer.y.to", window=g)
    step("labtalk_evaluate 裸调用 vs 指定 window",
         withwin.get("ok") and withwin.get("value") == 25.0
         and raw.get("value") != 25.0,
         f"裸(工作簿活动)={raw.get('value')} 指定window={withwin.get('value')}")

    # 追加曲线后轴是否仍保持锁定（#4 的 -2 边距问题）
    session.page_activate(book)
    session.plot_create(book, pairs=[[7, 8]], graph_name="Probe2")
    rget2 = session.axis_get(g, "y")
    step("其它操作后轴范围未被自动缩放改回",
         rget2.get("from") == 0.0 and rget2.get("to") == 25.0,
         f"from={rget2.get('from')} to={rget2.get('to')}")

    # ------------------------------------------------------------- #5
    print("\n--- #5 参考线（原生 refline） ---")
    before = session.axis_get(g, "y")
    rref = session.add_ref_line(graph=g, axis="x", pos=10.0, dash=True)
    step("add_ref_line 走原生 refline 且成功",
         rref.get("ok") and rref.get("native") == "refline",
         f"index={rref.get('refline_index')} pos={rref.get('pos')}")

    v1 = session.labtalk_evaluate("layer.x.refline1.value", window=g)
    step("参考线位置精确落在 x=10",
         v1.get("ok") and v1.get("value") == 10.0,
         f"value={v1.get('value')}")

    after = session.axis_get(g, "y")
    step("原生 refline 不触发轴缩放（范围不变）",
         after.get("from") == before.get("from")
         and after.get("to") == before.get("to"),
         f"before={before.get('from')}~{before.get('to')} "
         f"after={after.get('from')}~{after.get('to')}")

    # 同轴追加第二条，确认不覆盖第一条
    rref2 = session.add_ref_line(graph=g, axis="x", pos=12.0, dash=True)
    v1b = session.labtalk_evaluate("layer.x.refline1.value", window=g)
    v2b = session.labtalk_evaluate("layer.x.refline2.value", window=g)
    step("同轴追加不覆盖前一条",
         rref2.get("refline_index") == 2 and v1b.get("value") == 10.0
         and v2b.get("value") == 12.0,
         f"idx={rref2.get('refline_index')} v1={v1b.get('value')} "
         f"v2={v2b.get('value')}")

    # Y 轴独立参考线
    rref3 = session.add_ref_line(graph=g, axis="y", pos=20.0, color=4)
    vy1 = session.labtalk_evaluate("layer.y.refline1.value", window=g)
    step("Y 轴参考线独立生效",
         rref3.get("ok") and vy1.get("value") == 20.0,
         f"y refline1={vy1.get('value')}")

    rleg = session.legend_remove_last(g, count=1)
    step("legend_remove_last 工具仍可用", rleg.get("ok"),
         f"legend={str(rleg.get('legend'))[:50]}")

    # ------------------------------------------------------------- #6
    print("\n--- #6 工程保护 ---")
    opju = OUT / "fixes_e2e.opju"
    rsave = session.project_save(str(opju))
    step("正常保存工程", rsave.get("ok"),
         f"size={rsave.get('size')} pages={rsave.get('pages')}")
    size1 = opju.stat().st_size if opju.exists() else 0

    session.graph_export(str(OUT / "fixes_e2e.png"), graph=g)
    step("导出 PNG 真实生成", (OUT / "fixes_e2e.png").exists(),
         f"{(OUT / 'fixes_e2e.png').stat().st_size if (OUT / 'fixes_e2e.png').exists() else 0} bytes")

    # 模拟会话假死后的空工程覆盖
    if size1 > 100_000:
        session.project_new()
        rguard = session.project_save(str(opju))
        step("空工程覆盖被拦截", rguard.get("ok") is False,
             f"error={str(rguard.get('error'))[:70]}")
        step("拦截后原文件未被破坏",
             opju.exists() and opju.stat().st_size == size1,
             f"{opju.stat().st_size if opju.exists() else 0} bytes")
        rforce = session.project_save(str(opju), force=True)
        step("force=True 可强制覆盖（并留 .bak）",
             rforce.get("ok") and Path(str(opju) + ".bak").exists(),
             f"warning={str(rforce.get('warning'))[:60]}")
    else:
        step("空工程覆盖被拦截", False, f"首个工程仅 {size1} 字节，无法构造场景")

    # -------------------------------------------------------------
    print("\n" + "=" * 60)
    bad = [n for n, ok, _ in RESULTS if not ok]
    print(f"通过 {len(RESULTS) - len(bad)}/{len(RESULTS)}")
    if bad:
        print("失败项：")
        for n in bad:
            print("  -", n)
    print("输出目录:", OUT)
    return 0 if not bad else 1


if __name__ == "__main__":
    raise SystemExit(main())
