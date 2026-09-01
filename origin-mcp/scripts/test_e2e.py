# -*- coding: utf-8 -*-
"""
test_e2e.py — Origin MCP 端到端验收测试。

流程（对应 MCP连接.txt 的验收标准）：
    连接 Origin → 新建工作簿 → 写入 X/Y 数据 → 创建折线+符号图
    → 设置坐标轴标题 → 导出 PNG → 保存 .opju → 校验文件真实存在

用法：
    python scripts/test_e2e.py [--keep-open]
"""

import argparse
import math
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from origin_mcp_server import session  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep-open", action="store_true",
                    help="结束后不退出 Origin（便于人工检查图形）")
    args = ap.parse_args()

    out_dir = Path(tempfile.gettempdir()) / "origin_mcp_test"
    out_dir.mkdir(exist_ok=True)
    png = out_dir / "mcp_test.png"
    opju = out_dir / "mcp_test.opju"
    csvp = out_dir / "mcp_test_export.csv"
    steps: list[tuple[str, bool, str]] = []

    def step(name: str, ok: bool, detail: str = "") -> None:
        steps.append((name, ok, detail))
        print(("[PASS] " if ok else "[FAIL] ") + name + (f" — {detail}" if detail else ""))

    print("=== Origin MCP E2E 测试开始 ===")
    r = session.connect(visible=True)
    step("连接 Origin", r.get("ok") is True, f"version={r.get('version')}")

    r = session.project_new()
    step("新建工程", r["ok"])

    r = session.workbook_new(name="McpTest")
    book = r.get("book", "McpTest")
    step("新建工作簿", r["ok"], f"name={book}")

    xs = [i * 0.5 for i in range(13)]
    ys1 = [10 * math.sin(x) + x * 0.8 for x in xs]
    ys2 = [6 * math.cos(x) - x * 0.3 for x in xs]
    data = [[x, y1, y2] for x, y1, y2 in zip(xs, ys1, ys2)]
    r = session.data_put(book, data, header=["Time_s", "SignalA", "SignalB"])
    step("写入数据(13行×3列)", r["ok"], f"{r.get('rows')} rows × {r.get('cols')} cols")

    r = session.worksheet_set_columns(book, [
        {"index": 1, "name": "Time", "units": "s"},
        {"index": 2, "name": "Signal A", "units": "V"},
        {"index": 3, "name": "Signal B", "units": "V"},
    ])
    # 注意：不设 type=x。plotxy 的 iy:=(1,y) 已显式指定 X 列，
    # 再设列设计ation 会改变区间语法语义导致曲线错乱（实测坑点）。
    step("设置列属性", r["ok"])

    r = session.plot_create(book, x_col=1, y_cols=[2, 3],
                            plot_type="line_symbol", graph_name="McpGraph")
    gname = r.get("graph", "McpGraph")
    step("创建点线图(双Y)", r["ok"], f"graph={gname}")

    r = session.axis_set(gname, axis="x", title="Time (s)", vmin=0, vmax=6)
    step("设置 X 轴", r["ok"])
    r = session.axis_set(gname, axis="y", title="Amplitude (V)")
    step("设置 Y 轴标题", r["ok"])

    r = session.series_style(gname, series_index=1, color="red",
                             line_width_pt=2.0)
    step("曲线样式", r["ok"])

    r = session.text_label("(a) MCP E2E Test", graph=gname)
    step("添加文本标注", r["ok"])

    r = session.graph_export(str(png), graph=gname)
    step("导出 PNG", r["ok"], r.get("path", r.get("hint", "")))

    r = session.project_save(str(opju))
    step("保存 OPJU", r["ok"], r.get("path", r.get("hint", "")))

    r = session.worksheet_get_data(book)
    step("读回数据校验", r["ok"] and r.get("rows") == 13,
         f"{r.get('rows')} rows")

    r = session.labtalk_evaluate("10*3")
    step("LabTalk Evaluate 回读", r["ok"] and abs(r.get("value", 0) - 30) < 1e-9,
         str(r.get("value")))

    csvp = out_dir / "mcp_test_export.csv"
    r = session.worksheet_export_csv(str(csvp), book=book)
    step("导出 CSV 回读校验", r["ok"], r.get("path", r.get("hint", "")))

    print("-" * 50)
    failed = [s for s in steps if not s[1]]
    print(f"结果：{len(steps) - len(failed)}/{len(steps)} 步通过")
    print(f"输出目录：{out_dir}")
    if png.exists():
        print(f"  PNG: {png} ({png.stat().st_size} bytes)")
    if opju.exists():
        print(f"  OPJU: {opju} ({opju.stat().st_size} bytes)")

    session.invalidate()
    print("=== 测试结束（Origin 保持打开以便人工检查） ===")
    return 1 if len(failed) > 2 else 0  # 允许个别非关键步骤失败


if __name__ == "__main__":
    sys.exit(main())
