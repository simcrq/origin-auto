# -*- coding: utf-8 -*-
"""
probe_ref_line.py — 实测 Origin 原生参考线（line 对象）是否可控。

背景：当前 add_ref_line 用「两点数据图 + ogl:=[G]1 追加」的 hack，因为早前
结论是 draw -l 坐标不可控。这里重新系统实测 draw -l（原生 annotation line 对象）：
若坐标可控，则原生 line 对象不进图例、不触发轴缩放、不污染数据表，更优。

测试点：
  A. draw -l 各语法的坐标是否可控（用读回验证）
  B. 颜色 / 虚线 / 宽度是否生效
  C. 画线后轴范围是否被改（是否触发重缩放）
  D. 是否进入图例（DataPlots.Count 是否增加）
  E. line 对象的 LabTalk 属性名（x1/y1/x2/y2/color/linetype 等）读回
"""

from __future__ import annotations

import time

import pythoncom
import win32com.client


def ev(app, expr):
    try:
        return app.Evaluate(expr)
    except Exception as exc:  # noqa: BLE001
        return f"EXC {exc}"


def main() -> int:
    pythoncom.CoInitialize()
    app = win32com.client.gencache.EnsureDispatch("Origin.ApplicationSI")
    app.Visible = True
    for _ in range(120):
        try:
            if app.Evaluate("1"):
                break
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.5)
    print("version =", app.Evaluate("@V"))
    app.NewProject()
    time.sleep(0.5)

    # 建数据 + 图
    app.Execute("newbook;")
    time.sleep(0.2)
    app.Execute("win -r %H RefBook;")
    time.sleep(0.3)
    book = str(app.ActivePage.Name)
    xs = [i * 1.0 for i in range(11)]
    ys = [5.0 * i for i in range(11)]
    app.Execute(f"win -a {book};")
    app.PutWorksheet(book, [tuple(v) for v in zip(xs, ys)], 0, 0)
    time.sleep(0.3)
    app.Execute(f"win -a {book};")
    app.Execute("plotxy iy:=(1,2) plot:=200 ogl:=[<new>];")
    time.sleep(1.0)
    app.Execute("win -r %H RefGraph;")
    time.sleep(0.5)
    g = str(app.ActivePage.Name)
    gl = app.FindGraphLayer(g)
    print(f"图名={g!r}")
    print(f"初始轴 x=[{gl.GetNumProp('x.from')},{gl.GetNumProp('x.to')}] "
          f"y=[{gl.GetNumProp('y.from')},{gl.GetNumProp('y.to')}]")
    print(f"初始 DataPlots.Count = {gl.DataPlots.Count}")

    # 目标：竖线 x=5，纵跨 y 轴 [-10,60]
    X = 5.0
    y0, y1 = gl.GetNumProp("y.from"), gl.GetNumProp("y.to")

    print("\n=== A. draw -l 各语法 ===")
    cases = [
        ("A1", f"draw -l -n RL1 -c 2 {X} {y0} {X} {y1};"),
        ("A2", f"draw -n RL2 -l -c 2 {X} {y0} {X} {y1};"),
        ("A3", f"draw -l -n RL3 -c 2 -w 1.5 -l 2 {X} {y0} {X} {y1};"),
        ("A4", f"draw -n RL4 -l -c 4 -w 2 {X} {y0} {X} {y1};"),
    ]
    for tag, cmd in cases:
        app.Execute(f"win -a {g};")
        time.sleep(0.2)
        r = app.Execute(cmd)
        time.sleep(0.3)
        # 读回坐标
        got = {p: ev(app, f"{tag}.{p}") for p in
               ("x1", "y1", "x2", "y2", "color", "linetype", "linewidth")}
        print(f"  {tag}: Execute={r} -> {got}")

    print("\n=== B. 读回 line 对象属性（LabTalk 对象名） ===")
    for tag in ("RL1", "RL2", "RL3", "RL4"):
        props = {}
        for p in ("x1", "y1", "x2", "y2", "color", "type"):
            props[p] = ev(app, f"{tag}.{p}")
        print(f"  {tag}: {props}")

    print("\n=== C. 画线后轴范围是否被改 ===")
    print(f"  现在 x=[{gl.GetNumProp('x.from')},{gl.GetNumProp('x.to')}] "
          f"y=[{gl.GetNumProp('y.from')},{gl.GetNumProp('y.to')}]")
    print(f"  DataPlots.Count = {gl.DataPlots.Count}（不变则不进图例）")

    print("\n=== D. GraphObjects 枚举（看 line 对象的真实坐标/类型） ===")
    cnt = int(gl.GraphObjects.Count)
    print(f"  GraphObjects.Count = {cnt}")
    for i in range(0, cnt + 1):
        try:
            o = gl.GraphObjects.Item(i)
            nm = str(o.Name)
        except Exception:  # noqa: BLE001
            continue
        if not nm or nm == "None":
            continue
        # 尝试读对象属性
        info = {}
        for p in ("Type", "type"):
            try:
                info[p] = o.GetNumProp(p)
            except Exception:  # noqa: BLE001
                pass
        print(f"  [{i}] {nm!r} {info}")

    print("\n=== E. 原生 line 对象能否设虚线/颜色（对象名赋值） ===")
    app.Execute(f"win -a {g};")
    for cmd in ("RL1.color = 4;", "RL1.linetype = 2;", "RL1.linewidth = 3;",
                "RL1.width = 3;"):
        r = app.Execute(cmd)
        print(f"  {cmd:24} Execute={r}")

    hr = "=" * 66
    print(f"\n{hr}\n完成（Origin 保持打开，便于人工目检 RL1~RL4 的位置与样式）\n{hr}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
