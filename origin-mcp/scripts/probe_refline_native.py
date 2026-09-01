# -*- coding: utf-8 -*-
"""
probe_refline_native.py — 实测 Origin 原生参考线对象 layer.axis.refline# 在 COM 下是否可用。

官方文档（Origin 2018 SR1+）：
    layer.x.reflines.count = 2;      # 数量
    layer.x.refline1.value = 4;      # 在 x=4 加参考线
    layer.x.reflines.lineshow = 1;   # 显示
    layer.x.refline1.linecolor = color(blue);   # 颜色
    layer.x.refline1.linestyle = 2;  # 线型（虚线）
    layer.x.refline1.linethickness = 1.5;
    layer.x.refline1.labeltext$ = "...";
    layer.x.reflines.dataset$ = "2 4 7";   # 批量

若这些属性赋值在 COM Execute 下有效，就能把 add_ref_line 从「两点数据图」换成原生 refline：
不进图例、不触发轴缩放、不污染数据表。
"""

from __future__ import annotations

import time

import pythoncom
import win32com.client


def ev(app, expr):
    try:
        v = app.Evaluate(expr)
        return v
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

    hr = "=" * 66
    print(f"\n{hr}\n=== 1. 单条 X 轴参考线（x=5） ===\n{hr}")
    app.Execute(f"win -a {g};")
    cmds = [
        "layer.x.reflines.count = 1;",
        "layer.x.refline1.value = 5;",
        "layer.x.reflines.lineshow = 1;",
    ]
    for c in cmds:
        r = app.Execute(c)
        print(f"  {c:32} Execute={r}")
    time.sleep(0.4)
    for p in ("count", "refline1.value", "reflines.lineshow",
              "refline1.linecolor", "refline1.linestyle",
              "refline1.linethickness"):
        print(f"  读 {p:22} = {ev(app, f'layer.x.{p}')}")

    print(f"\n{hr}\n=== 2. 颜色 / 虚线 / 线宽 / 标签 ===\n{hr}")
    app.Execute(f"win -a {g};")
    cmds2 = [
        "layer.x.refline1.linecolor = 2;",      # 2=red 索引
        "layer.x.refline1.linestyle = 2;",      # 虚线？
        "layer.x.refline1.linethickness = 1.5;",
        'layer.x.refline1.labeltext$ = "yield";',
        "layer.x.refline1.labelshow = 1;",
    ]
    for c in cmds2:
        r = app.Execute(c)
        print(f"  {c:42} Execute={r}")
    time.sleep(0.4)
    for p in ("refline1.linecolor", "refline1.linestyle",
              "refline1.linethickness", "refline1.labelshow",
              "refline1.labeltext$"):
        print(f"  读 {p:22} = {ev(app, f'layer.x.{p}')}")

    print(f"\n{hr}\n=== 3. Y 轴参考线（y=30） ===\n{hr}")
    app.Execute(f"win -a {g};")
    cmds3 = [
        "layer.y.reflines.count = 1;",
        "layer.y.refline1.value = 30;",
        "layer.y.reflines.lineshow = 1;",
        "layer.y.refline1.linecolor = 4;",  # 4=blue
    ]
    for c in cmds3:
        r = app.Execute(c)
        print(f"  {c:32} Execute={r}")
    time.sleep(0.4)
    print(f"  读 layer.y.refline1.value = {ev(app, 'layer.y.refline1.value')}")

    print(f"\n{hr}\n=== 4. 加参考线后轴范围 / 图例是否被改 ===\n{hr}")
    print(f"  轴 x=[{gl.GetNumProp('x.from')},{gl.GetNumProp('x.to')}] "
          f"y=[{gl.GetNumProp('y.from')},{gl.GetNumProp('y.to')}]")
    print(f"  DataPlots.Count = {gl.DataPlots.Count}（不变=不进图例）")

    print(f"\n{hr}\n=== 5. dataset$ 批量 ===\n{hr}")
    app.Execute(f"win -a {g};")
    app.Execute('layer.x.reflines.dataset$ = "2 8";')
    time.sleep(0.3)
    print(f"  count = {ev(app, 'layer.x.reflines.count')}")
    print(f"  dataset$ = {ev(app, 'layer.x.reflines.dataset$')}")

    print(f"\n{hr}\n完成（Origin 保持打开，可目检：x=5 红线虚线 + y=30 蓝线）\n{hr}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
