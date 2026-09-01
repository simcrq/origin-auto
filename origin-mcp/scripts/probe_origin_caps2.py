# -*- coding: utf-8 -*-
"""
probe_origin_caps2.py — 第二轮能力探测（承接 probe_origin_caps.py）。

钉死以下细节：
  E. 窗口名净化：非字母数字删除 + 截断长度到底是多少
  F. PutWorksheet 在 wks.nrows 不足时是否自动扩行（分块回退的前提）
  G. 参考线：追加数据图后轴是否会重新缩放，能否用 rescale=0 锁回
  H. 图例文本读回与 LegendsAutoUpdate
  I. NANUM 哨兵复现条件（工作簿活动时 layer.* 读到的是什么）
  J. GraphLayer.GetNumProp 支持哪些轴属性
  K. （已移除）重命名为【已存在】窗口名会触发 Origin 模态对话框并永久挂死 COM，
     结论见 origin_mcp_server.py 顶部第 9 条。
"""

from __future__ import annotations

import time

import pythoncom
import win32com.client


def hr(t: str) -> None:
    print("\n" + "=" * 66 + f"\n{t}\n" + "=" * 66, flush=True)


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
    print("version =", app.Evaluate("@V"), flush=True)
    app.NewProject()
    time.sleep(0.5)

    # ------------------------------------------------------------------ E
    hr("E. 窗口名净化：截断长度")
    for want in ("AAAAAAAAAAAAAAAA", "BBBBBBBBBBBB", "CCCCCCCCCCCCC",
                 "DDDDDDDDDDDDDDDDDDDD", "E1E2E3E4E5E6E7E8",
                 "Graph_0.6_Force", "测试窗口名称很长很长"):
        app.Execute("newbook;")
        time.sleep(0.15)
        app.Execute(f"win -r %H {want};")
        time.sleep(0.25)
        try:
            actual = str(app.ActivePage.Name)
        except Exception as exc:  # noqa: BLE001
            actual = f"EXC {exc}"
        print(f"  {want!r:24}(len={len(want):2}) -> {actual!r:16}"
              f"（长度 {len(actual)}）", flush=True)

    # ------------------------------------------------------------------ F
    hr("F. PutWorksheet 是否自动扩行 / 扩列")
    app.Execute("newbook;")
    time.sleep(0.2)
    app.Execute("win -r %H ExpandBook;")
    time.sleep(0.3)
    eb = str(app.ActivePage.Name)
    app.Execute(f"win -a {eb};")
    print("  初始 nrows =", app.Evaluate("wks.nrows"),
          " ncols =", app.Evaluate("wks.ncols"), flush=True)
    data = [[float(r * 10 + c) for c in range(3)] for r in range(120)]
    ok = app.PutWorksheet(eb, [tuple(r) for r in data], 0, 0)
    print("  PutWorksheet(120x3, 未预设尺寸) =", ok, flush=True)
    print("  之后 nrows =", app.Evaluate("wks.nrows"),
          " ncols =", app.Evaluate("wks.ncols"), flush=True)
    back = app.GetWorksheet(eb)
    print("  回读行数 =", len(back) if back else None,
          " 末行 =", list(back[-1]) if back else None, flush=True)

    app.Execute("newbook;")
    time.sleep(0.2)
    app.Execute("win -r %H ChunkBook2;")
    time.sleep(0.3)
    cb = str(app.ActivePage.Name)
    app.Execute(f"win -a {cb};")
    ok_all = True
    for off in range(0, 300, 100):
        blk = [[float(off + r * 10 + c) for c in range(3)] for r in range(100)]
        ok_all = bool(app.PutWorksheet(cb, [tuple(r) for r in blk], off, 0)) \
            and ok_all
    print("  分块(起始行 0/100/200) 全成功 =", ok_all,
          " nrows =", app.Evaluate("wks.nrows"), flush=True)
    back = app.GetWorksheet(cb)
    if back:
        print(f"  回读 {len(back)} 行; row0={list(back[0])}, "
              f"row100={list(back[100])}, row200={list(back[200])}", flush=True)

    # ------------------------------------------------------------------ G/H
    hr("G/H. 参考线 + 轴锁定 + 图例")
    app.Execute("newbook;")
    time.sleep(0.2)
    app.Execute("win -r %H RefBook;")
    time.sleep(0.3)
    rb = str(app.ActivePage.Name)
    xs = [float(i) for i in range(11)]
    ys = [5.0 * i for i in range(11)]
    app.Execute(f"win -a {rb};")
    app.PutWorksheet(rb, [tuple(v) for v in zip(xs, ys)], 0, 0)
    time.sleep(0.3)
    app.Execute(f"win -a {rb};")
    app.Execute("plotxy iy:=(1,2) plot:=200 ogl:=[<new>];")
    time.sleep(1.0)
    app.Execute("win -r %H RefGraph;")
    time.sleep(0.5)
    rg = str(app.ActivePage.Name)
    gl = app.FindGraphLayer(rg)
    print(f"  图名={rg!r}", flush=True)
    y0, y1 = gl.GetNumProp("y.from"), gl.GetNumProp("y.to")
    x0, x1 = gl.GetNumProp("x.from"), gl.GetNumProp("x.to")
    print(f"  绘后轴范围: x=[{x0},{x1}] y=[{y0},{y1}]", flush=True)

    app.Execute(f"win -a {rb};")
    app.Execute("wks.ncols = 4; wks.nrows = 2;")
    app.PutWorksheet(rb, [(5.0, y0), (5.0, y1)], 0, 2)
    time.sleep(0.2)
    app.Execute(f"win -a {rb};")
    app.Execute(f"plotxy iy:=(3,4) plot:=200 ogl:=[{rg}]1;")
    time.sleep(1.0)
    y0b, y1b = gl.GetNumProp("y.from"), gl.GetNumProp("y.to")
    x0b, x1b = gl.GetNumProp("x.from"), gl.GetNumProp("x.to")
    print(f"  追加参考线后: x=[{x0b},{x1b}] y=[{y0b},{y1b}]", flush=True)
    print(f"  --> y 轴被改动: {(y0b, y1b) != (y0, y1)}", flush=True)
    print("  DataPlots 数 =", gl.DataPlots.Count, flush=True)

    for i in range(gl.GraphObjects.Count):
        try:
            o = gl.GraphObjects.Item(i)
        except Exception:  # noqa: BLE001
            continue
        if o.Name == "Legend":
            print("  图例文本 =", repr(str(o.Text)), flush=True)
    try:
        gp = None
        for i in range(app.GraphPages.Count):
            pg = app.GraphPages.Item(i)
            if str(pg.Name) == rg:
                gp = pg
                break
        if gp is not None:
            print("  LegendsAutoUpdate(前) =", gp.LegendsAutoUpdate, flush=True)
            gp.LegendsAutoUpdate = 0
            print("  LegendsAutoUpdate(后) =", gp.LegendsAutoUpdate, flush=True)
    except Exception as exc:  # noqa: BLE001
        print("  GraphPage 访问 EXC", exc, flush=True)

    app.Execute(f"win -a {rg};")
    app.Execute(f"layer.y.from = {y0}; layer.y.to = {y1}; layer.y.rescale = 0;")
    app.Execute(f"layer.x.from = {x0}; layer.x.to = {x1}; layer.x.rescale = 0;")
    time.sleep(0.5)
    print(f"  锁轴后: x=[{gl.GetNumProp('x.from')},{gl.GetNumProp('x.to')}] "
          f"y=[{gl.GetNumProp('y.from')},{gl.GetNumProp('y.to')}]", flush=True)

    # ------------------------------------------------------------------ I
    hr("I. 工作簿活动时 layer.* 读到什么")
    app.Execute(f"win -a {rb};")
    time.sleep(0.3)
    for expr in ("layer.x.from", "layer.x.to", "layer.y.from", "layer.y.to"):
        try:
            v = app.Evaluate(expr)
        except Exception as exc:  # noqa: BLE001
            v = f"EXC {exc}"
        print(f"  {expr:16} = {v}", flush=True)
    app.Execute(f"win -a {rg};")
    time.sleep(0.3)
    print("  -- 切回图窗口 --", flush=True)
    for expr in ("layer.x.from", "layer.y.from"):
        try:
            v = app.Evaluate(expr)
        except Exception as exc:  # noqa: BLE001
            v = f"EXC {exc}"
        print(f"  {expr:16} = {v}", flush=True)

    # ------------------------------------------------------------------ J
    hr("J. GraphLayer.GetNumProp 轴属性名")
    for prop in ("x.from", "x.to", "x.inc", "x.type", "x.rescale",
                 "y.from", "y.to", "y.inc", "y.type", "showframe"):
        try:
            print(f"  {prop:12} = {gl.GetNumProp(prop)}", flush=True)
        except Exception as exc:  # noqa: BLE001
            print(f"  {prop:12} EXC {exc}", flush=True)

    # 注意：不再做"重命名为已存在窗口名"的实测——已确认这会触发 Origin 的模态
    # 重名对话框并永久挂死 COM（180s 无返回），见 origin_mcp_server.py 顶部
    # 结论 9 与 _unique_name() 的净化后去重实现。

    hr("完成")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
