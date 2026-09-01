# -*- coding: utf-8 -*-
"""
probe_origin_caps.py — Origin COM 能力探测脚本（只读诊断 + 少量临时窗口）。

目的：在动手改 origin_mcp_server.py 之前，用真实 Origin 实测把几个"拍脑袋"
的结论钉死：
  A. 窗口名净化规则（Data_0.6 -> ?）
  B. 图层坐标轴范围到底怎么读（layer.x.from 哨兵 / [G]1! 语法 / GetNumProp）
  C. PutWorksheet 的 start_row 分块是否可行、大表耗时
  D. add_ref_line 场景下从【图形】窗口读 layer.y.from/to 是否可用

运行（需要 pywin32）：
    python scripts/probe_origin_caps.py [--hidden]

注意：脚本会调用 NewProject 清理现场，不要在打开重要工程时运行。
"""

from __future__ import annotations

import argparse
import time

import pythoncom
import win32com.client


def hr(title: str) -> None:
    print("\n" + "=" * 66)
    print(title)
    print("=" * 66)


def connect(visible: bool = True):
    pythoncom.CoInitialize()
    app = win32com.client.gencache.EnsureDispatch("Origin.ApplicationSI")
    app.Visible = visible
    deadline = time.time() + 60
    while time.time() < deadline:
        try:
            if app.Evaluate("1"):
                break
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.5)
    return app


def ev(app, expr):
    """带异常保护的求值，返回 (结果, 是否哨兵 NANUM)。"""
    try:
        v = app.Evaluate(expr)
    except Exception as exc:  # noqa: BLE001
        return f"EXC {exc}", True
    try:
        f = float(v)
    except (TypeError, ValueError):
        return repr(v), False
    return f, abs(f + 1.23456789e-300) < 1e-290 or abs(f) > 1e299


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--hidden", action="store_true", help="后台启动 Origin")
    ap.add_argument("--skip-big", action="store_true", help="跳过大数据量写入测试")
    args = ap.parse_args()

    app = connect(visible=not args.hidden)
    hr("0. 版本")
    print("version =", app.Evaluate("@V"))

    app.NewProject()
    time.sleep(0.5)

    # ---------------------------------------------------------------- A
    hr("A. 窗口名净化规则")
    for want in ("Data_0.6", "Data_0.9", "Force_Displacement_0.6",
                 "My Book", "A.B.C", "Book-1"):
        app.Execute("newbook;")
        time.sleep(0.2)
        app.Execute(f"win -r %H {want};")
        time.sleep(0.3)
        actual = ""
        try:
            actual = str(app.ActivePage.Name)
        except Exception as exc:  # noqa: BLE001
            actual = f"EXC {exc}"
        print(f"  请求 {want!r:28} -> 实际 {actual!r}")

    # ---------------------------------------------------------------- B
    hr("B. 坐标轴范围读取（关键）")
    app.Execute("newbook;")
    time.sleep(0.2)
    app.Execute("win -r %H ProbeBook;")
    time.sleep(0.3)
    book = str(app.ActivePage.Name)
    rows = [[i * 0.5, float(i) ** 1.5, 2.0 * i + 1] for i in range(1, 21)]
    app.Execute("wks.ncols = 3; wks.nrows = 20;")
    app.PutWorksheet(book, [tuple(r) for r in rows], 0, 0)
    time.sleep(0.3)

    app.Execute(f"win -a {book};")
    app.Execute("plotxy iy:=(1,2:3) plot:=202 ogl:=[<new>];")
    time.sleep(1.0)
    app.Execute("win -r %H ProbeGraph;")
    time.sleep(0.5)
    gname = str(app.ActivePage.Name)
    print(f"  图窗口实际名 = {gname!r}")

    print("\n  -- 未激活任何窗口（活动窗口=图） --")
    for expr in ("layer.x.from", "layer.x.to", "layer.y.from", "layer.y.to",
                 "layer.x.rescale"):
        v, bad = ev(app, expr)
        print(f"    {expr:20} = {v}   {'<-- NANUM 哨兵' if bad else ''}")

    print("\n  -- 先 win -a 图 再读 --")
    app.Execute(f"win -a {gname};")
    time.sleep(0.3)
    for expr in ("layer.x.from", "layer.x.to", "layer.y.from", "layer.y.to",
                 "layer.x.rescale", "layer.y.rescale"):
        v, bad = ev(app, expr)
        print(f"    {expr:20} = {v}   {'<-- NANUM 哨兵' if bad else ''}")

    print("\n  -- 工作簿活动（模拟 add_ref_line 的错误姿势） --")
    app.Execute(f"win -a {book};")
    time.sleep(0.3)
    for expr in ("layer.y.from", "layer.y.to"):
        v, bad = ev(app, expr)
        print(f"    {expr:20} = {v}   {'<-- NANUM 哨兵' if bad else ''}")

    print("\n  -- 跨窗口语法（不切换活动窗口） --")
    app.Execute(f"win -a {book};")
    time.sleep(0.3)
    for expr in (f"[{gname}]1!layer.x.from", f"{gname}!layer.x.from",
                 f"[{gname}]layer.x.from"):
        v, bad = ev(app, expr)
        print(f"    {expr:28} = {v}   {'<-- NANUM/失败' if bad else ''}")

    print("\n  -- COM GraphLayer 通道 --")
    try:
        gl = app.FindGraphLayer(gname)
        print("    FindGraphLayer ->", "OK" if gl is not None else "None")
        if gl is not None:
            for prop in ("x.from", "x.to", "y.from", "y.to", "X.From"):
                try:
                    v = gl.GetNumProp(prop)
                    print(f"    gl.GetNumProp({prop!r:10}) = {v}")
                except Exception as exc:  # noqa: BLE001
                    print(f"    gl.GetNumProp({prop!r:10}) EXC {exc}")
    except Exception as exc:  # noqa: BLE001
        print("    FindGraphLayer EXC", exc)

    print("\n  -- 设置 axis 后 rescale 行为 --")
    app.Execute(f"win -a {gname}; layer.x.from = 2; layer.x.to = 8;")
    time.sleep(0.5)
    for expr in ("layer.x.from", "layer.x.to"):
        v, bad = ev(app, expr)
        print(f"    {expr:20} = {v}   {'<-- NANUM' if bad else ''}")
    app.Execute("layer.x.rescale = 0;")
    time.sleep(0.3)
    for expr in ("layer.x.from", "layer.x.to", "layer.x.rescale"):
        v, bad = ev(app, expr)
        print(f"    rescale=0 后 {expr:20} = {v}   {'<-- NANUM' if bad else ''}")

    # ---------------------------------------------------------------- C
    hr("C. PutWorksheet 分块 / 大表")
    if args.skip_big:
        print("  (跳过)")
    else:
        app.Execute("newbook;")
        time.sleep(0.3)
        app.Execute("win -r %H ChunkBook;")
        time.sleep(0.3)
        cb = str(app.ActivePage.Name)
        nrows, ncols = 5000, 4
        app.Execute(f"wks.ncols = {ncols}; wks.nrows = {nrows};")
        big = [[float(r * ncols + c) for c in range(ncols)] for r in range(nrows)]
        rect = [tuple(r) for r in big]
        t0 = time.time()
        ok_all = True
        step = 1000
        off = 0
        while off < nrows:
            blk = rect[off:off + step]
            ok_all = bool(app.PutWorksheet(cb, blk, off, 0)) and ok_all
            off += step
        t1 = time.time()
        print(f"  分块 {nrows}x{ncols} step={step}: ok={ok_all} "
              f"耗时 {t1 - t0:.2f}s")
        back = app.GetWorksheet(cb)
        if back:
            print(f"  回读 {len(back)} 行; row0={list(back[0])}, "
                  f"row{step}={list(back[step])}, "
                  f"row{nrows-1}={list(back[nrows - 1])}")

        nrows2, ncols2 = 68804, 20
        app.Execute("newbook;")
        time.sleep(0.3)
        app.Execute("win -r %H BigBook;")
        time.sleep(0.3)
        bb = str(app.ActivePage.Name)
        app.Execute(f"wks.ncols = {ncols2}; wks.nrows = {nrows2};")
        big2 = [[float(r)] * ncols2 for r in range(nrows2)]
        rect2 = [tuple(r) for r in big2]
        t0 = time.time()
        ok1 = bool(app.PutWorksheet(bb, rect2, 0, 0))
        t1 = time.time()
        print(f"  整块 {nrows2}x{ncols2}: ok={ok1} 耗时 {t1 - t0:.2f}s")

        app.Execute("newbook;")
        time.sleep(0.3)
        app.Execute("win -r %H BigBook2;")
        time.sleep(0.3)
        bb2 = str(app.ActivePage.Name)
        app.Execute(f"wks.ncols = {ncols2}; wks.nrows = {nrows2};")
        t0 = time.time()
        step = 20000
        off, ok2 = 0, True
        while off < nrows2:
            ok2 = bool(app.PutWorksheet(bb2, rect2[off:off + step],
                                        off, 0)) and ok2
            off += step
        t1 = time.time()
        print(f"  分块 {nrows2}x{ncols2} step={step}: ok={ok2} "
              f"耗时 {t1 - t0:.2f}s")
        back2 = app.GetWorksheet(bb2)
        if back2:
            print(f"  回读 {len(back2)} 行; last={list(back2[-1])[:3]}...")

    # ---------------------------------------------------------------- D
    hr("D. add_ref_line 场景：从图形窗口读 y 轴范围")
    app.Execute(f"win -a {gname};")
    time.sleep(0.5)
    for expr in ("layer.y.from", "layer.y.to"):
        v, bad = ev(app, expr)
        print(f"  {expr:20} = {v}   {'<-- NANUM' if bad else ''}")

    hr("完成")
    print("提示：脚本故意不退出 Origin，便于人工检查。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
