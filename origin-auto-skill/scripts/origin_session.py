# -*- coding: utf-8 -*-
"""
origin_session.py — 可独立复用的 Origin COM 会话封装（Skill 回退脚本用）。

与 origin-mcp 服务端同源，浓缩了全部实测教训（Origin 2021 / 9.8002）：
- 专用 COM 线程；ApplicationSI 单实例连接
- newbook 后 win -r %H 重命名并【回读实际名】——Origin 会净化窗口名
  （删除非 ASCII 字母数字并截断到 13 字符：Data_0.6 -> Data06、
   Force_Displacement_0.6 -> ForceDisplace）
- 唯一化必须在【净化后】的名字上做，否则重名会触发模态框并永久挂死 COM
- 大表写 PutWorksheet 分块（start_row 递增），绝不逐格
- 读轴范围用 COM GraphLayer.GetNumProp("x.from"/...)，不用 LabTalk layer.*
- axis_set 给 vmin/vmax 时必须 rescale=0 锁轴，否则后续重绘会把范围改回去
- expGraph 只用裸文件名导出到用户文件夹(UFF)再复制
- app.Save 必须反斜杠绝对路径
"""

from __future__ import annotations

import os
import queue
import re
import shutil
import threading
import time
import concurrent.futures
from datetime import datetime
from pathlib import Path

PLOT_TYPE_IDS = {
    "line": 200, "scatter": 201, "symbol": 201,
    "line_symbol": 202, "linesymbol": 202,
    "column": 203, "bar": 204,
}
TEMPLATE_NAMES = {"line": "Line", "scatter": "Scatter",
                  "line_symbol": "LineSymb", "column": "Column"}
# 实测确认: y=1, x=4, yerr=3（选区绘图时误差列自动关联）。xerr=5/label=7 按文档。
COL_TYPE_CODES = {"none": 0, "y": 1, "yerr": 3, "x": 4, "xerr": 5, "label": 7}
COLOR_INDEX = {"black": 1, "red": 2, "green": 3, "blue": 4, "cyan": 5,
               "magenta": 6, "yellow": 7, "orange": 8, "navy": 14, "violet": 16}
EXPORT_TYPES = {"png": "png", "pdf": "pdf", "eps": "eps", "tif": "tif",
                "tiff": "tif", "emf": "emf", "jpg": "jpg", "jpeg": "jpg"}

# 窗口名净化规则（实测）
PAGE_NAME_MAXLEN = 13
# 单次 PutWorksheet 行数上限；超过则分块
PUT_CHUNK_ROWS = 20000
# 逐格写入的单元格上限，超过则直接报错（68804x20=137 万格逐格写不可接受）
CELL_FALLBACK_MAX_CELLS = 20000


def _clean(s) -> str:
    return str(s).replace('"', "'").replace("\n", " ").strip()


def _bsl(p) -> str:
    return str(Path(p).resolve())


def _col_letter(i1: int) -> str:
    s = ""
    while i1 > 0:
        i1, r = divmod(i1 - 1, 26)
        s = chr(65 + r) + s
    return s


def sanitize_page_name(name) -> str:
    """按 Origin 实测规则预测窗口短名：删非 ASCII 字母数字 + 截断 13 字符。"""
    s = re.sub(r"[^0-9A-Za-z]", "", str(name))
    return s[:PAGE_NAME_MAXLEN] or "A"


# 兼容别名
_sanitize_page_name = sanitize_page_name


def _name_key(name) -> str:
    """归一化键：忽略大小写与非字母数字。Data_0.6 与 Data06 同键。"""
    return re.sub(r"[^0-9a-z]", "", str(name).lower())


class ComWorker:
    def __init__(self):
        self._q = queue.Queue()
        self._t = None
        self._ready = threading.Event()

    def start(self):
        if self._t and self._t.is_alive():
            return
        self._t = threading.Thread(target=self._loop, daemon=True)
        self._t.start()
        self._ready.wait(15)

    def _loop(self):
        import pythoncom
        pythoncom.CoInitialize()
        self._ready.set()
        while True:
            fn, fut = self._q.get()
            if fn is None:
                break
            try:
                fut.set_result(fn())
            except BaseException as exc:
                fut.set_exception(exc)
        pythoncom.CoUninitialize()

    def submit(self, fn, timeout=60.0):
        self.start()
        fut = concurrent.futures.Future()
        self._q.put((fn, fut))
        return fut.result(timeout=timeout)

    def stop(self):
        if self._t and self._t.is_alive():
            self._q.put((None, None))


class OriginSession:
    def __init__(self):
        self.worker = ComWorker()
        self.app = None
        self._visible = True

    # -- 基础 ---------------------------------------------------------------

    def connect(self, visible=True):
        self._visible = bool(visible)

        def _go():
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            last = None
            app = None
            for f in (
                lambda: win32com.client.gencache.EnsureDispatch("Origin.ApplicationSI"),
                lambda: win32com.client.Dispatch("Origin.ApplicationSI"),
                lambda: win32com.client.Dispatch("Origin.Application"),
            ):
                try:
                    app = f()
                    break
                except Exception as exc:  # noqa: BLE001
                    last = exc
            if app is None:
                raise RuntimeError(f"COM 连接失败: {last}")
            try:
                app.Visible = bool(visible)
            except Exception:  # noqa: BLE001
                pass
            deadline = time.time() + 40
            while time.time() < deadline:
                try:
                    if app.Evaluate("1"):
                        break
                except Exception:  # noqa: BLE001
                    pass
                time.sleep(1)
            return app
        self.app = self.worker.submit(_go, timeout=90)
        return True

    def ex(self, script: str) -> bool:
        def _r():
            return bool(self.app.Execute(script))
        return self.worker.submit(_r, timeout=60)

    def evaluate(self, expr: str):
        def _r():
            return self.app.Evaluate(_clean(expr))
        return self.worker.submit(_r, timeout=60)

    def close(self):
        self.app = None
        self.worker.stop()

    # -- 页面 ---------------------------------------------------------------

    def all_pages(self):
        def _p():
            out = []
            coll = self.app.Pages
            for i in range(0, int(coll.Count) + 1):  # Item() 从 0 开始
                try:
                    nm = str(coll.Item(i).Name)
                except Exception:  # noqa: BLE001
                    continue
                if nm and nm != "None":
                    try:
                        t = str(coll.Item(i).TypeName)
                    except Exception:  # noqa: BLE001
                        t = "?"
                    out.append((nm, t))
            return out
        return self.worker.submit(_p, timeout=30)

    def active_page_name(self):
        def _p():
            try:
                return str(self.app.ActivePage.Name)
            except Exception:  # noqa: BLE001
                return ""
        return self.worker.submit(_p, timeout=30)

    def _all_pages_names(self):
        return [n for n, _ in self.all_pages()]

    def _unique_name(self, base: str) -> str:
        """在【净化后】的名字上唯一化。重名会触发 Origin 模态框并永久挂死 COM。"""
        s = _sanitize_page_name(base)
        taken = {_name_key(n) for n in self._all_pages_names()}
        if _name_key(s) not in taken:
            return s
        for k in range(2, 1000):
            suffix = str(k)
            cand = _sanitize_page_name(s[:PAGE_NAME_MAXLEN - len(suffix)] + suffix)
            if _name_key(cand) not in taken:
                return cand
        return _sanitize_page_name(s[:9] + str(int(time.time()) % 10000))

    def _resolve_page(self, name: str, kind: str | None = None) -> str:
        """把用户给的名字（可能含被净化的字符）映射到实际窗口名。"""
        nm = _clean(name)
        pages = self.all_pages()
        if not pages:
            return nm
        by_name = dict(pages)
        if nm in by_name:
            return nm
        if kind:
            want = kind.lower()
            pages = [(n, t) for n, t in pages if want in t.lower()]
        key = _name_key(nm)
        for n, _ in pages:
            if _name_key(n) == key:
                return n
        return nm

    def _rename_active(self, name: str) -> tuple[str, bool]:
        self.ex(f"win -r %H {name};")
        time.sleep(0.3)
        return self._actual_after_new(name)

    def _actual_after_new(self, expected: str) -> tuple[str, bool]:
        time.sleep(0.2)
        actual = ""
        for _ in range(5):
            actual = self.active_page_name()
            if actual and _name_key(actual) == _name_key(expected):
                return actual, True
            time.sleep(0.25)
        return actual, False

    # -- 数据 ---------------------------------------------------------------

    def project_new(self):
        """新建空白工程（会丢弃当前未保存内容）。"""

        def _r():
            return bool(self.app.NewProject())
        return self.worker.submit(_r, timeout=60)

    def workbook_new(self, name="Book"):
        base = _sanitize_page_name(_clean(name) if name else "Book")
        nm = self._unique_name(base)
        if not self.ex("newbook;"):
            raise RuntimeError("newbook 失败")
        # 关键：Origin 会净化窗口名（Data_0.6 -> Data06），必须回读实际名
        actual, _ = self._rename_active(nm)
        if not actual:
            raise RuntimeError(f"重命名后未能识别窗口（请求 {nm}）")
        return actual

    def _put_once(self, bk, rect, r1, c1, timeout=300):
        def _p():
            return bool(self.app.PutWorksheet(bk, rect, int(r1), int(c1)))
        try:
            return bool(self.worker.submit(_p, timeout=timeout))
        except Exception as exc:  # noqa: BLE001
            # 会话假死（长写入后实测偶发）：重连一次即可恢复
            self.connect(visible=self._visible)
            return False

    def _put_block(self, bk, rect, r1, c1):
        if self._put_once(bk, rect, r1, c1):
            return True
        self.ex(f"win -a {_clean(bk)};")
        return self._put_once(bk, rect, r1, c1)

    def _put_chunked(self, bk, rect, r1, c1):
        total = len(rect)
        if total <= PUT_CHUNK_ROWS:
            return False
        off = 0
        while off < total:
            block = rect[off:off + PUT_CHUNK_ROWS]
            if not self._put_once(bk, block, r1 + off, c1):
                return False
            off += PUT_CHUNK_ROWS
        return True

    def data_put(self, book, rows, header=None, start_row=1, start_col=1):
        if not rows:
            return 0, 0
        nrows = len(rows)
        ncols = max(len(r) for r in rows)
        bk = self._resolve_page(book, "worksheet")
        need_cols = start_col - 1 + ncols
        need_rows = start_row - 1 + nrows
        pre = [f"win -a {_clean(bk)};",
               f"if(wks.ncols < {need_cols}) wks.ncols = {need_cols};",
               # 只扩不缩：无脑赋值 wks.nrows 会截掉已有数据（实测教训）
               f"if(wks.nrows < {max(need_rows, 1)}) wks.nrows = {max(need_rows, 1)};"]
        if header:
            for j, h in enumerate(header):
                i = start_col + j
                L = _col_letter(i)
                pre.append(f"range hh{L} = {i};")
                pre.append(f'hh{L}[L]$ = "{_clean(h)}";')
        self.ex("\n".join(pre))
        rect = [tuple(r) + ("",) * (ncols - len(r)) for r in rows]
        r1, c1 = start_row - 1, start_col - 1

        if self._put_block(bk, rect, r1, c1):
            return nrows, ncols
        if self._put_chunked(bk, rect, r1, c1):
            return nrows, ncols
        cells = nrows * ncols
        if cells > CELL_FALLBACK_MAX_CELLS:
            raise RuntimeError(
                f"PutWorksheet 整块与分块均失败，且数据量 {nrows}x{ncols}="
                f"{cells} 格超过逐格兜底上限，已放弃")
        # 小表逐格兜底
        body = []
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                L = _col_letter(start_col + ci)
                ri_1 = start_row + ri
                if val is None or (isinstance(val, str) and not val):
                    continue
                if isinstance(val, str):
                    body.append(f'col({L})[{ri_1}]$ = "{_clean(val)}";')
                else:
                    body.append(f"col({L})[{ri_1}] = {val};")
        for k in range(0, len(body), 300):
            if not self.ex("\n".join(body[k:k + 300])):
                raise RuntimeError(f"第 {k}~{k+300} 单元格写入失败")
        return nrows, ncols

    # -- 绘图 ---------------------------------------------------------------

    @staticmethod
    def _runs(ys):
        ys = sorted(set(int(v) for v in ys))
        runs, cur = [], [ys[0]]
        for v in ys[1:]:
            if v == cur[-1] + 1:
                cur.append(v)
            else:
                runs.append(cur)
                cur = [v]
        runs.append(cur)
        return runs

    @staticmethod
    def _iy(x_col, cols):
        # 多 Y 必须写 (x,a:b)：内层再套括号会被 Origin 解析成 Y-vs-Y（实测坑）
        if len(cols) == 1:
            return f"({x_col},{cols[0]})"
        return f"({x_col},{cols[0]}:{cols[-1]})"

    def plot_create(self, book, x_col, y_cols, plot_type="line_symbol",
                    template=None, graph_name=None, pairs=None):
        """pairs=[[x1,y1],[x2,y2],...] 时按显式 XY 对绘制（各 X 独立）。
        返回【回读到的实际】图形窗口名列表（Origin 会净化名字）。"""
        pid_key = plot_type if isinstance(plot_type, int) else \
            str(plot_type).lower().strip()
        pid = PLOT_TYPE_IDS.get(pid_key, pid_key if isinstance(pid_key, int) else None)
        if pid is None:
            raise ValueError(f"未知 plot_type: {plot_type}")
        tpl = template or TEMPLATE_NAMES.get(
            pid_key if isinstance(pid_key, str) else "")
        clause = f" template:={_clean(tpl)}" if tpl else ""
        base = _sanitize_page_name(_clean(graph_name) if graph_name else "Graph")
        bk = self._resolve_page(book, "worksheet")

        if pairs is not None:
            iy = "(" + ",".join(f"({int(a)},{int(b)})" for a, b in pairs) + ")"
            gname = self._unique_name(base)
            self.ex("\n".join([f"win -a {_clean(bk)};",
                               f"plotxy iy:={iy} plot:={pid} ogl:=[<new{clause}>];",
                               f"win -r %H {gname};"]))
            time.sleep(1.0)
            actual, _ = self._actual_after_new(gname)
            exists = _name_key(actual) in {
                _name_key(n) for n, t in self.all_pages() if "graph" in t.lower()}
            if not exists:
                raise RuntimeError(f"绘图失败（请求 {gname}，实际 {actual or '无'}）")
            return [actual]

        graphs = []
        for k, run in enumerate(self._runs(y_cols)):
            gname = self._unique_name(base if k == 0 else f"{base}_p{k+1}")
            cmds = [f"win -a {_clean(bk)};",
                    f"plotxy iy:={self._iy(x_col, run)} plot:={pid} "
                    f"ogl:=[<new{clause}>];",
                    f"win -r %H {gname};"]
            self.ex("\n".join(cmds))
            time.sleep(1.0)
            actual, _ = self._actual_after_new(gname)
            exists = _name_key(actual) in {
                _name_key(n) for n, t in self.all_pages() if "graph" in t.lower()}
            if exists:
                graphs.append(actual)
        if not graphs:
            raise RuntimeError("绘图失败：检查列号与数据有效性")
        return graphs

    def _find_graph_layer(self, graph):
        g = self._resolve_page(graph, "graph")

        def _f():
            return self.app.FindGraphLayer(g)
        gl = self.worker.submit(_f, timeout=30)
        if gl is None:
            self.ex(f"win -a {g};")
            time.sleep(0.4)
            gl = self.worker.submit(_f, timeout=30)
        return gl

    def axis_get(self, graph, axis="x"):
        """读轴范围：走 COM GraphLayer.GetNumProp，不依赖活动窗口。

        返回 {"ok": bool, "from":.., "to":.., "increment":.., "scale":..}，
        任一项读失败则 ok=False 并带 invalid_fields。"""
        gl = self._find_graph_layer(graph)
        if gl is None:
            return {"ok": False, "graph": graph,
                    "error": f"找不到图形窗口或图层：{graph}"}

        def _g():
            out = {}
            for p in ("from", "to", "inc", "type"):
                try:
                    out[p] = float(gl.GetNumProp(f"{axis.lower()}.{p}"))
                except Exception:  # noqa: BLE001
                    out[p] = None
            return out
        vals = self.worker.submit(_g, timeout=30)
        bad = [k for k, v in vals.items() if v is None]
        return {"ok": not bad, "graph": graph, "axis": axis.lower(),
                "from": vals.get("from"), "to": vals.get("to"),
                "increment": vals.get("inc"), "scale": vals.get("type"),
                "invalid_fields": bad or None}

    def axis_set(self, graph, axis="x", title=None, vmin=None, vmax=None,
                 scale=None, major_inc=None, rescale=None):
        ax = axis.lower()
        assert ax in ("x", "y")
        g = self._resolve_page(graph, "graph")
        lines = [f"win -a {g};"]
        if title is not None:
            lines.append(f'{"xb" if ax == "x" else "yl"}.text$ = '
                         f'"{_clean(title)}";')
        if vmin is not None:
            lines.append(f"layer.{ax}.from = {float(vmin)};")
        if vmax is not None:
            lines.append(f"layer.{ax}.to = {float(vmax)};")
        if scale:
            sc = {"linear": 1, "log10": 2, "log": 2}[scale.lower()]
            lines.append(f"layer.{ax}.type = {sc};")
        if major_inc is not None:
            lines.append(f"layer.{ax}.inc = {float(major_inc)};")
        lock = rescale if rescale is not None else (vmin is not None
                                                    or vmax is not None)
        if lock:
            # 关键：不锁轴的话后续重绘会把手动范围改回去（出现 -2 之类自动边距）
            lines.append(f"layer.{ax}.rescale = 0;")
        return self.ex("\n".join(lines))

    def series_style(self, graph, series_index=1, color=None, line_width_pt=None,
                     symbol_size=None, symbol_shape=None, y_col=None):
        """实测唯一可靠通道：COM DataPlot.SetNumProp（set 命令静默无效）。
        color 用 Origin 索引名/数字；line_width_pt 单位 pt。
        带误差棒时用 y_col（Y 列号）定位曲线，series_index 会错位。"""
        g = self._resolve_page(graph, "graph")

        def _s():
            gl = self.app.FindGraphLayer(g)
            if gl is None:
                raise RuntimeError(f"找不到图形窗口或图层：{graph}（解析为 {g}）")
            if y_col is not None:
                target = None
                n = int(gl.DataPlots.Count)
                for i in range(n):
                    try:
                        dp_i = gl.DataPlots.Item(i)
                        nm = str(dp_i.GetDatasetName())
                    except Exception:  # noqa: BLE001
                        continue
                    if nm.upper().endswith("_" + _col_letter(int(y_col))):
                        target = dp_i
                        break
                if target is None:
                    raise RuntimeError(f"未找到 y_col={y_col} 对应的数据图")
                dp = target
            else:
                dp = gl.DataPlots.Item(int(series_index) - 1)
            if color is not None:
                c = str(color).lower()
                ci = COLOR_INDEX.get(c, int(c) if str(c).isdigit() else 1)
                dp.SetNumProp("color", float(ci))
            if line_width_pt is not None:
                dp.SetNumProp("line.width", float(line_width_pt))
            if symbol_size is not None:
                dp.SetNumProp("symbol.size", float(symbol_size))
            if symbol_shape is not None:
                dp.SetNumProp("symbol.shape", float(symbol_shape))
            return True
        return self.worker.submit(_s, timeout=30)

    def graph_frame(self, graph, boxed=True):
        """显示/隐藏图层边框（上/右轴线）。唯一有效通道：SetNumProp('showframe')。"""
        g = self._resolve_page(graph, "graph")

        def _f():
            gl = self.app.FindGraphLayer(g)
            if gl is None:
                raise RuntimeError(f"找不到图形窗口或图层：{graph}（解析为 {g}）")
            gl.SetNumProp("showframe", 1.0 if boxed else 0.0)
            return float(gl.GetNumProp("showframe"))
        return bool(self.worker.submit(_f, timeout=30))

    # -- 输出 ---------------------------------------------------------------

    def uff(self):
        def _g():
            return str(self.app.LTStr("%Y")).strip().rstrip("\\/")
        return self.worker.submit(_g, timeout=30)

    def graph_export(self, filepath, fmt=None, graph=None):
        p = Path(_bsl(filepath))
        p.parent.mkdir(parents=True, exist_ok=True)
        ftype = EXPORT_TYPES.get((fmt or p.suffix.lower().lstrip(".")))
        if not ftype:
            raise ValueError(f"不支持的导出格式: {fmt or p.suffix}")
        pre = f"win -a {_clean(self._resolve_page(graph, 'graph'))};" if graph else ""
        bare = f"skill_{datetime.now().strftime('%H%M%S%f')}.{ftype}"
        if not self.ex(pre + f'expGraph type:={ftype} filename:="{bare}";'):
            raise RuntimeError("expGraph 执行失败（检查目标图是否有效）")
        src = Path(self.uff()) / bare
        for _ in range(10):
            if src.exists():
                break
            time.sleep(1)
        if not src.exists():
            raise RuntimeError("expGraph 未生成文件（静默失败）")
        shutil.copyfile(src, p)
        try:
            src.unlink()
        except OSError:
            pass
        return p

    def project_save(self, path, force=False, backup=True):
        p = Path(_bsl(path))
        p.parent.mkdir(parents=True, exist_ok=True)
        old_size = p.stat().st_size if p.exists() else 0

        # 安全护栏（对应实测踩坑：会话假死后保存会用空工程覆盖已有成果）。
        # 当前工程枚举到 0 个页面、目标文件已存在且明显非空时，默认拒绝覆盖。
        if old_size > 100_000 and not force:
            pages = self.all_pages()
            if not pages:
                raise RuntimeError(
                    f"当前工程枚举到 0 个页面，但目标文件已存在且为 {old_size} 字节。"
                    "拒绝用空工程覆盖（多半是 COM 会话假死）。确认无误后传 force=True")

        bak = None
        if backup and old_size > 0:
            bak = p.with_name(p.name + ".bak")
            shutil.copy2(p, bak)

        def _s():
            return bool(self.app.Save(str(p)))
        if self.worker.submit(_s, timeout=90) and p.exists():
            new_size = p.stat().st_size
            if bak is not None:
                if old_size and new_size < old_size * 0.5:
                    # 体积异常缩水 = 空工程覆盖，保留备份
                    raise RuntimeError(
                        f"保存后体积从 {old_size} 缩到 {new_size} 字节，疑似空工程覆盖！"
                        f"原文件已备份为 {bak}")
                try:
                    bak.unlink()
                except OSError:
                    pass
            return p
        raise RuntimeError("保存失败（请在 Origin 中手动 Ctrl+S）")

    def read_worksheet(self, book):
        def _g():
            raw = self.app.GetWorksheet(_clean(book))
            return [list(r) for r in raw] if raw else []
        return self.worker.submit(_g, timeout=120)
