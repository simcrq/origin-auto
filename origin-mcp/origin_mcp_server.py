# -*- coding: utf-8 -*-
"""
origin_mcp_server.py — 通过 MCP 协议让 AI Agent 自动化操作 OriginLab Origin。

在 Origin 2021 (9.85) 实测得出的关键结论（务必遵守，勿凭文档想象）：
1. 所有 COM 调用固定在单一专用线程执行（CoInitialize 一次），避免跨套间崩溃。
2. 连接用 Origin.ApplicationSI；EnsureDispatch(gencache) 后才有 FindWorksheet/
   GetWorksheet/PutWorksheet/Evaluate 等方法。
3. workbook 命名：`newbook name:=X` 的 name 参数在 COM 下不生效！
   正确姿势：newbook; 之后立即 win -r %H X 重命名，再用页面集合验证。
4. expGraph 静默失败的真凶是【绝对路径】和【width/height/tr1.* 尺寸子句】——
   二者任选其一都会导致返回 True 却不生成文件。正确姿势：
   激活目标图后用裸文件名导出到用户文件夹(UFF)，再由 Python 复制到目标路径。
   实测可用格式：png/pdf/eps/tif/emf；svg 在 2021 COM 下不可用。
5. 项目保存：app.Save(反斜杠绝对路径) 可靠；正斜杠会静默失败。
6. 输出回读：Execute 只返回 bool，但 Evaluate(expr)/LTStr(name)/LTVar(name)
   可以读取 LabTalk 表达式与变量。
7. 数据读写优先走 COM：PutWorksheet(name, data, r1, c1) 批量写（r1/c1 从 0 起），
   GetWorksheet(name) 整表读回。

【第二轮实测补充（Origin 2021 / 9.8002，scripts/probe_origin_caps*.py 验证）】
8. 窗口名会被 Origin 净化：**删除所有非 ASCII 字母数字字符**（点/下划线/空格/
   连字符/中文全部被删）后**截断到 13 个字符**。Data_0.6 -> Data06、
   Force_Displacement_0.6 -> ForceDisplace、中文名 -> "A"。
   因此绝不能把用户传入的名字当作实际窗口名，必须重命名后回读。
9. `win -r %H X` 若 X 净化后与已有窗口重名，Origin 弹出模态框并**永久挂死 COM**
   （实测 180s 无返回，只能外部关闭对话框）。唯一化必须在【净化后】的名字上做。
10. 图层坐标轴范围的可靠读法不是 LabTalk `layer.x.from`（活动窗口是工作簿时会读到
    工作表的 layer，返回 0/4/8 这类垃圾值甚至 NANUM 哨兵），而是 COM：
    `GraphLayer.GetNumProp("x.from"/"x.to"/"y.from"/"y.to"/"x.inc"/"x.type")`
    ——纯 COM、不依赖活动窗口、大小写不敏感。跨窗口 LabTalk 语法
    `[Graph1]1!layer.x.from` 也可用，但优先用 COM。
11. PutWorksheet 会**自动扩行扩列**（向默认 32 行的表写 120x3 后自动变成 120x3），
    因此 wks.nrows 预检只是保险；分块写（start_row 递增）数据完全正确，
    68804x20 整块约 0.4s、分块约 0.5s，慢的只有逐格兜底路径。
12. 往图层追加数据图（如追加第二条曲线）会**触发轴重新缩放**：实测 y 从
    [-10,60] 变成 [-20,70]。因此追加后必须重新写回 layer.*.from/to 并
    layer.*.rescale=0 锁定。
13. 参考线应走 Origin 原生对象 `layer.x.refline#`（2018 SR1+，probe_refline_native.py
    实测）：不进图例、不触发轴缩放、不写数据表，位置/颜色/虚线/线宽/标签均可控。
    早前「两点数据图 + ogl:=[G]1」与 `draw -l`（坐标在 COM 下不可控）均已弃用。

运行：
    python origin_mcp_server.py          # stdio transport
"""

from __future__ import annotations

import logging
import os
import queue
import re
import shutil
import sys
import threading
import time
from datetime import datetime
import concurrent.futures
from pathlib import Path

try:
    import winreg
except ImportError:
    winreg = None

# ---------------------------------------------------------------------------
# 日志：MCP stdio 模式下只允许写 stderr
# ---------------------------------------------------------------------------
logging.basicConfig(stream=sys.stderr, level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("origin-mcp")

IS_WINDOWS = sys.platform.startswith("win")


# ---------------------------------------------------------------------------
# COM 工作线程
# ---------------------------------------------------------------------------
class ComWorker:
    """把所有 COM 调用串行化到单一专用线程，保证套间(apartment)一致性。"""

    def __init__(self) -> None:
        self._tasks: "queue.Queue[tuple]" = queue.Queue()
        self._thread: threading.Thread | None = None
        self._ready = threading.Event()

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._loop, daemon=True,
                                        name="origin-com-worker")
        self._thread.start()
        if not self._ready.wait(timeout=15):
            raise RuntimeError("COM 工作线程启动失败")

    def _loop(self) -> None:
        import pythoncom
        pythoncom.CoInitialize()
        self._ready.set()
        while True:
            fn, fut = self._tasks.get()
            if fn is None:
                break
            try:
                fut.set_result(fn())
            except BaseException as exc:  # noqa: BLE001
                fut.set_exception(exc)
        pythoncom.CoUninitialize()

    def stop(self) -> None:
        if self._thread and self._thread.is_alive():
            self._tasks.put((None, None))

    def submit(self, fn, timeout: float = 60.0):
        self.start()
        fut: concurrent.futures.Future = concurrent.futures.Future()
        self._tasks.put((fn, fut))
        return fut.result(timeout=timeout)


# ---------------------------------------------------------------------------
# 常量表
# ---------------------------------------------------------------------------
PLOT_TYPE_IDS: dict[str, int] = {
    # plotxy plot:= ID。若你的 Origin 版本表现不符，可直接传整数覆盖。
    "line": 200,
    "scatter": 201,
    "symbol": 201,
    "line_symbol": 202,
    "linesymbol": 202,
    "column": 203,
    "bar": 204,
}

TEMPLATE_NAMES: dict[str, str] = {
    # 本机(Origin 2021)实测可用的模板名
    "line": "Line",
    "scatter": "Scatter",
    "line_symbol": "LineSymb",
    "column": "Column",
}

COL_TYPE_CODES: dict[str, int] = {
    # 实测确认: y=1, x=4, yerr=3（选区绘图时误差列自动关联）。
    # xerr=5 / label=7 按文档推测未验证；2/6 作用未确认。
    "none": 0, "y": 1, "yerr": 3, "x": 4, "xerr": 5, "label": 7,
}

COLOR_INDEX: dict[str, int] = {
    # Origin 颜色索引（渲染实测）：14=深蓝 16=紫；灰色索引未确认，勿用
    "black": 1, "red": 2, "green": 3, "blue": 4, "cyan": 5,
    "magenta": 6, "yellow": 7, "orange": 8,
    "navy": 14, "violet": 16,
}

EXPORT_TYPES: dict[str, str] = {
    # 实测经裸文件名导出可用的格式
    "png": "png", "pdf": "pdf", "eps": "eps", "tif": "tif", "tiff": "tif",
    "emf": "emf", "jpg": "jpg", "jpeg": "jpg",
}


def _clean_text(s) -> str:
    """LabTalk 字符串里双引号易引发解析事故，统一替换掉。"""
    return str(s).replace('"', "'").replace("\n", " ").strip()


def _posix(p) -> str:
    return str(Path(p).resolve()).replace("\\", "/")


def _bsl(p) -> str:
    """COM Save/Load 需要反斜杠绝对路径。"""
    return str(Path(p).resolve())


def _col_letter(idx1: int) -> str:
    s, n = "", idx1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------------------
# 窗口名净化（Origin 2021 实测规则）
# ---------------------------------------------------------------------------
# Origin 短窗口名只保留 ASCII 字母数字，并截断到 13 字符：
#   Data_0.6                -> Data06
#   Force_Displacement_0.6  -> ForceDisplace
#   My Book                 -> MyBook
#   测试窗口名               -> A          （净化后为空时 Origin 回退为 "A"）
PAGE_NAME_MAXLEN = 13

# LabTalk 的 NANUM（缺失值）哨兵，约 -1.23456789e-300。
_NANUM = -1.23456789e-300


def _sanitize_page_name(name) -> str:
    """按 Origin 实测规则预测窗口短名：删非 ASCII 字母数字 + 截断 13 字符。

    预测值只用于去重/校验，真正使用的永远是回读到的实际窗口名。"""
    s = re.sub(r"[^0-9A-Za-z]", "", str(name))
    return s[:PAGE_NAME_MAXLEN] or "A"


def _name_key(name) -> str:
    """归一化键：忽略大小写与非字母数字。Data_0.6 与 Data06 同键。"""
    return re.sub(r"[^0-9a-z]", "", str(name).lower())


def _is_nanum(v) -> bool:
    """判断是否为 Origin 的 NANUM 哨兵值（读轴状态失败时的典型表现）。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return False
    if f != f:  # NaN
        return True
    if abs(f) > 1e299:  # 溢出/无穷
        return True
    # 必须按相对误差精确匹配哨兵：用绝对容差会把正常的 0.0 误判成哨兵
    # （实测活动窗口是工作簿时 layer.y.from 读到的正是 0.0）
    return f < 0 and abs(f - _NANUM) <= abs(_NANUM) * 1e-6


# 单次 PutWorksheet 的行数上限；超过则分块（实测整块 68804x20 约 0.4s，
# 但超大表的 COM 编组在部分机器上不稳定，分块更稳且同样快）。
PUT_CHUNK_ROWS = 20000
# 逐格写入的单元格上限。超过这个量级逐格写是不可接受的（68804x20=137万格）。
CELL_FALLBACK_MAX_CELLS = 20000


class OriginSession:
    """管理一个长驻的 Origin COM 连接。"""

    def __init__(self) -> None:
        self.worker = ComWorker()
        self.app = None
        self.connected = False
        self.visible = True
        self._uff_cache: str | None = None
        # 曾经成功枚举到过页面。用于区分"工程真的是空的"与"COM 代理假死"——
        # 后者需要重连恢复，前者绝不能触发重连风暴。
        self._ever_had_pages = False
        self._last_recovery: float = 0.0

    # -- 连接管理 -----------------------------------------------------------

    def invalidate(self, exc: Exception | None = None) -> None:
        self.app = None
        self.connected = False
        self._uff_cache = None
        if exc is not None:
            log.warning("Origin 连接失效: %s", exc)

    def _dispatch(self, visible: bool):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        last_err: Exception | None = None
        app = None
        for factory in (
            lambda: win32com.client.gencache.EnsureDispatch("Origin.ApplicationSI"),
            lambda: win32com.client.Dispatch("Origin.ApplicationSI"),
            lambda: win32com.client.Dispatch("Origin.Application"),
        ):
            try:
                app = factory()
                break
            except Exception as exc:  # noqa: BLE001
                last_err = exc
        if app is None:
            raise RuntimeError(
                f"无法通过 COM 连接 Origin：{last_err}。"
                "请确认已安装 Origin 并完成过首次启动/许可验证")
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
            time.sleep(1.0)
        return app

    def connect(self, visible: bool = True) -> dict:
        self.visible = visible
        app = self.worker.submit(lambda: self._dispatch(visible), timeout=90)
        self.app = app
        self.connected = True
        self._uff_cache = None
        # 关键：读版本也必须走 COM 工作线程（主线程未 CoInitialize，
        # 直调 app.Evaluate 会抛 com_error 导致 version 恒为空）
        version = ""
        try:
            version = str(self.worker.submit(lambda: app.Evaluate("@V"),
                                             timeout=30))
        except Exception:  # noqa: BLE001
            pass
        return {"ok": True, "connected": True, "version_labtalk": version}

    def ensure_connected(self) -> None:
        if not self.connected:
            self.connect(visible=self.visible)

    def ex(self, script: str) -> bool:
        def _run():
            if self.app is None:
                raise RuntimeError("Origin 未连接")
            return bool(self.app.Execute(script))
        return self.worker.submit(_run, timeout=60)

    def safe_op(self, fn, retry: bool = True):
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001
            msg = str(exc)
            # 业务性失败（对象不存在等）重连也没用，直接抛出
            if not retry or "未连接" in msg or "未找到" in msg:
                raise
            log.warning("COM 操作失败(%s)，重连后重试一次", exc)
            self.invalidate(exc)
            self.ensure_connected()
            return fn()

    # -- 内部工具 -----------------------------------------------------------

    def _user_files_folder(self) -> str:
        if self._uff_cache:
            return self._uff_cache

        def _get():
            v = self.app.LTStr("%Y")
            return str(v).strip().rstrip("\\/")
        self._uff_cache = self.worker.submit(_get, timeout=30)
        return self._uff_cache

    @staticmethod
    def _coll_names(coll) -> list[str]:
        names: list[str] = []
        try:
            n = coll.Count
        except Exception:  # noqa: BLE001
            return names
        for i in range(1, max(n, 0) + 2):  # +2: 集合尾部可能带一个空项
            try:
                p = coll.Item(i)
                nm = p.Name
            except Exception:  # noqa: BLE001
                continue
            if nm:
                names.append(str(nm))
        return list(dict.fromkeys(names))

    def _all_pages(self, recover: bool = True) -> list[tuple[str, str]]:
        """枚举所有页面，返回 [(名称, 类型)]。

        注意：WorksheetPages/GraphPages 集合的 Item 对象在部分状态下访问
        .Name 会失败，实测统一走 Pages 集合最可靠。

        recover=True 时，若枚举结果为空但本会话此前见过页面，判定为 COM 代理
        假死（长 PutWorksheet 后实测会出现），重连同一 Origin 实例(SI)后重试。
        工程内容不会因为重连丢失。
        """
        def _run():
            out: list[tuple[str, str]] = []
            coll = self.app.Pages
            cnt = int(coll.Count)
            # 注意：COM 集合 Item() 为 0 基索引（实测），Count 为元素个数
            for i in range(0, cnt + 1):
                try:
                    p = coll.Item(i)
                    nm = str(p.Name)
                except Exception:  # noqa: BLE001
                    continue
                if not nm or nm == "None":
                    continue
                tname = ""
                try:
                    tname = str(p.TypeName)
                except Exception:  # noqa: BLE001
                    pass
                out.append((nm, tname))
            return out

        try:
            pages = self.safe_op(lambda: self.worker.submit(_run, timeout=30)) or []
        except Exception:  # noqa: BLE001
            pages = []
        if pages:
            self._ever_had_pages = True
            return pages
        if not (recover and self.connected and self._ever_had_pages):
            return pages
        # 限流：10 秒内不重复做恢复尝试，避免空工程下的重连风暴
        if time.time() - self._last_recovery < 10:
            return pages
        self._last_recovery = time.time()
        log.warning("页面枚举为空（此前有内容），判定 COM 代理假死，重连后重试")
        self.invalidate()
        self.ensure_connected()
        try:
            pages = self.safe_op(lambda: self.worker.submit(_run, timeout=30)) or []
        except Exception:  # noqa: BLE001
            pages = []
        if pages:
            self._ever_had_pages = True
            log.warning("重连后恢复，共 %d 个页面", len(pages))
        return pages

    def _ws_names(self) -> list[str]:
        return [n for n, t in self._all_pages() if "worksheet" in t.lower()]

    def _graph_names(self) -> list[str]:
        return [n for n, t in self._all_pages() if "graph" in t.lower()]

    def _unique_name(self, base: str) -> str:
        """生成不与现有窗口冲突的名字。

        关键：唯一化必须在【净化后】的名字上做。Origin 会把 Data_0.6 净化成
        Data06，若按原始名去重，两次请求的名字净化后相同 -> Origin 弹出模态框
        并永久挂死 COM（实测 180s 无返回）。
        """
        s = _sanitize_page_name(base)
        taken = {_name_key(n) for n, _ in self._all_pages()}
        if _name_key(s) not in taken:
            return s
        # 追加序号时要给序号留位置，否则会被 13 字符截断后再次撞名
        for k in range(2, 1000):
            suffix = str(k)
            cand = _sanitize_page_name(s[:PAGE_NAME_MAXLEN - len(suffix)] + suffix)
            if _name_key(cand) not in taken:
                return cand
        return _sanitize_page_name(s[:9] + str(int(time.time()) % 10000))

    def _all_pages_names(self) -> list[str]:
        return [n for n, _ in self._all_pages()]

    def _resolve_page(self, name: str, kind: str | None = None) -> str:
        """把用户给的名字映射到 Origin 实际存在的窗口名。

        解决"传 Data_0.6、实际叫 Data06"导致 win -a 静默失败的问题：
        先精确匹配，再按净化键匹配（忽略大小写与非字母数字）。
        kind 可传 "worksheet"/"graph" 限定类型。
        枚举失败时原样返回，保持向后兼容。
        """
        nm = _clean_text(name)
        pages = self._all_pages()
        if not pages:
            return nm
        by_name = dict(pages)
        if nm in by_name:
            return nm
        if kind:
            want = kind.lower()
            pages = [(n, t) for n, t in pages if want in t.lower()]
        key = _name_key(nm)
        if not key:
            return nm
        for n, _ in pages:
            if _name_key(n) == key:
                return n
        return nm

    def _rename_active_page(self, name: str) -> tuple[str, bool]:
        """重命名当前活动窗口并回读 Origin 实际采用的名字。

        返回 (实际窗口名, 重命名是否命中预期)。
        Origin 会净化窗口名，因此必须以回读值为准——否则后续所有
        win -a / FindGraphLayer 都会静默失败。
        """
        self.ex(f"win -r %H {name};")
        time.sleep(0.3)
        return self._actual_after_new(name)

    def _actual_after_new(self, expected: str) -> tuple[str, bool]:
        """回读刚新建/重命名的窗口实际名（命令里已含 win -r 时用这个）。

        Origin 会净化名字，所以只按"净化键"比对：
        Data_0.6 与 Data06 视为命中，其它情况视为未命中（但名字照样返回）。
        """
        time.sleep(0.2)
        actual = ""
        for _ in range(5):
            actual = self._active_page_name()
            if actual and _name_key(actual) == _name_key(expected):
                return actual, True
            time.sleep(0.25)
        return actual, False

    def _find_graph_layer(self, graph: str):
        """按（可能未净化的）图形名取 GraphLayer，取不到返回 None 而不是抛异常。

        实测 FindGraphLayer 返回 None 时，紧接着访问 .DataPlots 会刷屏
        "'NoneType' object has no attribute 'DataPlots'"，必须在这里拦住。
        """
        g = self._resolve_page(graph, "graph")
        if g != graph:
            log.info("图形名 %s 解析为实际窗口 %s", graph, g)
        try:
            gl = self.safe_op(
                lambda: self.worker.submit(
                    lambda: self.app.FindGraphLayer(g), timeout=30))
        except Exception:  # noqa: BLE001
            return None
        if gl is None:
            # 激活一次再取：部分状态下图层尚未就绪
            self.ex(f"win -a {g};")
            time.sleep(0.4)
            try:
                gl = self.safe_op(
                    lambda: self.worker.submit(
                        lambda: self.app.FindGraphLayer(g), timeout=30))
            except Exception:  # noqa: BLE001
                return None
        return gl

    def _axis_prop(self, graph: str, axis: str, prop: str):
        """用 COM 读取坐标轴属性（唯一可靠通道）。

        实测：GraphLayer.GetNumProp("x.from"/"x.to"/"y.from"/"y.to"/"x.inc"/
        "x.type") 可用，且不依赖活动窗口；而 LabTalk layer.x.from 在活动窗口
        是工作簿时读到的是工作表图层，返回垃圾值或 NANUM 哨兵。
        """
        gl = self._find_graph_layer(graph)
        if gl is None:
            return None

        def _get():
            return float(gl.GetNumProp(f"{axis.lower()}.{prop}"))
        try:
            return self.safe_op(lambda: self.worker.submit(_get, timeout=30))
        except Exception:  # noqa: BLE001
            return None

    def _active_page_name(self) -> str:
        def _run():
            try:
                return str(self.app.ActivePage.Name)
            except Exception:  # noqa: BLE001
                return ""
        try:
            return self.safe_op(lambda: self.worker.submit(_run, timeout=30)) or ""
        except Exception:  # noqa: BLE001
            return ""

    # -- 项目级 -------------------------------------------------------------

    def project_new(self) -> dict:
        def _run():
            return bool(self.app.NewProject())
        ok = self.safe_op(lambda: self.worker.submit(_run, timeout=60))
        return {"ok": bool(ok), "action": "new_project"}

    def project_save(self, path: str, force: bool = False,
                     backup: bool = True) -> dict:
        """保存工程为 .opju/.opj。

        安全护栏（因实测踩过坑：会话假死后 pages_list 返回空，此时保存会用
        0.6KB 空工程覆盖掉 6.9MB 的已有成果）：
        1. 当前工程为空且目标文件已存在且体积明显不是空工程时，默认拒绝覆盖，
           除非显式 force=True；
        2. 覆盖前把已有文件备份为 <原名>.bak，若保存后体积正常（不小于原体积
           的一半）则删除备份，否则保留备份并在返回值里给出警告。
        """
        p = Path(_bsl(path))
        p.parent.mkdir(parents=True, exist_ok=True)
        old_size = p.stat().st_size if p.exists() else 0

        pages = self._all_pages()
        if not pages and old_size > 100_000 and not force:
            return {"ok": False, "path": str(p),
                    "error": (f"当前工程枚举到 0 个页面，但目标文件已存在且为 "
                              f"{old_size} 字节。拒绝用空工程覆盖"
                              f"（多半是 COM 会话假死）"),
                    "hint": ("先 origin_disconnect 再 origin_connect 重连，"
                             "确认 pages_list 能列出内容后重试；"
                             "确需强制覆盖请传 force=True")}

        bak: Path | None = None
        if backup and old_size > 0:
            bak = p.with_name(p.name + ".bak")
            try:
                shutil.copy2(p, bak)
            except OSError as exc:  # noqa: BLE001
                log.warning("备份 %s 失败: %s", bak, exc)
                bak = None

        def _save():
            return bool(self.app.Save(str(p)))

        attempts: list[str] = []
        saved = False
        try:
            ok = self.safe_op(lambda: self.worker.submit(_save, timeout=90))
            attempts.append(f"COM Save(backslash)={ok}")
            saved = bool(ok and p.exists())
        except Exception as exc:  # noqa: BLE001
            attempts.append(f"COM Save 异常: {exc}")
            self.invalidate(exc)
        if not saved:
            try:
                ok = self.ex(f'save "{p}";')
                attempts.append(f"LabTalk save={ok}")
                saved = bool(ok and p.exists())
            except Exception as exc:  # noqa: BLE001
                attempts.append(f"LabTalk save 异常: {exc}")
                self.invalidate(exc)
        if not saved:
            return {"ok": False, "path": str(p), "attempts": attempts,
                    "backup": str(bak) if bak else None,
                    "hint": "请在 Origin 中手动 Ctrl+S"}

        new_size = p.stat().st_size
        res: dict = {"ok": True, "path": str(p), "size": new_size,
                     "pages": len(pages)}
        # 体积异常缩水 = 典型的"空工程覆盖"，保留备份并告警
        if bak is not None:
            if old_size and new_size < old_size * 0.5:
                res["backup"] = str(bak)
                res["warning"] = (f"保存后体积从 {old_size} 字节缩到 {new_size} "
                                  f"字节，疑似空工程覆盖！原文件已备份为 {bak}")
            else:
                try:
                    bak.unlink()
                except OSError:
                    pass
        if not pages:
            res["warning"] = (res.get("warning", "") +
                              " 当前工程枚举到 0 个页面，请确认内容是否完整").strip()
        return res

    def project_open(self, path: str) -> dict:
        p = Path(_bsl(path))
        if not p.exists():
            return {"ok": False, "error": f"文件不存在: {path}"}

        def _load():
            return bool(self.app.Load(str(p)))
        ok = self.safe_op(lambda: self.worker.submit(_load, timeout=180))
        return {"ok": bool(ok), "path": str(p),
                "hint": None if ok else "Load 返回失败，确认文件版本兼容"}

    # -- 页面 ---------------------------------------------------------------

    def pages_list(self) -> dict:
        pages = self._all_pages()
        return {"ok": bool(pages),
                "worksheets": [n for n, t in pages if "worksheet" in t.lower()],
                "graphs": [n for n, t in pages if "graph" in t.lower()],
                "count": len(pages)}

    def page_activate(self, name: str) -> dict:
        nm = self._resolve_page(name)
        ok = self.ex(f"win -a {nm};")
        active = self._active_page_name()
        verified = _name_key(active) == _name_key(nm)
        res = {"ok": bool(ok and verified), "activated": nm,
               "verified": verified, "active_page": active}
        if nm != _clean_text(name):
            res["resolved_from"] = _clean_text(name)
        return res

    # -- 工作簿/工作表 ------------------------------------------------------

    def workbook_new(self, name: str | None = None, sheets: int = 1) -> dict:
        base = _sanitize_page_name(_clean_text(name) if name else "Book")
        nm = self._unique_name(base)
        if not self.ex("newbook;"):
            return {"ok": False, "error": "newbook 执行失败"}
        # 关键 1：name:= 参数在 COM 下不生效，必须创建后立即重命名
        # 关键 2：Origin 会净化窗口名（Data_0.6 -> Data06，且截断 13 字符），
        #         必须回读实际名，否则调用方拿到的是永远 win -a 不到的假名字
        actual, matched = self._rename_active_page(nm)
        if not actual:
            return {"ok": False, "error": f"重命名后未能识别窗口（请求 {nm}）",
                    "existing": self._all_pages_names()}
        if int(sheets) > 1:
            self.ex(f"newsheet name:=Sheet2 number:={int(sheets) - 1};")
        res: dict = {"ok": True, "book": actual}
        if actual != _clean_text(name):
            res["requested"] = _clean_text(name)
            res["warning"] = (f"Origin 净化了窗口名：{_clean_text(name)} -> "
                              f"{actual}（非字母数字被删除并截断到 "
                              f"{PAGE_NAME_MAXLEN} 字符）。后续调用请使用 {actual}")
        if not matched:
            res["warning"] = (f"重命名未完全生效：请求 {nm}，实际窗口名 {actual}。"
                              "后续调用请使用实际名")
        return res

    def worksheet_set_columns(self, book: str, columns: list[dict]) -> dict:
        bk = self._resolve_page(book, "worksheet")
        lines = [f"win -a {bk};"]
        applied: list[int] = []
        for c in columns:
            try:
                i = int(c.get("index", 0))
            except (TypeError, ValueError):
                continue
            if i <= 0:
                continue
            tag = _col_letter(i)
            lines.append(f"range rr{tag} = {i};")
            if c.get("name"):
                lines.append(f'rr{tag}[L]$ = "{_clean_text(c["name"])}";')
            if c.get("units"):
                lines.append(f'rr{tag}[U]$ = "{_clean_text(c["units"])}";')
            if c.get("comments"):
                lines.append(f'rr{tag}[C]$ = "{_clean_text(c["comments"])}";')
            if c.get("type"):
                t = COL_TYPE_CODES.get(str(c["type"]).lower())
                if t is not None:
                    lines.append(f"wks.col{i}.type = {t};")
            applied.append(i)
        ok = self.ex("\n".join(lines))
        return {"ok": bool(ok), "applied_columns": applied}

    def data_put(self, book: str, data: list[list], start_row: int = 1,
                 start_col: int = 1, header: list[str] | None = None) -> dict:
        if not data:
            return {"ok": False, "error": "data 为空"}
        rows = len(data)
        ncols = max(len(r) for r in data)
        # 关键：用户可能传的是未净化的名字（Data_0.6），解析成实际窗口（Data06）
        bk = self._resolve_page(book, "worksheet")
        need_cols = start_col - 1 + ncols
        need_rows = start_row - 1 + rows
        pre = [
            f"win -a {bk};",
            f"if(wks.ncols < {need_cols}) wks.ncols = {need_cols};",
            # 关键：只扩不缩——无脑赋值 wks.nrows 会截掉已有数据（实测教训）
            f"if(wks.nrows < {max(need_rows, 1)}) wks.nrows = {max(need_rows, 1)};",
        ]
        if header:
            for j, h in enumerate(header):
                i = start_col + j
                tag = _col_letter(i)
                pre.append(f"range hh{tag} = {i};")
                pre.append(f'hh{tag}[L]$ = "{_clean_text(h)}";')
        ok_pre = self.ex("\n".join(pre))

        # 主通道：COM PutWorksheet 批量写入（r1/c1 从 0 起）
        # 实测 PutWorksheet 会自动扩行扩列，所以 wks 预检只是保险不是前提
        rect = [tuple(r) + ("",) * (ncols - len(r)) for r in data]
        r1, c1 = int(start_row - 1), int(start_col - 1)

        # 1) 整块写入
        if self._put_block(bk, rect, r1, c1):
            return {"ok": True, "method": "PutWorksheet", "rows": rows,
                    "cols": ncols, "pre_ok": bool(ok_pre)}

        # 2) 分块写入：大表 COM 编组失败时的主力回退，实测数据完全正确
        if self._put_chunked(bk, rect, r1, c1):
            n_chunks = -(-rows // PUT_CHUNK_ROWS)
            return {"ok": True, "method": "PutWorksheet(chunked)", "rows": rows,
                    "cols": ncols, "pre_ok": bool(ok_pre),
                    "chunks": n_chunks}

        # 3) 逐格兜底：只对小表开放。68804x20=137 万格逐格写是不可接受的
        cells = rows * ncols
        if cells > CELL_FALLBACK_MAX_CELLS:
            return {"ok": False, "rows": rows, "cols": ncols,
                    "pre_ok": bool(ok_pre),
                    "error": (f"PutWorksheet 整块与分块均失败，且数据量 "
                              f"{rows}x{ncols}={cells} 单元格超过逐格兜底上限 "
                              f"{CELL_FALLBACK_MAX_CELLS}，已放弃写入以免长时间阻塞"),
                    "hint": ("多为 COM 会话假死：先 origin_disconnect 再 "
                             "origin_connect 重连，或新建工程后重试")}
        body: list[str] = []
        for ri, row in enumerate(data):
            for ci, val in enumerate(row):
                col_i = start_col + ci
                row_i = start_row + ri
                L = _col_letter(col_i)
                if val is None or (isinstance(val, str) and not val):
                    continue
                if isinstance(val, str):
                    body.append(f'col({L})[{row_i}]$ = "{_clean_text(val)}";')
                else:
                    body.append(f"col({L})[{row_i}] = {val};")
        for k in range(0, len(body), 300):
            if not self.ex("\n".join(body[k:k + 300])):
                return {"ok": False, "error": f"第 {k}~{k+300} 单元格写入失败",
                        "pre_ok": bool(ok_pre)}
        return {"ok": True, "method": "cell-by-cell", "rows": rows,
                "cols": ncols, "pre_ok": bool(ok_pre)}

    # -- data_put 的三级写入通道 -----------------------------------------

    def _put_once(self, bk: str, rect, r1: int, c1: int, timeout: int = 300):
        """执行一次 PutWorksheet；异常即判定会话失效并重连。"""
        def _put():
            return bool(self.app.PutWorksheet(bk, rect, int(r1), int(c1)))
        try:
            return bool(self.safe_op(
                lambda: self.worker.submit(_put, timeout=timeout)))
        except Exception as exc:  # noqa: BLE001
            log.warning("PutWorksheet 异常(%s)，标记会话失效", exc)
            self.invalidate(exc)
            return False

    def _put_block(self, bk: str, rect, r1: int, c1: int) -> bool:
        """整块写入，失败后重连再给一次机会（实测会话假死重连即可恢复）。"""
        if self._put_once(bk, rect, r1, c1):
            return True
        log.warning("整块 PutWorksheet 失败，重连后重试一次")
        self.ensure_connected()
        self.ex(f"win -a {bk};")
        return self._put_once(bk, rect, r1, c1)

    def _put_chunked(self, bk: str, rect, r1: int, c1: int) -> bool:
        """按 PUT_CHUNK_ROWS 分块写入（起始行递增）。

        实测整块 68804x20 约 0.4s、20000 行分块约 0.5s，性能可忽略，
        但分块能绕开超大数组的一次性 COM 编组失败。
        """
        total = len(rect)
        if total <= PUT_CHUNK_ROWS:
            return False  # 小表分块没有意义，直接走逐格兜底
        off = 0
        while off < total:
            block = rect[off:off + PUT_CHUNK_ROWS]
            if not self._put_once(bk, block, r1 + off, c1):
                log.warning("分块写入在第 %d 行处失败（共 %d 行）", off, total)
                return False
            off += PUT_CHUNK_ROWS
        return True

    def worksheet_get_data(self, book: str) -> dict:
        bk = self._resolve_page(book, "worksheet")

        def _get():
            raw = self.app.GetWorksheet(bk)
            if raw is None:
                return None
            return [list(r) for r in raw]
        try:
            data = self.safe_op(lambda: self.worker.submit(_get, timeout=120))
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        if data is None:
            return {"ok": False, "error": f"找不到工作簿 {book}",
                    "existing": self._ws_names()}
        return {"ok": True, "rows": len(data), "cols": len(data[0]) if data else 0,
                "data": data}

    def worksheet_export_csv(self, filepath: str, book: str | None = None) -> dict:
        # expasc 的静默失败与 expGraph 同源，因此直接用 COM 读数后由 Python 落盘
        r = self.worksheet_get_data(book)
        if not r.get("ok"):
            return {"ok": False, "error": r.get("error", "读数失败")}
        p = Path(_bsl(filepath))
        p.parent.mkdir(parents=True, exist_ok=True)
        import csv as _csv
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            _csv.writer(f).writerows(r["data"])
        return {"ok": True, "path": str(p), "rows": r.get("rows")}

    # -- 绘图 ---------------------------------------------------------------

    def plot_create(self, book: str, x_col: int = 1,
                    y_cols: list[int] | None = None,
                    plot_type: str | int = "line_symbol",
                    template: str | None = None,
                    graph_name: str | None = None,
                    pairs: list[list[int]] | None = None,
                    err_cols: list[int | None] | None = None) -> dict:
        """实测：plotxy 多 Y 仅支持连续列区间 iy:=(x,(a:b))；
        非连续列的追加语法（ogl:=[G]1 / layer -i 等）在 COM 下均无效。
        因此非连续列自动拆分为多个图，并在返回值中说明。
        pairs=[[1,2],[3,4],...] 时按显式 XY 对绘制（各 X 列可独立），
        实测 iy:=((1,2),(3,4)) 语法有效且曲线数正确。"""
        ys = sorted(set(int(v) for v in (y_cols or [2])))
        if isinstance(plot_type, int):
            pid, key = plot_type, None
        else:
            key = str(plot_type).lower().strip()
            pid = PLOT_TYPE_IDS.get(key)
            if pid is None:
                return {"ok": False,
                        "error": f"未知 plot_type: {plot_type}，"
                                 f"可选 {sorted(set(PLOT_TYPE_IDS))} 或整数 ID"}
        tpl = template if template else TEMPLATE_NAMES.get(key or "")
        tpl_clause = f" template:={_clean_text(tpl)}" if tpl else ""
        base_name = _sanitize_page_name(
            _clean_text(graph_name) if graph_name else "Graph")
        bk = self._resolve_page(book, "worksheet")

        def _count_curves(g: str):
            """统计图层里的曲线数；图层取不到时返回 None 而不是抛异常。"""
            gl = self._find_graph_layer(g)
            if gl is None:
                return None

            def _c():
                dp = gl.DataPlots
                cnt = int(dp.Count)
                names = []
                for i in range(0, cnt + 1):
                    try:
                        names.append(str(dp.Item(i).Name))
                    except Exception:  # noqa: BLE001
                        continue
                return names
            try:
                names = self.safe_op(
                    lambda: self.worker.submit(_c, timeout=30))
            except Exception:  # noqa: BLE001
                return None
            return len([n for n in names if n])

        def _iy(cols: list[int]) -> str:
            # 关键语法：多 Y 必须用 (x,a:b) —— 内层再套括号会被解析成 Y-vs-Y
            if len(cols) == 1:
                return f"({x_col},{cols[0]})"
            return f"({x_col},{cols[0]}:{cols[-1]})"

        # 拆分连续段
        runs: list[list[int]] = []
        cur = [ys[0]]
        for v in ys[1:]:
            if v == cur[-1] + 1:
                cur.append(v)
            else:
                runs.append(cur)
                cur = [v]
        runs.append(cur)

        if pairs:
            # err_cols 与 pairs 一一对应（元素可为 None）：该对用
            # 选区方式绘制（worksheet -s + 无 iy 的 plotxy），
            # 选区含 [X,Y,Err] 三列时 Origin 自动加误差棒（实测）。
            err_cols = list(err_cols) if err_cols else []
            gname = self._unique_name(base_name)
            if not any(err_cols):
                iy_expr = "(" + ",".join(
                    f"({int(a)},{int(b)})" for a, b in pairs) + ")"
                ok = self.ex("\n".join([
                    f"win -a {bk};",
                    f"plotxy iy:={iy_expr} plot:={pid} ogl:=[<new{tpl_clause}>];",
                    f"win -r %H {gname};"]))
                # 关键：Origin 会净化图形名，必须用回读到的实际名，
                # 否则 verified 恒为 False、FindGraphLayer 恒为 None
                actual, _ = self._actual_after_new(gname)
            else:
                nrows = 0
                try:
                    def _nr():
                        self.app.Execute(f"win -a {bk};")
                        return float(self.app.Evaluate("wks.nrows"))
                    nrows = int(self.safe_op(
                        lambda: self.worker.submit(_nr, timeout=30)))
                except Exception:  # noqa: BLE001
                    pass
                if nrows <= 0:
                    return {"ok": False,
                            "error": "无法读取工作簿行数，选区绘图失败"}
                # 第一对：新建图并立即改名
                a0, b0 = pairs[0]
                e0 = err_cols[0] if err_cols else None
                if e0:
                    first = (f"win -a {bk};"
                             f"worksheet -s {int(a0)} 1 {int(e0)} {nrows};"
                             f"plotxy plot:={pid} ogl:=[<new{tpl_clause}>];")
                else:
                    first = (f"win -a {bk};"
                             f"plotxy iy:=({int(a0)},{int(b0)}) plot:={pid} "
                             f"ogl:=[<new{tpl_clause}>];")
                ok = self.ex(first)
                actual, matched = self._rename_active_page(gname)
                if not actual:
                    return {"ok": False, "graph": None,
                            "error": f"绘图后未能识别图形窗口（请求 {gname}）",
                            "graphs": self._graph_names()}
                if not matched:
                    log.warning("图形名被净化: %s -> %s", gname, actual)
                # 剩余对：用【回读到的实际名】追加，否则 ogl:=[名字] 找不到图层
                for k, (a, b) in enumerate(pairs[1:], start=1):
                    e = err_cols[k] if k < len(err_cols) else None
                    # 关键：每对绘制前重新激活工作簿——上一次 plotxy 会把
                    # 活动窗口切到新图，导致 worksheet -s 选区失效（实测）
                    if e:
                        cmd = (f"win -a {bk};"
                               f"worksheet -s {int(a)} 1 {int(e)} {nrows};"
                               f"plotxy plot:={pid} ogl:=[{actual}]1;")
                    else:
                        cmd = (f"win -a {bk};"
                               f"plotxy iy:=({int(a)},{int(b)}) plot:={pid} "
                               f"ogl:=[{actual}]1;")
                    ok = self.ex(cmd) and ok
            time.sleep(1.2)
            exists = _name_key(actual) in {_name_key(n) for n in self._graph_names()}
            curves = _count_curves(actual)
            res: dict = {"ok": bool(exists), "graph": actual if exists else None,
                         "book": bk, "curves_detected": curves,
                         "plot_type_id": pid, "template": tpl, "pairs": pairs,
                         "verified_by": "pages" if exists else None}
            if not exists:
                res["error"] = f"图形窗口 {actual} 未出现在页面列表中"
                res["graphs"] = self._graph_names()
            elif curves is None:
                # 图已生成但图层对象取不到时不要判 ok=False——这是 COM 侧的
                # 读取问题，不是绘图失败
                res["note"] = ("图形已生成，但 FindGraphLayer 未取到图层对象，"
                               "曲线数未能核对")
            if actual != _clean_text(graph_name or ""):
                res["requested"] = _clean_text(graph_name) if graph_name else None
                res["warning"] = (f"Origin 净化了图形窗口名："
                                  f"{graph_name or 'Graph'} -> {actual}")
            return res

        graphs: list[str] = []
        notes: list[str] = []
        for k, run in enumerate(runs):
            gname = self._unique_name(
                base_name if k == 0 else f"{base_name}_p{k+1}")
            cmds = [f"win -a {bk};",
                    f"plotxy iy:={_iy(run)} plot:={pid} "
                    f"ogl:=[<new{tpl_clause}>];",
                    f"win -r %H {gname};"]
            self.ex("\n".join(cmds))
            time.sleep(1.2)
            # 关键：回读实际名（Origin 会净化），否则 verified 恒为 False
            actual, _ = self._actual_after_new(gname)
            exists = _name_key(actual) in {_name_key(n)
                                           for n in self._graph_names()}
            if not exists:
                notes.append(f"第 {k+1} 组列 {run} 绘图未验证成功"
                             f"（请求名 {gname}，实际 {actual or '无'}）")
                continue
            graphs.append(actual)
        curves = _count_curves(graphs[0]) if graphs else None
        if len(runs) > 1 and graphs:
            notes.append(f"Y 列不连续，已拆分为 {len(graphs)} 个图；"
                         "若需单图请让 Y 列相邻")
        res = {"ok": bool(graphs), "graph": graphs[0] if graphs else None,
               "extra_graphs": graphs[1:], "curves_detected": curves,
               "plot_type_id": pid, "template": tpl, "y_cols": ys,
               "notes": notes or None}
        if graphs and graphs[0] != _clean_text(graph_name or ""):
            res["warning"] = (f"Origin 净化了图形窗口名："
                              f"{graph_name or 'Graph'} -> {graphs[0]}")
        if graphs and curves is None:
            res["note"] = ("图形已生成，但 FindGraphLayer 未取到图层对象，"
                           "曲线数未能核对")
        return res

    def axis_get(self, graph: str, axis: str = "x") -> dict:
        """读回坐标轴真实状态（from/to/inc/type）。

        唯一可靠通道是 COM GraphLayer.GetNumProp，不是 LabTalk layer.x.from：
        - LabTalk 的 layer.* 永远作用于【当前活动窗口】的图层，活动窗口是工作簿
          时会读到工作表图层的垃圾值（实测 0/4/8）或 NANUM 哨兵
        - GetNumProp 直接问图形图层，不依赖活动窗口，大小写不敏感
        """
        g = self._resolve_page(graph, "graph")
        ax = axis.lower()
        if ax not in ("x", "y"):
            return {"ok": False, "error": "axis 仅支持 x/y"}
        gl = self._find_graph_layer(g)
        if gl is None:
            return {"ok": False, "graph": g,
                    "error": f"找不到图形窗口或图层 {g}",
                    "graphs": self._graph_names()}

        def _read():
            out = {}
            for p in ("from", "to", "inc", "type"):
                try:
                    out[p] = float(gl.GetNumProp(f"{ax}.{p}"))
                except Exception:  # noqa: BLE001
                    out[p] = None
            return out
        try:
            vals = self.safe_op(lambda: self.worker.submit(_read, timeout=30))
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "graph": g, "error": str(exc)}
        bad = [k for k, v in vals.items() if v is None or _is_nanum(v)]
        return {"ok": not bad, "graph": g, "axis": ax,
                "from": vals.get("from"), "to": vals.get("to"),
                "increment": vals.get("inc"), "scale": vals.get("type"),
                "invalid_fields": bad or None,
                "note": None if not bad else
                        f"字段 {bad} 读到无效值，确认 {g} 是否为图形窗口"}

    def axis_set(self, graph: str, axis: str = "x", title: str | None = None,
                 vmin: float | None = None, vmax: float | None = None,
                 scale: str | None = None,
                 major_increment: float | None = None,
                 rescale: bool | None = None) -> dict:
        """设置坐标轴：标题、范围、刻度类型、主刻度增量。

        rescale（是否锁定轴）：
          None（默认）—— 仅在显式给了 vmin/vmax 时自动锁定（layer.x.rescale=0）。
          这是必须的：Origin 默认的自动缩放会在后续任何重绘时把手动范围
          改回去（实测出现 -2 之类的自动边距），不锁定等于没设。
          True/False 可显式控制。
        """
        g = self._resolve_page(graph, "graph")
        ax = axis.lower()
        if ax not in ("x", "y"):
            return {"ok": False, "error": "axis 仅支持 x/y"}
        lines = [f"win -a {g};"]
        if title is not None:
            obj = "xb" if ax == "x" else "yl"
            lines.append(f'{obj}.text$ = "{_clean_text(title)}";')
        if vmin is not None:
            lines.append(f"layer.{ax}.from = {float(vmin)};")
        if vmax is not None:
            lines.append(f"layer.{ax}.to = {float(vmax)};")
        if scale:
            sc = {"linear": 1, "log10": 2, "log": 2}.get(scale.lower())
            if sc is None:
                return {"ok": False, "error": "scale 仅支持 linear/log10"}
            lines.append(f"layer.{ax}.type = {sc};")
        if major_increment is not None:
            lines.append(f"layer.{ax}.inc = {float(major_increment)};")
        lock = rescale if rescale is not None else (vmin is not None
                                                    or vmax is not None)
        if lock:
            # 实测：不加 rescale=0，后续操作（追加数据图/导出）会把范围改回去
            lines.append(f"layer.{ax}.rescale = 0;")
        ok = self.ex("\n".join(lines))
        res: dict = {"ok": bool(ok), "graph": g, "axis": ax, "locked": bool(lock)}
        if g != _clean_text(graph):
            res["resolved_from"] = _clean_text(graph)
        # 回读校验：写没写进去只有读回来才知道
        if vmin is not None or vmax is not None:
            read = self.axis_get(g, ax)
            res["verified"] = read.get("ok") and read.get("from") is not None
            res["actual"] = {"from": read.get("from"), "to": read.get("to")}
        return res

    def series_style(self, graph: str, series_index: int = 1,
                     color: str | None = None, line_width_pt: float | None = None,
                     symbol_size: float | None = None,
                     symbol_shape_index: int | None = None,
                     y_col: int | None = None) -> dict:
        """实测：LabTalk set 命令在 COM 下静默无效（返回 True 但不改样式）。
        唯一可靠通道是 COM DataPlot.SetNumProp：
        - 'color' 只接受 Origin 颜色索引(1-24)，RGB 复合值不渲染
        - 'line.width' 单位 pt（默认 0.6）
        y_col 给定时按数据集名(如 Demo_B)定位曲线——误差棒模式下
        DataPlots 索引会错位，推荐用 y_col。"""
        g = self._resolve_page(graph, "graph")
        idx = int(series_index) - 1
        # 关键：图不存在时 FindGraphLayer 返回 None，继续访问 .DataPlots 会抛
        # "'NoneType' object has no attribute 'DataPlots'"。这里先拦住，
        # 返回可读的错误而不是刷屏异常。
        if self._find_graph_layer(g) is None:
            return {"ok": False, "graph": g,
                    "error": f"找不到图形窗口或图层：{graph}"
                             f"（解析为 {g}）",
                    "graphs": self._graph_names(),
                    "hint": "先用 plot_create 建图，并确认用的是回读到的实际窗口名"}

        if color:
            cname = str(color).lower().strip()
            if cname in COLOR_INDEX:
                ci = COLOR_INDEX[cname]
            elif cname.isdigit():
                ci = int(cname)
            else:
                return {"ok": False,
                        "error": f"颜色 {color} 需为索引表内名称或 1-24 数字"
                                 "（COM 通道不支持 R,G,B）"}
        applied: list[str] = []

        def _style():
            gl = self.app.FindGraphLayer(g)
            if gl is None:
                raise RuntimeError(f"找不到图形图层 {g}（可能已在 Origin 中关闭）")
            if y_col is not None:
                target = None
                n = int(gl.DataPlots.Count)
                for i in range(n):
                    try:
                        dp_i = gl.DataPlots.Item(i)
                        name = str(dp_i.GetDatasetName())
                    except Exception:  # noqa: BLE001
                        continue
                    # 数据集名形如 Book_C（C = y_col 对应列字母）
                    if name.upper().endswith("_" + _col_letter(int(y_col))):
                        target = dp_i
                        break
                if target is None:
                    raise RuntimeError(f"未找到 y_col={y_col} 对应的数据图")
                dp = target
            else:
                dp = gl.DataPlots.Item(idx)
            if color:
                dp.SetNumProp("color", float(ci))
                applied.append("color")
            if line_width_pt is not None:
                dp.SetNumProp("line.width", float(line_width_pt))
                applied.append("line_width")
            if symbol_size is not None:
                dp.SetNumProp("symbol.size", float(symbol_size))
                applied.append("symbol_size")
            if symbol_shape_index is not None:
                dp.SetNumProp("symbol.shape", float(symbol_shape_index))
                applied.append("symbol_shape")
            return True
        try:
            self.safe_op(lambda: self.worker.submit(_style, timeout=30))
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        return {"ok": True, "graph": g, "applied": applied,
                "note": "颜色用 Origin 索引: black1 red2 green3 blue4 cyan5 "
                        "magenta6 yellow7 dark_yellow8 ... grey14"}

    def text_label(self, text: str, name: str = "mcpLabel",
                   x: float | None = None, y: float | None = None,
                   font_size: int | None = None,
                   graph: str | None = None) -> dict:
        pre = (f"win -a {self._resolve_page(graph)};" if graph else "")
        txt = _clean_text(text)
        nm = _clean_text(name)
        fs = f" -fs {int(font_size)}" if font_size else ""
        if x is not None and y is not None:
            cmd = f'label -px {float(x)} {float(y)} -s{fs} -n {nm} "{txt}";'
        else:
            cmd = f'label -s -sa{fs} -n {nm} "{txt}";'
        return {"ok": self.ex(pre + cmd), "label": nm}

    def graph_frame(self, graph: str, boxed: bool = True) -> dict:
        """显示/隐藏图层边框（上/右轴线）。实测唯一有效通道：
        COM GraphLayer.SetNumProp('showframe', 1/0)。"""
        g = self._resolve_page(graph, "graph")
        if self._find_graph_layer(g) is None:
            return {"ok": False, "graph": g,
                    "error": f"找不到图形窗口或图层：{graph}（解析为 {g}）",
                    "graphs": self._graph_names()}

        def _f():
            gl = self.app.FindGraphLayer(g)
            if gl is None:
                raise RuntimeError(f"找不到图形图层 {g}")
            gl.SetNumProp("showframe", 1.0 if boxed else 0.0)
            return float(gl.GetNumProp("showframe"))
        try:
            v = self.safe_op(lambda: self.worker.submit(_f, timeout=30))
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        return {"ok": True, "boxed": bool(v), "graph": g}

    def add_ref_line(self, graph: str, axis: str = "x", pos: float = 0.0,
                     color: int = 1, width: float = 1.0, dash: bool = True,
                     label: str | None = None, book: str | None = None,
                     span: list[float] | None = None) -> dict:
        """添加参考线（axis='x' → 竖线 x=pos；axis='y' → 横线 y=pos）。

        改用 Origin 原生参考线对象 `layer.{axis}.refline#`（Origin 2018 SR1+，
        见 scripts/probe_refline_native.py 实测），替代早前「两点数据图 +
        ogl:=[G]1 追加」的 hack。原生 refline 三胜：
        1. 不进图例（DataPlots 不变），无需再调 legend_remove_last；
        2. 不触发轴重新缩放（轴范围不变），无需锁轴回写；
        3. 不往数据表追加列。
        `book` / `span` 已无实际用途，仅为兼容旧调用保留（直接忽略）。
        label 给定时写入 refline 原生标签（labeltext$ + labelshow）。
        """
        g = self._resolve_page(graph, "graph")
        ax = axis.lower()
        if ax not in ("x", "y"):
            return {"ok": False, "error": "axis 仅支持 x/y"}
        if self._find_graph_layer(g) is None:
            return {"ok": False, "graph": g,
                    "error": f"找不到图形窗口或图层：{graph}（解析为 {g}）",
                    "graphs": self._graph_names()}

        def _run():
            app = self.app
            app.Execute(f"win -a {g};")
            # 读当前参考线条数决定新索引（读不到就从 1 开始）
            cnt = 0
            for _ in range(4):
                try:
                    v = app.Evaluate(f"layer.{ax}.reflines.count")
                    if v is not None and not _is_nanum(v):
                        cnt = int(float(v))
                        break
                except Exception:  # noqa: BLE001
                    pass
                time.sleep(0.2)
            idx = cnt + 1
            style = 2 if dash else 1  # refline linestyle: 2=虚线, 1=实线
            cmds = [
                f"layer.{ax}.reflines.count = {idx};",
                f"layer.{ax}.refline{idx}.value = {float(pos)};",
                f"layer.{ax}.reflines.lineshow = 1;",
                f"layer.{ax}.refline{idx}.linecolor = {int(color)};",
                f"layer.{ax}.refline{idx}.linestyle = {style};",
                f"layer.{ax}.refline{idx}.linethickness = {float(width)};",
            ]
            if label:
                cmds.append(f'layer.{ax}.refline{idx}.labeltext$ = '
                            f'"{_clean_text(label)}";')
                cmds.append(f"layer.{ax}.refline{idx}.labelshow = 1;")
            ok = app.Execute("\n".join(cmds))
            time.sleep(0.3)
            got = None
            try:
                got = app.Evaluate(f"layer.{ax}.refline{idx}.value")
            except Exception:  # noqa: BLE001
                got = None
            return idx, bool(ok), got
        try:
            idx, ok, got = self.safe_op(
                lambda: self.worker.submit(_run, timeout=60))
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        verified = bool(ok) and got is not None and \
            abs(float(got) - float(pos)) < 1e-9
        res: dict = {"ok": verified, "graph": g, "axis": ax,
                     "pos": float(pos), "refline_index": idx,
                     "native": "refline"}
        if g != _clean_text(graph):
            res["resolved_from"] = _clean_text(graph)
        if not verified:
            res["error"] = (f"参考线写入未验证：读回 value={got}，"
                            f"期望 {float(pos)}")
        return res

    def legend_remove_last(self, graph: str, count: int = 1) -> dict:
        """删除图例最后 count 行。

        必须在所有绘图/坐标轴操作之后调用——Origin 会在后续操作时
        重建图例，提前删除会被恢复（实测）。

        注意：add_ref_line 已改用原生 refline，**不再**进图例；本工具
        用于清理其它会进图例的条目（如多余的数据曲线、误差棒分组等）。
        """
        g = self._resolve_page(graph, "graph")
        if self._find_graph_layer(g) is None:
            return {"ok": False, "graph": g,
                    "error": f"找不到图形窗口或图层：{graph}（解析为 {g}）",
                    "graphs": self._graph_names()}

        def _run():
            # 关键：先关掉图例自动更新，否则导出时 Origin 会重建图例，
            # 手工裁剪的文本被覆盖（实测）
            try:
                n = int(self.app.GraphPages.Count)
                for i in range(0, n + 1):
                    try:
                        pg = self.app.GraphPages.Item(i)
                    except Exception:  # noqa: BLE001
                        continue
                    if str(pg.Name) == g:
                        pg.LegendsAutoUpdate = 0
                        break
            except Exception:  # noqa: BLE001
                pass
            gl = self.app.FindGraphLayer(g)
            for i in range(0, int(gl.GraphObjects.Count) + 1):
                try:
                    o = gl.GraphObjects.Item(i)
                except Exception:  # noqa: BLE001
                    continue
                if o.Name != "Legend":
                    continue
                lines = [l for l in str(o.Text).split("\r\n") if l.strip()]
                o.Text = "\r\n".join(lines[:-count] if count > 0 else lines)
                return str(o.Text)
            return None
        try:
            txt = self.safe_op(lambda: self.worker.submit(_run, timeout=30))
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        return {"ok": txt is not None, "legend": txt}

    # -- 导出 ---------------------------------------------------------------

    def graph_export(self, filepath: str, fmt: str | None = None,
                     width: int | None = None, height: int | None = None,
                     dpi: int | None = None, graph: str | None = None) -> dict:
        p = Path(_bsl(filepath))
        p.parent.mkdir(parents=True, exist_ok=True)
        ext = p.suffix.lower().lstrip(".")
        ftype = EXPORT_TYPES.get((fmt or ext).lower())
        if not ftype:
            return {"ok": False, "error":
                    f"不支持的格式: {fmt or ext}，可用 {sorted(EXPORT_TYPES)}"
                    "（svg 在部分版本的 COM 导出中不可用）"}
        gname = self._resolve_page(graph) if graph else None
        pre = f"win -a {gname};" if gname else ""
        bare = f"mcp_exp_{datetime.now().strftime('%H%M%S%f')}.{ftype}"
        uff = self._user_files_folder()

        attempts: list[str] = []
        # 方式一（实测可靠）：裸文件名导出 → UFF → 复制到目标。
        # 注意绝不能带 width/height/tr1.* 子句——它们会导致静默失败。
        ok = self.ex(pre + f'expGraph type:={ftype} filename:="{bare}";')
        attempts.append(f"expGraph bare={bool(ok)}")
        src = Path(uff) / bare
        for _ in range(10):
            if src.exists():
                break
            time.sleep(1.0)
        if src.exists():
            shutil.copyfile(src, p)
            try:
                src.unlink()
            except OSError:
                pass
            return {"ok": True, "path": str(p), "size": p.stat().st_size,
                    "note": "分辨率由 Origin 默认导出设置决定（通常已满足出版需求）；"
                            "如需自定义尺寸请在 Origin 图形导出对话框中预设"}
        # 方式二回退：page.export
        ok2 = self.ex(pre + f'page.export(type:={ftype}, filename:="{bare}", '
                          f"overwrite:=replace);")
        attempts.append(f"page.export={bool(ok2)}")
        for _ in range(6):
            if src.exists():
                break
            time.sleep(1.0)
        if src.exists():
            shutil.copyfile(src, p)
            try:
                src.unlink()
            except OSError:
                pass
            return {"ok": True, "path": str(p), "size": p.stat().st_size}
        return {"ok": False, "path": str(p), "attempts": attempts,
                "hint": "两种导出方式都未生成文件；检查目标图是否存在且有效"}

    # -- LabTalk ------------------------------------------------------------

    def labtalk_execute(self, script: str) -> dict:
        return {"ok": self.ex(script),
                "note": "需要读回数值时请改用 labtalk_evaluate"}

    def labtalk_evaluate(self, expr: str, window: str | None = None) -> dict:
        """求值 LabTalk 表达式。

        window 强烈建议在读 `layer.*` 时给出：LabTalk 的 layer.* 永远作用于
        【当前活动窗口】的图层。活动窗口是工作簿时，layer.y.from 读到的是
        工作表图层的垃圾值（实测 0/4/8），很多情况下直接返回 NANUM 哨兵
        -1.23456789e-300。给 window 会先执行 win -a 再求值。

        读坐标轴范围请优先用 axis_get（走 COM GetNumProp，不依赖活动窗口）。
        """
        script = _clean_text(expr)
        tgt = None
        if window:
            tgt = self._resolve_page(window)
            # 关键：Evaluate() 只接受【表达式】，不能把 win -a X; 拼进去当语句
            # 求值（实测返回 None）。必须先用 Execute 激活，再单独求值表达式。
            self.ex(f"win -a {tgt};")
            time.sleep(0.2)

        def _eval():
            return self.app.Evaluate(script)
        try:
            val = self.safe_op(lambda: self.worker.submit(_eval, timeout=60))
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}
        res: dict = {}
        if tgt:
            res["window"] = tgt
        if val is None:
            res.update({"ok": False, "value": None, "type": "none",
                        "error": "Evaluate 返回 None（表达式无效或不被支持）",
                        "hint": ("Evaluate 只能求值表达式，不能执行语句；"
                                 "执行命令请用 labtalk_execute")})
            return res
        if _is_nanum(val):
            res.update({"ok": False, "value": None, "type": "nanum",
                        "error": (f"表达式在当前活动窗口下无效，返回 Origin "
                                  f"缺失值哨兵 {val}"),
                        "hint": ("读 layer.* 请传 window=<图形名>，或改用 "
                                 "axis_get(graph, axis)（COM 通道）")})
            return res
        try:
            num = float(val)
        except (TypeError, ValueError):
            res.update({"ok": True, "value": val, "type": type(val).__name__})
            return res
        res.update({"ok": True, "value": num, "type": "number"})
        return res

    def quit_origin(self, force: bool = False) -> dict:
        if not force:
            return {"ok": False,
                    "hint": "拒绝退出：可能有未保存工作。确认后 force=true 且征得用户同意"}
        def _quit():
            self.app.Exit()
        try:
            self.worker.submit(_quit, timeout=30)
        finally:
            self.invalidate()
        return {"ok": True}


def _is_number(s: str) -> bool:
    try:
        float(s)
        return True
    except ValueError:
        return False


# ---------------------------------------------------------------------------
# 环境检测（纯 Python，不触碰 COM）
# ---------------------------------------------------------------------------
def _clsid_exe_path(progid: str) -> str | None:
    """从 COM 注册表读 Origin 可执行文件路径（比扫目录可靠）。

    注意注册表 32/64 位视图：Origin 安装程序常把 CLSID 写进 WOW6432Node，
    因此需要 KEY_WOW64_32KEY 显式访问。
    """
    if winreg is None:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID") as k:
            clsid = winreg.QueryValueEx(k, "")[0]
    except OSError:
        return None
    for view in (0, winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                                f"CLSID\\{clsid}\\LocalServer32",
                                0, winreg.KEY_READ | view) as k2:
                val = winreg.QueryValueEx(k2, "")[0]
            if val:
                return os.path.expandvars(val.strip().strip('"'))
        except OSError:
            continue
    return None


def detect_origin_install() -> dict:
    result: dict = {"exe_paths": [], "registry_entries": {},
                    "running_processes": [], "progids_registered": []}
    if winreg is not None:
        for pid in ("Origin.ApplicationSI", "Origin.Application"):
            try:
                winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, pid)
                result["progids_registered"].append(pid)
            except OSError:
                pass
        for pid in result["progids_registered"]:
            exe = _clsid_exe_path(pid)
            if exe and exe not in result["exe_paths"]:
                result["exe_paths"].append(exe)
        for sub in (r"SOFTWARE\OriginLab", r"WOW6432Node\OriginLab"):
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub) as key:
                    i = 0
                    while True:
                        try:
                            ver = winreg.EnumKey(key, i)
                            i += 1
                        except OSError:
                            break
                        result["registry_entries"].setdefault(ver, "?")
            except OSError:
                pass
    bases = [os.environ.get("ProgramFiles", r"C:\Program Files"),
             r"D:\Program Files", r"C:\Program Files (x86)",
             r"D:\Program Files (x86)", r"E:\Program Files"]
    for base in bases:
        olab = Path(base) / "OriginLab"
        if olab.is_dir():
            for d in sorted(olab.iterdir()):
                for cand in (d / "Origin64.exe", d / "Origin.exe"):
                    if cand.exists() and str(cand) not in result["exe_paths"]:
                        result["exe_paths"].append(str(cand))
    try:
        out = os.popen("tasklist /FO CSV /NH").read()
        procs = []
        for line in out.splitlines():
            parts = line.replace('"', "").split(",")
            if parts and parts[0].lower().startswith("origin"):
                procs.append(parts[0])
        result["running_processes"] = sorted(set(procs))
    except Exception:  # noqa: BLE001
        pass
    result["installed"] = bool(result["exe_paths"] or result["registry_entries"])
    return result


# ---------------------------------------------------------------------------
# MCP 服务定义（兼容 mcp SDK 1.x / 2.x）
# ---------------------------------------------------------------------------
try:
    from mcp.server.mcpserver import MCPServer as _FastMCP  # mcp >= 2.0
except ImportError:  # pragma: no cover
    from mcp.server.fastmcp import FastMCP as _FastMCP      # mcp 1.x  # noqa: N813

session = OriginSession()

mcp = _FastMCP(
    "origin-mcp",
    instructions=(
        "Origin 科研绘图自动化服务器。典型流程：origin_status → origin_connect → "
        "workbook_new → data_put → plot_create → axis_set → graph_export / "
        "project_save。所有涉及文件的返回值均已做真实存在性校验。"
        "未经用户同意不要 quit_origin 或新建工程覆盖其未保存内容。"
        "重要：Origin 会净化窗口名（删除非 ASCII 字母数字并截断到 13 字符，"
        "如 Data_0.6 -> Data06），workbook_new/plot_create 返回的是回读到的"
        "实际名，后续调用一律使用返回值里的名字。"
        "读坐标轴范围用 axis_get，不要用 labtalk_evaluate('layer.x.from')。"
    ),
)


@mcp.tool()
def origin_status() -> dict:
    """检测本机 Origin 安装状态（安装路径、注册表、运行中进程、COM 注册）。不启动 Origin。"""
    return detect_origin_install()


@mcp.tool()
def origin_connect(visible: bool = True) -> dict:
    """连接（必要时启动）Origin 并保持长驻会话。首次启动需等待加载。"""
    session.visible = visible
    return session.connect(visible=visible)


@mcp.tool()
def origin_disconnect() -> dict:
    """断开 MCP 与 Origin 的 COM 会话（不关闭 Origin 程序本身）。"""
    session.invalidate()
    return {"ok": True}


@mcp.tool()
def project_new() -> dict:
    """新建空白 Origin 工程。注意：当前工程中未保存的内容会丢失，调用前应提醒用户。"""
    session.ensure_connected()
    return session.project_new()


@mcp.tool()
def workbook_new(name: str | None = None, sheets: int = 1) -> dict:
    """新建工作簿并按 name 命名（自动去重），返回**实际**窗口名。

    注意：Origin 会净化窗口名（删掉所有非 ASCII 字母数字并截断到 13 字符），
    例如 Data_0.6 -> Data06、Force_Displacement_0.6 -> ForceDisplace。
    返回值里的 "book" 是回读到的真实名字，后续所有调用都必须用它。"""
    session.ensure_connected()
    return session.workbook_new(name=name, sheets=sheets)


@mcp.tool()
def worksheet_set_columns(book: str, columns: list[dict]) -> dict:
    """设置列属性。columns 元素形如 {"index":1,"name":"Time","units":"s","comments":"...","type":"x"}；
    type 可选 x/y/z/xerr/yerr/label/none，index 从 1 开始。"""
    session.ensure_connected()
    return session.worksheet_set_columns(book=book, columns=columns)


@mcp.tool()
def data_put(book: str, data: list[list], start_row: int = 1,
             start_col: int = 1, header: list[str] | None = None) -> dict:
    """向工作簿活动表批量写入二维数组（数字或字符串混合），start_row/start_col 从 1 起。
    header 为可选列长名称列表。
    写入链路：整块 PutWorksheet →（失败）重连重试 →（失败）按 20000 行分块 →
    （失败且为小表）逐格。大表绝不会退化成逐格写入。
    book 可以传未净化的名字（Data_0.6），会自动解析到实际窗口（Data06）。"""
    session.ensure_connected()
    return session.data_put(book=book, data=data, start_row=start_row,
                            start_col=start_col, header=header)


@mcp.tool()
def worksheet_get_data(book: str) -> dict:
    """读回工作簿活动表的全部数据（用于核对写入结果或分析）。"""
    session.ensure_connected()
    return session.worksheet_get_data(book=book)


@mcp.tool()
def plot_create(book: str, x_col: int = 1, y_cols: list[int] | None = None,
                plot_type: str | int = "line_symbol",
                template: str | None = None, graph_name: str | None = None,
                pairs: list[list[int]] | None = None,
                err_cols: list[int | None] | None = None) -> dict:
    """从工作簿数据创建折线/散点/点线/柱状图到新 Graph 窗口。
    plot_type 可选 line/scatter/line_symbol/column/bar 或 plotxy 整数 ID；
    共用 X 列时用 y_cols（需相邻），各 X 独立时用 pairs=[[x1,y1],[x2,y2],...]。
    err_cols 与 pairs 对应（元素可为 null）：给出误差列号时该对以选区方式
    绘制并自动加 Y 误差棒（误差列需已用 worksheet_set_columns 设为 yerr）。
    返回的 graph 是回读到的**实际**窗口名（Origin 会净化图形名），曲线数已核对；
    若图形已生成但图层对象取不到，ok 仍为 True 并附 note。"""
    session.ensure_connected()
    return session.plot_create(book=book, x_col=x_col, y_cols=y_cols,
                               plot_type=plot_type, template=template,
                               graph_name=graph_name, pairs=pairs,
                               err_cols=err_cols)


@mcp.tool()
def axis_get(graph: str, axis: str = "x") -> dict:
    """读回坐标轴真实范围/增量/刻度类型（from/to/increment/scale）。

    这是读轴状态的**首选**通道（COM GraphLayer.GetNumProp），不依赖活动窗口。
    LabTalk 的 layer.x.from 只在活动窗口恰好是该图时才对，否则会读到工作表
    图层的垃圾值或 NANUM 哨兵 -1.23456789e-300。"""
    session.ensure_connected()
    return session.axis_get(graph=graph, axis=axis)


@mcp.tool()
def axis_set(graph: str, axis: str = "x", title: str | None = None,
             vmin: float | None = None, vmax: float | None = None,
             scale: str | None = None, major_increment: float | None = None,
             rescale: bool | None = None) -> dict:
    """设置图形坐标轴：标题、范围(min/max)、刻度类型(linear/log10)、主刻度增量。

    给了 vmin/vmax 时默认会同时锁定该轴（layer.x.rescale = 0）——不锁的话
    Origin 的自动缩放会在后续任何重绘时把手动范围改回去（表现为多出 -2 之类
    的自动边距）。传 rescale=False 可关闭锁定。"""
    session.ensure_connected()
    return session.axis_set(graph=graph, axis=axis, title=title, vmin=vmin,
                            vmax=vmax, scale=scale,
                            major_increment=major_increment, rescale=rescale)


@mcp.tool()
def series_style(graph: str, series_index: int = 1, color: str | None = None,
                 line_width_pt: float | None = None,
                 symbol_size: float | None = None,
                 symbol_shape_index: int | None = None,
                 y_col: int | None = None) -> dict:
    """设置曲线样式。color 用 Origin 颜色索引名(red/blue/...)或 1-24 数字；
    line_width_pt 单位 pt。带误差棒时请用 y_col（Y 列号）定位曲线，
    series_index 在含误差棒的图中会错位。"""
    session.ensure_connected()
    return session.series_style(graph=graph, series_index=series_index,
                                color=color, line_width_pt=line_width_pt,
                                symbol_size=symbol_size,
                                symbol_shape_index=symbol_shape_index,
                                y_col=y_col)


@mcp.tool()
def text_label(text: str, name: str = "mcpLabel", x: float | None = None,
               y: float | None = None, font_size: int | None = None,
               graph: str | None = None) -> dict:
    """向图形添加文本标注（如 "(a) Sample A"）。不给坐标时附到页面上部居中。"""
    session.ensure_connected()
    return session.text_label(text=text, name=name, x=x, y=y,
                              font_size=font_size, graph=graph)


@mcp.tool()
def graph_export(filepath: str, fmt: str | None = None, graph: str | None = None) -> dict:
    """导出图为图片，格式由扩展名或 fmt 决定：png/pdf/eps/tif/jpg/emf。
    分辨率取 Origin 默认导出设置（通常 300dpi 级别）。返回前校验文件真实生成。"""
    session.ensure_connected()
    return session.graph_export(filepath=filepath, fmt=fmt, graph=graph)


@mcp.tool()
def worksheet_export_csv(filepath: str, book: str | None = None) -> dict:
    """把工作簿活动表导出为 CSV 文件。"""
    session.ensure_connected()
    return session.worksheet_export_csv(filepath=filepath, book=book)


@mcp.tool()
def pages_list() -> dict:
    """列出当前打开的工作簿与图形窗口名。"""
    session.ensure_connected()
    return session.pages_list()


@mcp.tool()
def page_activate(name: str) -> dict:
    """按窗口名激活工作簿或图形页面，返回值含激活结果校验。"""
    session.ensure_connected()
    return session.page_activate(name=name)


@mcp.tool()
def project_save(path: str, force: bool = False, backup: bool = True) -> dict:
    """保存工程为 .opju/.opj（路径需为绝对路径）。返回值含文件存在性与大小校验。

    内置防误覆盖：当前工程枚举到 0 个页面但目标文件已存在且明显非空时，
    默认拒绝覆盖（多半是 COM 会话假死），需显式 force=True。
    backup=True 时覆盖前会先备份为 <原名>.bak，若保存后体积正常则自动删除备份。"""
    session.ensure_connected()
    return session.project_save(path=path, force=force, backup=backup)


@mcp.tool()
def project_open(path: str) -> dict:
    """打开已有 Origin 工程(.opju/.opj)，path 需为绝对路径。"""
    session.ensure_connected()
    return session.project_open(path=path)


@mcp.tool()
def graph_frame(graph: str, boxed: bool = True) -> dict:
    """显示/隐藏图形边框（补全上/右轴线），boxed=true 为封闭边框。"""
    session.ensure_connected()
    return session.graph_frame(graph=graph, boxed=boxed)


@mcp.tool()
def add_ref_line(graph: str, axis: str = "x", pos: float = 0.0,
                 color: int = 1, width: float = 1.0, dash: bool = True,
                 label: str | None = None, book: str | None = None,
                 span: list[float] | None = None) -> dict:
    """添加参考线：axis='x' 为竖线 x=pos，axis='y' 为横线 y=pos。
    走 Origin 原生参考线对象 layer.{axis}.refline#（2018 SR1+），
    不进图例、不触发轴缩放、不写数据表。
    color 用 Origin 颜色索引(1黑 2红 3绿 4蓝 14灰...)；dash=true 为虚线。
    label 给定时显示原生标签。book/span 已废弃，仅为兼容保留、会被忽略。"""
    session.ensure_connected()
    return session.add_ref_line(graph=graph, axis=axis, pos=pos,
                                color=color, width=width, dash=dash,
                                label=label, book=book, span=span)


@mcp.tool()
def legend_remove_last(graph: str, count: int = 1) -> dict:
    """删除图例最后 count 行（多余曲线/误差棒分组等条目）。
    必须在所有绘图与坐标轴操作全部完成之后调用，否则会被 Origin 重建恢复。
    参考线已走原生 refline，不会进图例，无需此清理。"""
    session.ensure_connected()
    return session.legend_remove_last(graph=graph, count=count)


@mcp.tool()
def labtalk_execute(script: str) -> dict:
    """直接执行任意 LabTalk 脚本（高级逃生舱）。只返回成功/失败；
    要读回数值请用 labtalk_evaluate。危险命令请优先用专用工具。"""
    session.ensure_connected()
    return session.labtalk_execute(script=script)


@mcp.tool()
def labtalk_evaluate(expr: str, window: str | None = None) -> dict:
    """求值 LabTalk 表达式并返回数值/字符串结果。例："10*3"、"@V"(版本)。
    window 可选：求值前先 win -a 到该窗口。读 layer.* 时**必须**给 window，
    否则读的是当前活动窗口的图层（工作簿活动时会得到垃圾值或 NANUM 哨兵）。
    读坐标轴范围请优先用 axis_get。"""
    session.ensure_connected()
    return session.labtalk_evaluate(expr=expr, window=window)


@mcp.tool()
def quit_origin(force: bool = False) -> dict:
    """退出 Origin 程序。默认拒绝以防丢失未保存工作，须显式 force=true 且征得用户同意。"""
    session.ensure_connected()
    return session.quit_origin(force=force)


def main() -> None:
    log.info("Origin MCP server 启动 (stdio)")
    mcp.run()


if __name__ == "__main__":
    main()
