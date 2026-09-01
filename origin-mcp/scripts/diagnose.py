# -*- coding: utf-8 -*-
"""
diagnose.py — Origin MCP 部署环境自检。

用法：
    python scripts/diagnose.py            # 基础检查（不启动 Origin）
    python scripts/diagnose.py --com      # 额外做一次真实 COM 连接测试
"""

import argparse
import importlib
import os
import sys

sys.path.insert(0, str(__import__("pathlib").Path(__file__).resolve().parents[1]))

CHECKS: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    CHECKS.append((name, ok, detail))
    mark = "PASS" if ok else "FAIL"
    print(f"[{mark}] {name}" + (f" — {detail}" if detail else ""))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--com", action="store_true", help="执行真实 COM 连接测试")
    args = ap.parse_args()

    print("=" * 60)
    print("Origin MCP 环境诊断")
    print("=" * 60)

    check("操作系统", sys.platform.startswith("win"), sys.platform)
    if not sys.platform.startswith("win"):
        return finish()

    v = sys.version_info
    check("Python >= 3.9", v >= (3, 9), f"{v.major}.{v.minor}.{v.micro}")

    for mod in ("mcp", "win32com.client", "pythoncom", "openpyxl"):
        try:
            m = importlib.import_module(mod)
            ver = getattr(m, "__version__", "") or getattr(m, "VERSION", "")
            check(f"依赖 {mod}", True, str(ver))
        except ImportError as exc:
            check(f"依赖 {mod}", False, f"{exc}（pip install {mod.split('.')[0]}）")

    from origin_mcp_server import detect_origin_install  # noqa: E402

    info = detect_origin_install()
    check("检测到 Origin 安装", info["installed"],
          f"exe: {info['exe_paths'] or '无'} | 注册表: {info['registry_entries'] or '无'}")
    check("COM ProgID 已注册", bool(info["progids_registered"]),
          ", ".join(info["progids_registered"]) or "未找到 Origin.Application(SI)")
    if info["running_processes"]:
        print(f"[INFO] Origin 正在运行: {info['running_processes']}")

    if args.com:
        print("-" * 60)
        print("执行真实 COM 连接测试（会启动/连接 Origin，约 5~30 秒）...")
        from origin_mcp_server import session  # noqa: E402

        try:
            r = session.connect(visible=False)
            print(f"[PASS] COM 连接成功 version={r.get('version')}")
            probe = session.ex('type -qs "";')
            print(f"[{'PASS' if probe else 'FAIL'}] LabTalk Execute 探测: {probe}")
            session.worker.stop()
        except Exception as exc:  # noqa: BLE001
            print(f"[FAIL] COM 连接失败: {exc}")
            print("       建议：手动启动一次 Origin 完成许可验证后再关闭重试。")

    return finish()


def finish() -> int:
    print("=" * 60)
    failed = [c for c in CHECKS if not c[1]]
    print(f"结果：{len(CHECKS) - len(failed)}/{len(CHECKS)} 通过")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
