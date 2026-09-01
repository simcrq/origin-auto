"""Live test for the set_axis_origin tool (RANGE-start semantics).

Two very different things — test only the intended one:
  - THIS tool: axis RANGE starts at (x0,y0) -> plot frame's lower-left
    corner IS the data origin. Values below the origin get clipped.
  - NOT this tool: atzero crossover (axis LINES drawn through 0).

Validates:
  1. default (0,0): both axes' `from` become 0.0, `to` unchanged
     (auto upper bound preserved); PNG differs from baseline
  2. explicit upper bound x1/y1 honoured
  3. clipping: data spanning negative Y emits a `warning`
  4. tool count = 28
"""
import os
import sys
import time
import hashlib
import subprocess

REPO = r"G:\USTB2024\origin-auto\origin-mcp"
sys.path.insert(0, REPO)
OUT = r"G:\Users\SIMCR\probe_out"
os.makedirs(OUT, exist_ok=True)


def md5(p): return hashlib.md5(open(p, "rb").read()).hexdigest()


def rng(s, g):
    """Read back the real axis range via COM (ground truth)."""
    def _r():
        gl = s.app.FindGraphLayer(g)
        return {f"{ax}.{p}": float(gl.GetNumProp(f"{ax}.{p}"))
                for ax in ("x", "y") for p in ("from", "to")}
    return s.worker.submit(_r)


def make_graph(s, name):
    """Asymmetric data: x 0..10, y -5..25 (spans negative on purpose)."""
    bk = s.workbook_new(name=name)["book"]
    data = [[float(x), 3.0 * x - 5.0] for x in range(0, 11)]
    s.data_put(bk, data)
    s.worksheet_set_columns(bk, [
        {"index": 1, "type": "x"}, {"index": 2, "type": "y"}])
    return s.plot_create(bk, x_col=1, y_cols=[2], graph_name=name + "G")["graph"]


def main():
    import origin_mcp_server as S
    s = S.session
    s.connect(visible=True)
    print("Origin", s.worker.submit(lambda: s.app.Evaluate("@V")))

    # ---- case 1: default (0,0), auto upper bound preserved ----
    g = make_graph(s, "AO1")
    before = rng(s, g)
    s.ex(f"win -a {g}; doc -uw; win -r; redraw; layer -r;")
    time.sleep(0.4)
    p0 = os.path.join(OUT, "t_axisorigin_1_before.png")
    s.graph_export(p0, graph=g)

    r = s.set_axis_origin(g)            # defaults: x0=0, y0=0
    print("[1] set_axis_origin() ->", {k: r[k] for k in ("ok", "after")})
    assert r["ok"] is True, r
    after = r["after"]
    assert after["x.from"] == 0.0, f"x.from should be 0, got {after['x.from']}"
    assert after["y.from"] == 0.0, f"y.from should be 0, got {after['y.from']}"
    # upper bound must be preserved (not collapsed / not auto-reverted)
    assert abs(after["x.to"] - before["x.to"]) < 1e-6, \
        f"x.to changed: {before['x.to']} -> {after['x.to']}"
    assert abs(after["y.to"] - before["y.to"]) < 1e-6, \
        f"y.to changed: {before['y.to']} -> {after['y.to']}"
    p1 = os.path.join(OUT, "t_axisorigin_1_after.png")
    s.graph_export(p1, graph=g)
    assert md5(p0) != md5(p1), "PNG should differ after origin change"
    print(f"    [PASS] (0,0): from 0/0, to preserved "
          f"(x {before['x.to']:g}, y {before['y.to']:g}), PNG differs")

    # ---- case 2: explicit upper bound ----
    g2 = make_graph(s, "AO2")
    r2 = s.set_axis_origin(g2, x0=0.0, y0=0.0, x1=10.0, y1=25.0)
    print("[2] set_axis_origin(x1=10, y1=25) ->", r2["after"])
    assert r2["ok"] is True, r2
    assert r2["after"]["x.to"] == 10.0, r2["after"]
    assert r2["after"]["y.to"] == 25.0, r2["after"]
    assert r2["after"]["x.from"] == 0.0 and r2["after"]["y.from"] == 0.0
    print("    [PASS] explicit x1/y1 honoured")

    # ---- case 3: clipping warning when data dips below the origin ----
    g3 = make_graph(s, "AO3")
    r3 = s.set_axis_origin(g3)          # data y spans -5..25 => would clip
    print("[3] clipping case -> warning:", r3.get("warning", "(none)"))
    assert r3["ok"] is True, r3
    assert "warning" in r3, "expected a clipping warning for negative data"
    print("    [PASS] clipping warning emitted")

    # ---- case 4: tool count ----
    import re
    src = open(os.path.join(REPO, "origin_mcp_server.py"), encoding="utf-8").read()
    n = len(re.findall(r"@mcp\.tool\(\)", src))
    print(f"[4] @mcp.tool() count = {n}")
    assert n == 28, f"expected 28, got {n}"
    print("    [PASS] tool count = 28")

    print("\nALL ASSERTIONS PASSED")
    s.invalidate()
    subprocess.run(["taskkill", "/IM", "Origin64.exe", "/F"], capture_output=True)


if __name__ == "__main__":
    main()
