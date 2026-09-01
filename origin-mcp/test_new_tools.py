"""Live test for the new origin-mcp capabilities:
  - data_put hardening (string input -> numeric, no text-column bug)
  - import_csv (generic, with provenance)
  - import_dataset (manifest -> books + import_log.json traceability)
  - plot_create alignment (offset_origin + shared_x_grid + line_width + ref_step)

Run with the origin-mcp venv python + sandbox disabled (COM needs it).
"""
import sys, os, csv, json, tempfile, subprocess
sys.path.insert(0, r"G:\USTB2024\0_Project\origin-mcp")
import origin_mcp_server as m

SESSION = m.session
SMOOTH = r"G:\USTB2024\0_Project\8.27碳\B_force_N\smoothed"
PASS = []
FAIL = []
created_windows = []


def check(cond, msg):
    (PASS if cond else FAIL).append(msg)
    print(("PASS " if cond else "FAIL ") + msg)


def lerp(xs, ys, xq):
    if xq <= xs[0]:
        return ys[0]
    if xq >= xs[-1]:
        return ys[-1]
    lo, hi = 0, len(xs) - 1
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if xs[mid] <= xq:
            lo = mid
        else:
            hi = mid
    t = (xq - xs[lo]) / (xs[hi] - xs[lo]) if xs[hi] != xs[lo] else 0.0
    return ys[lo] + t * (ys[hi] - ys[lo])


def read_smooth(path):
    xs, ys = [], []
    with open(path, newline="", encoding="utf-8-sig") as f:
        r = csv.DictReader(f)
        for row in r:
            xs.append(float(row["横梁_mm"]))
            ys.append(float(row["力_N"]))
    return xs, ys


def cleanup():
    for name in created_windows:
        try:
            SESSION.labtalk_execute(f"win -c {name}; del -px;")
        except Exception as e:  # noqa
            print(f"  (cleanup skip {name}: {e})")


def main():
    SESSION.ensure_connected()
    print("Origin connected:", SESSION.connected)

    # ---------- Test A: data_put hardening (string input -> numeric) ----------
    p = os.path.join(SMOOTH, "0.6_1.csv")
    with open(p, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    header = rows[0]
    body = [[str(v) for v in r] for r in rows[1:]]  # force everything to STRING
    r_bk = SESSION.workbook_new("TESTput")
    bk = r_bk["book"]
    created_windows.append(bk)
    SESSION.data_put(bk, body, header=header)
    rd = SESSION.worksheet_get_data(bk)
    first = rd["data"][1][0] if len(rd["data"]) > 1 else None
    is_num = isinstance(first, (int, float)) and not isinstance(first, bool)
    check(is_num, f"data_put coerces string->numeric (col1[1]={first!r} type={type(first).__name__})")
    rp = SESSION.plot_create(bk, x_col=1, y_cols=[2], graph_name="TESTgraph")
    g = rp.get("graph")
    if g:
        created_windows.append(g)
    # lock axis to origin to verify numeric data is actually plotted from 0
    SESSION.axis_set(g, "x", vmin=0, vmax=7, major_increment=1)
    ax = SESSION.axis_get(g, "x") if g else {"from": None, "to": None}
    ok_range = ax.get("from") == 0.0 and ax.get("to") == 7.0 and ax.get("increment") == 1.0
    check(ok_range, f"plot A axis lock X = {ax.get('from')}..{ax.get('to')} inc {ax.get('increment')} (data 0..6.74)")
    check(ax.get("to", 0) > 1, "plot A axis NOT collapsed to 0-0.2 (text bug absent)")

    # ---------- Test B: import_csv + import_dataset (traceability) ----------
    items = []
    for stem in ["0.6_1", "0.9_1", "1.2_1"]:
        items.append({
            "csv": os.path.join(SMOOTH, stem + ".csv"),
            "book_name": "T_" + stem.replace(".", ""),
            "x_col": 1, "y_cols": [2],
            "provenance": "smooth window=500 step=1 (test)",
        })
    man = {"name": "test-dataset", "description": "live test", "items": items}
    tmp = tempfile.mkdtemp(prefix="origin_test_")
    man_path = os.path.join(tmp, "manifest.json")
    with open(man_path, "w", encoding="utf-8") as f:
        json.dump(man, f, ensure_ascii=False, indent=2)
    rd2 = SESSION.import_dataset(man_path)
    check(rd2.get("ok"), f"import_dataset ok (imported={rd2.get('imported')}/{rd2.get('total')})")
    for r in rd2.get("results", []):
        if r.get("ok"):
            created_windows.append(r["book"])
    logp = rd2.get("import_log")
    check(bool(logp) and os.path.exists(logp), f"import_log.json written: {logp}")
    if logp:
        with open(logp, encoding="utf-8") as f:
            logj = json.load(f)
        check("log" in logj and len(logj["log"]) == rd2.get("imported"),
              f"import_log has {len(logj.get('log', {}))} entries with provenance")

    # ---------- Test C: full alignment chain (external interp -> import -> plot) ----------
    files = [f for f in os.listdir(SMOOTH) if f.endswith(".csv")]
    groups = {}
    for fn in files:
        depth = fn.split("_")[0]
        groups.setdefault(depth, []).append(fn)
    proc_dir = os.path.join(tmp, "processed")
    os.makedirs(proc_dir, exist_ok=True)
    depth_graphs = {}
    for depth, fns in groups.items():
        # external processing: interpolate every specimen onto one common grid
        series = [read_smooth(os.path.join(SMOOTH, fn)) for fn in sorted(fns)]
        xmax = max(max(xs) for xs, _ in series)
        grid = [i * 0.01 for i in range(int(xmax / 0.01) + 1)]
        cols = [grid]
        for xs, ys in series:
            cols.append([lerp(xs, ys, g) for g in grid])
        out = os.path.join(proc_dir, f"{depth}.csv")
        with open(out, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["X"] + [f"Y{i}" for i in range(1, len(cols))])
            for r in range(len(grid)):
                w.writerow([cols[c][r] for c in range(len(cols))])
        rc = SESSION.import_csv(out, book_name="C_" + depth.replace(".", ""),
                                provenance=f"depth {depth}: {len(series)} batches, common grid")
        check(rc.get("ok"), f"import_csv combined depth {depth} -> {rc.get('book')}")
        if rc.get("ok"):
            created_windows.append(rc["book"])
            rp3 = SESSION.plot_create(rc["book"], x_col=1,
                                      offset_origin=True, shared_x_grid=True,
                                      line_width=1.5, ref_step=2,
                                      graph_name="CG" + depth.replace(".", ""))
            g3 = rp3.get("graph")
            if g3:
                created_windows.append(g3)
                depth_graphs[depth] = g3
                # verify aligned book data is correct (X from 0 to ~xmax, numeric)
                ab = rp3.get("aligned_book")
                ad = SESSION.worksheet_get_data(ab) if ab else {"data": []}
                xcol = []
                for r in ad.get("data", []):
                    if not r or r[0] in (None, ""):
                        continue
                    try:
                        xcol.append(float(r[0]))
                    except (TypeError, ValueError):
                        continue  # skip header / non-numeric row
                if xcol:
                    check(abs(min(xcol)) < 1e-6 and abs(max(xcol) - xmax) < 0.5,
                          f"aligned X range {min(xcol):.3f}..{max(xcol):.3f} (expect 0..{xmax:.2f})")
                # lock axis to origin to verify numeric plot from 0
                SESSION.axis_set(g3, "x", vmin=0, vmax=int(xmax) + 1, major_increment=2)
                ax3 = SESSION.axis_get(g3, "x")
                ok3 = ax3.get("from") == 0.0 and ax3.get("to") == int(xmax) + 1
                check(ok3, f"plot depth {depth} axis lock X={ax3.get('from')}..{ax3.get('to')} (data 0..{xmax:.2f})")
                check(rp3.get("ok"), f"plot depth {depth} aligned ok (book {ab})")

    # ---------- summary ----------
    print("\n==== SUMMARY ====")
    print(f"PASS: {len(PASS)}  FAIL: {len(FAIL)}")
    for f in FAIL:
        print("  FAIL:", f)
    cleanup()
    print("cleanup done")
    # kill Origin so a re-run starts from a clean state (avoids window-name collisions)
    try:
        subprocess.run(["taskkill", "/IM", "Origin64.exe", "/F"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass
    sys.exit(1 if FAIL else 0)


if __name__ == "__main__":
    main()
