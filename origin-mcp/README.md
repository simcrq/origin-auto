# origin-mcp — 让 AI Agent 自动化操作 Origin 科研绘图 · 使用手册

通过 **MCP (Model Context Protocol)** 把 OriginLab Origin 暴露为一组工具，
让 Claude / opencode / Codex 等 Agent 用自然语言完成：连接 Origin → 导入数据 →
绘图 → 坐标轴/样式/边框/参考线/误差棒 → 导出图片 → 保存 `.opju` 工程。

**已在 Origin 2021 (9.8002, Windows 10/11) 全链路实测**：基础端到端验收共 15 步，
并完成真实科研数据（CarbonFiber 力-位移曲线，15 个试件）的完整出图验证。

```
你："用 Origin 画 CarbonFiber 的力-位移曲线，出版级"
        │ 自然语言
        ▼
Agent（Claude / opencode / Codex）
        │ MCP 协议 (stdio)
        ▼
origin_mcp_server.py  ←—— 28 个工具
        │ COM 自动化
        ▼
Origin64.exe（本机 Origin，图形界面可见）
        │
        └─→ PNG / PDF / OPJU 文件（全部经过真实存在性校验）
```

---

## 目录

1. [环境要求](#1-环境要求)
2. [安装（三步）](#2-安装三步)
3. [注册到 Agent 客户端](#3-注册到-agent-客户端)
4. [快速验证](#4-快速验证)
5. [工具总表](#5-工具总表)
   - [并发、超时与恢复](#51-并发超时与恢复)
6. [工具详细说明](#6-工具详细说明)
7. [典型工作流（照抄即可）](#7-典型工作流照抄即可)
8. [速查表（颜色/线型/列类型/绘图类型）](#8-速查表)
9. [实测坑点全集](#9-实测坑点全集)
10. [故障排查](#10-故障排查)
11. [安全约定](#11-安全约定)

---

## 1. 环境要求

| 项目 | 要求 | 说明 |
|---|---|---|
| 操作系统 | Windows 10/11 | COM 自动化仅限 Windows |
| OriginLab Origin | 2017+ 推荐（2021 实测） | 需完成过一次手动启动与许可验证 |
| Python | 3.9+（3.12 实测） | 64 位 |
| Python 包 | `mcp`、`pywin32`、`openpyxl` | 见 requirements.txt；Excel 导入才会用到 openpyxl |
| Agent 客户端 | opencode / Claude Desktop / Codex CLI 任一 | 用于注册 MCP |

## 2. 安装（三步）

```powershell
# ① 进入目录，创建独立虚拟环境（推荐，不污染系统 Python）
cd G:\USTB2024\0_Project\origin-mcp
python -m venv .venv

# ② 安装依赖
.\.venv\Scripts\pip install -r requirements.txt

# ③ 环境自检
python scripts\diagnose.py              # 基础 8 项（不启动 Origin）
python scripts\diagnose.py --com        # 真实 COM 连接测试（会启动 Origin）
```

`diagnose.py` 全部 PASS 即安装成功。它会输出：

- Origin 安装路径（从 COM 注册表 CLSID 反查，含 32/64 位视图处理）
- COM ProgID 注册情况（`Origin.ApplicationSI` / `Origin.Application`）
- Origin 是否正在运行
- Python 依赖版本

## 3. 注册到 Agent 客户端

`command`/`args` 一律使用**绝对路径**；用了 venv 就指向 venv 内的 python.exe。

### 3.1 opencode

方式一（命令行一键注册）：

```powershell
opencode mcp add origin -- G:\USTB2024\0_Project\origin-mcp\.venv\Scripts\python.exe G:\USTB2024\0_Project\origin-mcp\origin_mcp_server.py
```

方式二（编辑 `~/.config/opencode/opencode.json` 或项目根 `opencode.json`）：

```json
{
  "mcp": {
    "origin": {
      "type": "local",
      "command": ["G:\\USTB2024\\0_Project\\origin-mcp\\.venv\\Scripts\\python.exe",
                   "G:\\USTB2024\\0_Project\\origin-mcp\\origin_mcp_server.py"],
      "enabled": true
    }
  }
}
```

### 3.2 Claude Desktop

编辑 `%APPDATA%\Claude\claude_desktop_config.json`：

```json
{
  "mcpServers": {
    "origin": {
      "command": "G:\\USTB2024\\0_Project\\origin-mcp\\.venv\\Scripts\\python.exe",
      "args": ["G:\\USTB2024\\0_Project\\origin-mcp\\origin_mcp_server.py"]
    }
  }
}
```

### 3.3 Codex CLI

编辑 `~/.codex/config.toml`：

```toml
[mcp_servers.origin]
command = 'G:\USTB2024\0_Project\origin-mcp\.venv\Scripts\python.exe'
args = ['G:\USTB2024\0_Project\origin-mcp\origin_mcp_server.py']
startup_timeout_sec = 120.0
```

> ⚠️ 修改任何客户端全局配置前先备份原文件。
> 未使用 venv 时，把 command 换成系统 python.exe 绝对路径即可。
> 注册后**重启客户端**使配置生效。

## 4. 快速验证

### 4.1 端到端自动验收

```powershell
python scripts\test_e2e.py
```

自动完成：连接 Origin → 新建工程 → 建工作簿 → 写入 13×3 数据 → 画双 Y 点线图
→ 设坐标轴 → 曲线样式 → 文本标注 → 导出 PNG → 保存 OPJU → 读回数据核对
→ LabTalk 求值。15 步全部 PASS 即部署成功。
输出文件在 `%TEMP%\origin_mcp_test\`。

### 4.2 在 Agent 里说一句话验证

注册并重启客户端后，对 Agent 说：

```text
用 origin_status 检查 Origin 环境，然后连接 Origin，
画一条正弦曲线（x=0~6，步长0.5），导出 PNG 到 D:\temp\test.png
```

Agent 会依次调用 `origin_connect` → `workbook_new` → `data_put` →
`plot_create` → `graph_export`，并报告文件绝对路径。

---

## 5. 工具总表（28 个）

| 类别 | 工具 | 一句话说明 |
|---|---|---|
| 环境 | `origin_status` | 检测安装路径/注册表/进程，不启动 Origin |
| 会话 | `origin_connect` | 连接（必要时启动）Origin，长驻会话 |
| 会话 | `origin_disconnect` | 断开 COM（不关闭 Origin） |
| 会话 | `quit_origin` | 退出 Origin（须 force=true，防丢数据） |
| 项目 | `project_new` | 新建空白工程（清空当前，先提醒用户） |
| 项目 | `project_save` | 保存 .opju/.opj；**空工程覆盖已存在的大文件时默认拒绝**并留 `.bak` |
| 项目 | `project_open` | 打开已有工程 |
| 页面 | `workbook_new` | 新建工作簿并命名（**返回 Origin 净化后的实际名**） |
| 页面 | `pages_list` | 列出所有工作簿/图形窗口名（空结果会自动重连重试一次） |
| 页面 | `page_activate` | 按名激活窗口（支持未净化名自动解析） |
| 数据 | `data_put` | 二维数组批量写入；整块 → 重连重试 → 分块 → 逐格（仅小表） |
| 数据 | `worksheet_get_data` | 整表读回（核对/分析） |
| 数据 | `worksheet_set_columns` | 列长名/单位/注释/类型(designation) |
| 数据 | `worksheet_export_csv` | 活动表导出 CSV |
| 数据 | `import_csv` | 领域无关地导入任意 CSV（表头→列名、自动数值强转、可指定 X/Y 列、溯源戳记）|
| 数据 | `import_dataset` | 按 manifest.json 批量导入多个 CSV 并写 `import_log.json` 审计链（领域无关、可溯源）|
| 绘图 | `plot_create` | 折线/散点/点线/柱状；支持误差棒；新增通用对齐（`offset_origin`/`shared_x_grid`）、`line_width`、`ref_step` 参考线（**返回实际图形名**）|
| 格式 | `axis_get` | **读回轴真实范围/增量/刻度类型**（COM 通道，首选） |
| 格式 | `axis_set` | 轴标题/范围/线性对数/刻度增量（给范围时**自动锁轴**） |
| 格式 | `set_axis_origin` | 让两条轴线在 (0,0) 相交（修自动边距导致的"原点在 (0,-2)"）|
| 格式 | `series_style` | 曲线颜色/线宽/符号（支持 y_col 定位；图不存在时返回可读错误） |
| 格式 | `graph_frame` | 边框开关（上/右轴线） |
| 格式 | `text_label` | 文本标注 |
| 装饰 | `add_ref_line` | 参考线（x=竖线 / y=横线，可虚线/标签）；**原生 refline，不进图例** |
| 装饰 | `legend_remove_last` | 删除图例末尾条目（清理多余曲线/误差棒条目；参考线无需此清理） |
| 输出 | `graph_export` | 导出 png/pdf/tif/eps/emf/jpg |
| 高级 | `labtalk_execute` | 执行任意 LabTalk（只返回 bool） |
| 高级 | `labtalk_evaluate` | 求值 LabTalk 表达式（读 `layer.*` 请传 `window`） |

> 所有涉及文件的工具（导出/保存/打开）都在返回前做 `os.path.exists`
> 校验——不要相信 Origin 的布尔返回值，相信工具的返回值。
>
> 涉及窗口名的工具都接受**未净化**的名字（如 `Data_0.6`），会自动解析到
> Origin 实际采用的 `Data06`。但推荐直接用 `workbook_new`/`plot_create`
> 返回的实际名。

### 5.1 并发、超时与恢复

Origin 是单实例 STA，COM 调用不会并行。除只读的 `origin_status` 外，服务端对
所有 MCP 工具实行单飞：同一时间只允许一个有状态调用进入 Origin。

- 并发请求返回 `{"ok": false, "busy": true, "not_executed": true, ...}`，表示
  本次请求没有入队、没有产生副作用；等待当前调用完成后再发起。
- 排队中的 COM 任务若超时，会在开始执行前取消；已经开始的 COM 调用无法安全
  取消，服务端会等待真实结果，不会先返回超时再在后台继续修改 Origin。
- 客户端或传输层先超时、报“连接关闭”时，不能据此判断 Origin 进程已经退出。
  重新连接 `ApplicationSI`，先调用 `pages_list` 枚举实际状态，再只补建缺失结果。
- 不要并发拆开 `page_activate`/`win -a` 与后续 `layer.*` 读取；活动窗口是全局状态。
- 不要盲目重试 `plot_create`、`add_ref_line`、`graph_export`、`project_save`，也不要
  用 `taskkill Origin64.exe` 处理超时。

---

## 6. 工具详细说明

### 6.1 `origin_status()` → 环境检测

无参数。返回 `{installed, exe_paths, registry_entries, running_processes, progids_registered}`。
**不启动 Origin**，随时可安全调用。开始任何任务前先调它。

### 6.2 `origin_connect(visible=True)` → 连接

- `visible`：是否显示 Origin 窗口。默认 True（推荐，用户能看到过程）。
- 首次连接需 10~40 秒（Origin 启动 + splash 加载）；之后复用同一实例，秒级。
- 使用 `Origin.ApplicationSI`（单实例类）：用户手动打开的 Origin 不会被重复启动。

### 6.3 `workbook_new(name=None, sheets=1)` → 新建工作簿

- `name`：窗口短名，建议纯 ASCII 字母数字。**Origin 会净化窗口名**：
  删掉所有非 ASCII 字母数字并截断到 **13 字符**。

  | 请求名 | 实际窗口名 |
  |---|---|
  | `Data_0.6` | `Data06` |
  | `Force_Displacement_0.6` | `ForceDisplace` |
  | `My Book` | `MyBook` |
  | `中文名` | `A`（不可用） |

- 返回的是**回读到的实际名**：
  `{"ok": true, "book": "Data06", "requested": "Data_0.6", "warning": "..."}`。
  **后续所有操作用 `book` 字段的名字**，用请求名 `win -a` 会静默失败。
- 重名自动加后缀。注意唯一化是在**净化后**的名字上做的：
  `Force_Displacement_0.6` 与 `Force_Displacement_0.9` 净化后同名，
  后者会自动变成 `ForceDisplac2`。
  ⚠️ 若让重名发生，`win -r` 会弹模态框并**挂死 COM 366 秒**（实测）。
- 内部实现：`newbook;` 后立即 `win -r %H 名字` 重命名（`name:=` 参数在 COM 下
  不生效，实测坑），再回读 `ActivePage.Name` 并用 Pages 集合验证。

### 6.4 `data_put(book, data, start_row=1, start_col=1, header=None)` → 写数据

- `data`：二维数组，数字与字符串可混排。`[["Disp","Force"],...]` 外层是行。
- `start_row/start_col`：从 **1** 开始。
- `header`：列长名称列表（图例标签来源），与列一一对应；传 `""` 跳过该列。
- 走 COM `PutWorksheet` 批量写入，实测 `68804×20` 约 **0.4 秒**。
  PutWorksheet 会**自动扩行扩列**，`wks.nrows` 预检只是保险。
- 回退顺序：`整块` → `重连后整块` → `按 20000 行分块` → `逐格（仅 ≤20000 单元格）`。
  大表**不会**掉进逐格路径（137 万格逐格写不可接受，会直接报错提示重连）。
  返回值里的 `method` 字段告诉你实际走了哪条路。
- ⚠️ 列数不够会自动扩；行数**只扩不缩**（缩会截掉已有数据，已修复）。

### 6.5 `worksheet_set_columns(book, columns)` → 列属性

`columns` 元素：`{"index": 列号(1起), "name": 长名, "units": 单位, "comments": 注释, "type": 类型}`

`type` 取值（实测确认）：

| type | 含义 | 用途 |
|---|---|---|
| `"x"` | X | 配合选区/误差棒绘图时标记 X 列 |
| `"y"` | Y | 默认即是，一般不用设 |
| `"yerr"` | Y 误差 | **误差棒必需**（值=3） |
| `"xerr"` | X 误差 | 横向误差棒（按文档推测，未实测） |
| `"label"` | 标签 | 文本列 |

⚠️ 坑：给列设了 `type="x"` 之后，就不要再用 `plotxy iy:=(1,2)` 显式指定 X 列，
二者混用会导致曲线错乱（见 §9）。

### 6.6 `plot_create(...)` → 绘图（核心工具）

```python
plot_create(
    book,                    # 工作簿名
    x_col=1,                 # 共用 X 列号（与 y_cols 搭配）
    y_cols=[2, 3],           # Y 列号列表：必须相邻（连续区间）
    plot_type="line_symbol", # line / scatter / line_symbol / column / bar 或整数ID
    template=None,           # Origin 模板名（可选，默认不传更稳）
    graph_name=None,         # 图形窗口名（自动去重）
    pairs=[[1,2],[4,5]],     # XY 对模式：各 X 独立时用（与 y_cols 二选一）
    err_cols=[3, None],      # 与 pairs 对应的误差列号（None=该对无误差棒）
    offset_origin=False,     # 各曲线减去首点，偏置到原点
    shared_x_grid=False,     # 多曲线插值到公共 X 网格
    line_width=None,         # 统一线宽（pt）
    ref_step=None,           # 按固定步长添加原生参考线
    ref_axis="x",          # 参考线所在轴
    ref_dash=True,
    ref_color=1,
    ref_width=1.0,
    grid_step=None,          # 公共 X 网格步长
)
```

**三种用法：**

| 场景 | 参数组合 | 说明 |
|---|---|---|
| 共用 X 的相邻多 Y | `x_col=1, y_cols=[2,3,4]` | 一次画多条（内部用 `(x,a:b)` 区间语法）|
| 各行独立的 XY 对 | `pairs=[[1,2],[3,4]]` | 每对独立 X 列（如多试件各自位移）|
| 带误差棒 | `pairs=[[1,2],[4,5]], err_cols=[3,6]` | 误差列需先设 `type="yerr"`；内部走选区绘图 |

返回值：`{"ok", "graph", "book", "curves_detected", ...}` —— graph 名与曲线数
均已验证真实存在。

**plot_type 与渲染对照（Origin 2021 实测）：**

| plot_type | 整数 ID | 渲染效果 |
|---|---|---|
| `line` | 200 | 折线 |
| `scatter` | 201 | 散点（只有点）|
| `line_symbol` | 202 | 点线图（线+符号）|
| `column` | 203 | 柱状图 |

### 6.7 `series_style(graph, series_index=1, color=None, line_width_pt=None, symbol_size=None, symbol_shape_index=None, y_col=None)` → 曲线样式

- `color`：颜色名（`red/blue/green/cyan/magenta/yellow/orange/black/navy/violet`）
  或 1~24 的索引数字。**不支持 R,G,B**（COM 通道 RGB 渲染为黑，实测）。
- `line_width_pt`：线宽，单位 pt，默认 0.6，论文推荐 1.5。
- `y_col`：**推荐**。按 Y 列号定位曲线（如 `y_col=2` 定位画 col2 的曲线）。
  带误差棒的图中 DataPlots 索引会错位，务必用 y_col。
- `symbol_size` / `symbol_shape_index`：符号大小 / 形状编号（1=方 2=圆…）。

### 6.8 `axis_set(graph, axis="x", title=None, vmin=None, vmax=None, scale=None, major_increment=None)` → 坐标轴

- `axis`：`"x"` 或 `"y"`
- `title`：轴标题（如 `"Load (N)"`）
- `vmin/vmax`：范围；`scale`：`"linear"` 或 `"log10"`
- `major_increment`：主刻度间隔

### 6.9 `graph_frame(graph, boxed=True)` → 边框

`boxed=True` 补全上/右轴线（封闭边框），`False` 恢复只显示左/下。
内部走 COM `SetNumProp('showframe')`——LabTalk 无等效命令。

### 6.10 `add_ref_line(graph, axis="x", pos=0.0, color=1, width=1.0, dash=True, label=None)` → 参考线

走 Origin 原生参考线对象 `layer.{axis}.refline#`（2018 SR1+，实测验证），
**不进图例、不触发轴缩放、不写数据表**：

- `axis="x"` → 在 `x=pos` 画**竖线**；`axis="y"` → 在 `y=pos` 画**横线**
- `color`：Origin 颜色索引（1=黑 2=红 14=深蓝 16=紫…见 §8）
- `dash=True`：虚线（论文参考线标准样式）
- `label`：给字符串则显示原生参考线标签（`labeltext$`），不给则无标签
- 多次调用会**追加**多条（内部读 `reflines.count` 定位新索引，不覆盖旧线）
- 旧参数 `book`/`span` 已废弃，仅为兼容保留、会被忽略

### 6.11 `legend_remove_last(graph, count=1)` → 图例清理

删除图例最后 `count` 行。**必须放在所有绘图与坐标轴操作之后调用**——
Origin 会在后续操作（axis_set、导出等）时重建图例，提前删会被恢复。
典型用途：删掉多余的辅助数据曲线、误差棒分组等图例条目。
（**参考线走原生 refline，不进图例，无需此清理。**）

### 6.12 `graph_export(filepath, fmt=None, graph=None)` → 导出图片

- 格式由扩展名决定：`.png/.pdf/.tif/.eps/.emf/.jpg`
- 分辨率取 Origin 默认导出设置（实测 PNG 约 3216×2461，出版级够用）
- 返回 `{"ok": true, "path": "绝对路径", "size": 字节数}` —— 文件已确认真实生成
- ⚠️ svg 在 Origin 2021 COM 下不可用
- ⚠️ 自定义宽高/DPI 的参数子句会导致静默失败，已从接口移除；
  如需自定义尺寸请在 Origin 图形导出对话框中预设默认值

### 6.13 `project_save(path)` / `project_open(path)` / `project_new()`

- `path` 必须**反斜杠绝对路径**（如 `G:\data\result.opju`），正斜杠静默失败
- 保存返回 `{"ok", "path", "size"}`，文件已确认落盘
- `project_new` 会丢弃当前工程未保存内容，调用前应征得用户同意

### 6.14 `labtalk_evaluate(expr, window=None)` → 读状态（重要）

唯一能从 Origin 拿回数值的通道：

```python
labtalk_evaluate("10*3")                    # → {"ok": true, "value": 30.0}
labtalk_evaluate("layer.x.from")            # 危险：活动窗口是工作簿时读到垃圾值/NANUM
labtalk_evaluate("layer.x.from", window="Graph1")  # → 正确（先 win -a 再求值）
```

> **读坐标轴范围请用 `axis_get(graph, axis)`**（COM `GetNumProp`，不依赖活动窗口），
> 而不是裸的 `labtalk_evaluate("layer.x.from")`——后者永远作用于【当前活动窗口】
> 的图层，活动窗口是工作簿时读到的是工作表图层（实测 0/1/8 或 NANUM 哨兵）。

### 6.15 `labtalk_execute(script)` → 逃生舱

执行任意 LabTalk 脚本，只返回成功/失败。用于本手册未覆盖的高级操作。
危险命令（save/exit/删除）请优先使用专用工具。

### 6.16 其余

- `pages_list()`：返回 `{"worksheets": [...], "graphs": [...]}`
- `page_activate(name)`：激活窗口并验证
- `worksheet_get_data(book)`：整表读回，`{"ok", "rows", "cols", "data"}`
- `worksheet_export_csv(filepath, book)`：导出 CSV
- `origin_disconnect()`：断开会话（Origin 保持运行）
- `quit_origin(force=False)`：默认拒绝，`force=true` 且征得用户同意才退出

---

### 6.17 `import_csv(path, book_name=None, x_col=1, y_cols=None, provenance=None)` → 通用 CSV 导入

领域无关地把任意 CSV 读入新工作簿：

- 表头作为列 long-name；数值列经 `_coerce_cell` 自动强转 + `data_put` 的
  `numerictype` 预格式化，**避免整列被存成文本**（早期 X 轴塌缩到 0–0.2 的根因修复）。
- `x_col` / `y_cols` 指定列 designation（不指定则除 `x_col` 外全当 Y）。
- `provenance` 字符串盖进工作簿 comment（`page.info(8)$`，best-effort）并随结果返回，
  用于溯源；缺省为 `source=<文件名>`。
- 返回 `{"ok", "book", "rows", "cols", "header", "provenance", "write_method"}`——
  `book` 是 Origin 净化后的实际窗口名，后续调用用它。
- ⚠️ **重数据处理（单位换算 / 平滑 / 拆分 / 插值）应在外部 Python 完成**，
  只把成品 CSV 交给本工具。MCP 只负责「导入 + 绘图」，不做领域计算。

### 6.18 `import_dataset(manifest_path)` → 批量导入 + 溯源

领域无关：按 `manifest.json` 批量导入多个 CSV，并在**同目录**写 `import_log.json`
审计链，使每个结果工作簿都可回溯到源文件与处理说明。

manifest 结构：

```json
{
  "name": "8.27 carbon",
  "description": "拉拔力-位移，单位已转 N，窗长500平滑",
  "items": [
    {"csv": "processed/depth06.csv", "book_name": "D06",
     "x_col": 1, "y_cols": [2], "provenance": "smooth window=500 step=1"},
    {"csv": "processed/depth09.csv", "book_name": "D09",
     "x_col": 1, "y_cols": [2], "provenance": "smooth window=500 step=1"}
  ]
}
```

- `csv` 路径**相对 manifest 目录**解析（也接受绝对路径）。
- 每个 `item` 调 `import_csv`；`provenance` 缺省时回退到 manifest 的 `description`。
- 完成后写 `import_log.json` =
  `{"dataset": <name>, "log": {<book>: {"source", "provenance", "imported_at"}}}`。
- 返回 `{"ok", "imported", "total", "results": [...], "import_log": <路径>}`。

### 6.19 `plot_create` 的通用对齐参数（可选，不替代外部处理）

`plot_create` 新增一组与领域无关的可选参数，解决「多曲线起点 / 量程不一致」时的
轻量展示对齐：

| 参数 | 作用 |
|---|---|
| `offset_origin` | 所有曲线偏置到 (0,0)（各减自身首点），便于不同试件 / 工况对齐起点 |
| `shared_x_grid` | 多曲线 X 范围不一致时，插值到公共 X 网格再画（否则 plotxy 共用 X 列会把轴拉崩）|
| `grid_step` | 公共网格分辨率（默认取最小间距或 0.01）|
| `line_width` | 统一设曲线线宽(pt)，省去逐条 `series_style` |
| `ref_step` | 沿 `ref_axis` 以该步长画参考虚线（如 `ref_step=2` → x=2,4,6…）|
| `ref_axis` / `ref_dash` / `ref_color` / `ref_width` | 参考线轴别、虚线、颜色索引、线宽 |

机制：当 `offset_origin or shared_x_grid`（且非 `pairs` 模式）时，内部先调
`_build_aligned_book`——把所有曲线按各自首点偏置到 (0,0)，再线性插值到公共 X 网格，
写入 `*_align` 新工作簿，然后递归对对齐簿绘图，自动套用 `line_width` 与 `ref_step`
参考线。返回里多带 `aligned_book`、`xmax`。

> ⚠️ 对齐只解决「起点 / 量程不一致」的**展示**对齐，**不替代**外部的数据处理
> （单位换算、平滑、降采样等仍应在 Python 侧完成，再 `import_csv` 进 Origin）。

### 6.20 `set_axis_origin(graph, x0=0.0, y0=0.0, x1=None, y1=None)` → 轴范围从 (0,0) 开始

把两条坐标轴的**范围起点**钉到 `(x0, y0)`（默认 `(0, 0)`），并锁定——
图框的左下角就是数据原点。这是"原点在 (0,0) 而不是 (0,-2)"的正解：
Origin 自动缩放会给轴加负边距（实测 Y 轴被推到 `-2`），本工具两条都
清零。不传 `x1/y1` 时上界保持 Origin 当前的自动值。

| 参数 | 含义 |
|---|---|
| `x0` / `y0` | 两条轴的起点，默认都是 0 |
| `x1` / `y1` | 可选上界；不传则保持当前自动上界（不被改） |

**⚠️ 与"轴线穿过 0"（atzero 交叉）是两件不同的事，常见混淆：**

| 想要什么 | 用什么 | 后果 |
|---|---|---|
| **范围**从 (0,0) 开始（图框左下角 = 原点） | `set_axis_origin` | 低于原点的数据**被裁** |
| 轴线**穿过** 0 而保留所有数据 | `labtalk_execute("win -a <图>; layer.x.atzero=1; layer.y.atzero=1; doc -uw; win -r; redraw; layer -r;")` | 视觉上像"多画两条 x=0, y=0 线" |

数据含负值、又不想裁时，请用第二行的 atzero 路径。本工具会在检测到
"当前起点低于目标原点"时返回 `warning` 提示裁剪风险，但**照常执行**。

**Origin 2021 属性真相**（探测踩坑记录）：

| 候选属性 | `GetNumProp` 读 | 能否用 |
|---|---|---|
| `x.cross` / `x.crossN` / `x.crossAt` | 全部返回 sentinel (`-1.23e-300`) | ❌ 不存在 |
| `x.zeroline` / `x.at` / `x.axispos` | 全部返回 sentinel | ❌ 不存在 |
| `x.position` | 0/1/2/3 整数 | ⚠️ 写入生效但视觉不动（缓存问题） |
| **`x.atzero`** | **0.0 / 1.0** | ✅ atzero 路径用，但**不是**本工具的实现 |

实测坑：
- **必须 `layer.<ax>.rescale = 0` 锁定**——否则 Origin 自动边距会在后续任何重绘
  时把手动范围改回去。
- **必须强刷** `doc -uw; win -r; redraw; layer -r;`——只设属性不刷，`graph_export`
  抓到的是旧帧（readback 看起来对、PNG 不变，极迷惑）。本工具已自动包含强刷。

```python
# 典型：让 X/Y 都从 0 开始，上界保持自动
set_axis_origin("Fig1")                  # 默认 (0,0)

# 数据范围已知时显式锁两端
set_axis_origin("Fig1", x0=0, y0=0, x1=10, y1=25)

# 数据含负值、不想被裁——改用 atzero 路径
labtalk_execute("win -a Fig1; layer.x.atzero=1; layer.y.atzero=1; "
                "doc -uw; win -r; redraw; layer -r;")
```

**返回**：`{ok, graph, x0, y0, before, after, warning?}`。`ok=True` 以
**读回** `x.from == x0` 且 `y.from == y0` 为唯一标准，`before/after` 是
完整的 `{x.from, x.to, y.from, y.to}` 范围快照，方便核对。

---

## 7. 典型工作流（照抄即可）

以下"伪代码"展示 Agent 实际的工具调用序列。**你只需要对 Agent 说自然语言**，
它会自己翻译成这些调用；列出调用序列是为了便于你检查/复现。

### 7.1 单条曲线：CSV → 出版级折线图

```text
用户："把 D:\data\run1.csv 画成 Origin 力-位移曲线，X=位移列，Y=力列，导出 PNG"
```

```python
origin_status()                                    # ① 环境确认
origin_connect(visible=True)                       # ② 连接
workbook_new(name="Run1")                          # ③ 得到 book 名
data_put(book, rows, header=["","Force"])          # ④ 写数据（长名→图例）
plot_create(book, x_col=1, y_cols=[2],             # ⑤ 画点线图
            plot_type="line_symbol", graph_name="Fig1")
axis_set("Fig1", "x", title="Displacement (mm)", vmin=0)
axis_set("Fig1", "y", title="Load (N)")
series_style("Fig1", y_col=2, color=4, line_width_pt=1.5)
graph_frame("Fig1", boxed=True)                    # ⑥ 边框
graph_export(r"D:\out\fig1.png", graph="Fig1")     # ⑦ 导出（已校验）
project_save(r"D:\out\fig1.opju")                  # ⑧ 存档
```

### 7.2 多试件对比图（各自独立 X 列）

各试件位移范围不同时，用 `pairs` 模式（布局：每试件占两列 [X_i, Y_i]）：

```python
workbook_new(name="Group06")                       # 列布局: D1,F1,D2,F2,...
data_put(book, rows_14cols, header=["","试件1","","试件2",...])
plot_create(book, pairs=[[1,2],[3,4],[5,6],...],
            plot_type="line", graph_name="G06")
# 逐条上色（颜色索引见 §8）
series_style("G06", y_col=2, color=4, line_width_pt=1.5)
series_style("G06", y_col=4, color=2, line_width_pt=1.5)
series_style("G06", y_col=6, color=3, line_width_pt=1.5)
# ...
axis_set("G06", "x", title="Displacement (mm)", vmin=0)
axis_set("G06", "y", title="Load (N)")
graph_frame("G06", boxed=True)
graph_export(r"D:\out\group06.png", graph="G06")
```

### 7.3 均值 ± 标准差曲线（误差棒）

先把每组试件插值到公共位移网格，计算 mean/sd（Python 侧完成），然后：

```python
workbook_new(name="Mean06")                        # 列: Disp, Mean, SD
data_put(book, [[grid,mean,sd] 行...], header=["Disp","Mean ± SD",""])
worksheet_set_columns(book, [{"index":1,"type":"x"},
                             {"index":3,"type":"yerr"}])   # 误差列必须设 yerr
plot_create(book, pairs=[[1,2]], err_cols=[3],
            plot_type="line", graph_name="Mean06")
series_style("Mean06", y_col=2, color=1, line_width_pt=1.5)
graph_frame("Mean06", boxed=True)
axis_set("Mean06", "x", title="Displacement (mm)", vmin=0)
axis_set("Mean06", "y", title="Load (N)")
graph_export(r"D:\out\mean06.png", graph="Mean06")
```

误差棒太密时，把 SD 列大部分单元格留空（`""`=缺失=不画），
每 0.5mm 保留一个即可（CarbonFiber 实战采用此法）。

### 7.4 参考虚线

```python
# ...绘图与格式化完成前后均可调用（原生 refline 不触发轴缩放、不进图例）：
add_ref_line("Fig1", axis="x", pos=3.25, color=1, dash=True)   # x=3.25 竖虚线
add_ref_line("Fig1", axis="y", pos=5.0,  color=2, dash=True)   # y=5 横虚线
graph_export(...)
```

参考线走 Origin 原生 `layer.{axis}.refline#`，**不进图例**，无需
`legend_remove_last`；多次调用自动追加多条，不会覆盖。

### 7.5 修改已有工程

```text
用户："打开 D:\result.opju，把 Fig1 的 Y 轴标题改成 Stress (MPa)，重新导出"
```

```python
project_open(r"D:\result.opju")
page_activate("Fig1")                              # 或 pages_list() 查看名字
axis_set("Fig1", "y", title="Stress (MPa)")
graph_export(r"D:\out\fig1_new.png", graph="Fig1")
project_save(r"D:\result.opju")                    # 覆盖前确认用户同意
```

### 7.6 读回数据 / 检查 Origin 状态

```python
worksheet_get_data("Data1")            # 整表数据，核对写入结果
axis_get("Fig1", "x")                  # 读 X 轴范围（首选，COM 通道）
labtalk_evaluate("wks.ncols", window="Data1")  # 工作簿列数（带 window）
pages_list()                           # 当前打开的所有窗口
```

### 7.7 数据集批量导入 + 对齐绘图（推荐：重处理外部，导入仅绘图）

**原则**：单位换算 / 平滑 / 拆分 / 插值全部在外部 Python 完成，产出成品 CSV；
MCP 只负责把 CSV 导入 Origin 并绘图。这样处理过程可追溯、可复跑。

```text
用户："按 manifest.json 把 8.27 碳的三组深度数据导入 Origin，
       偏置到起点对齐画点线图，线宽 1.5，每 2 单位画参考虚线"
```

```python
origin_status()                                          # ① 环境确认
origin_connect(visible=True)                             # ② 连接
import_dataset("G:/8.27碳/processed/manifest.json")       # ③ 批量导入 + 写 import_log.json
#   → 返回各 book 名（D06 / D09 / D12）与溯源日志路径
plot_create("D06", x_col=1, y_cols=[2],                  # ④ 对齐绘图
            plot_type="line_symbol", graph_name="G06",
            offset_origin=True, shared_x_grid=True,
            line_width=1.5, ref_step=2)
plot_create("D09", x_col=1, y_cols=[2],
            plot_type="line_symbol", graph_name="G09",
            offset_origin=True, shared_x_grid=True,
            line_width=1.5, ref_step=2)
# ... 逐组导出 / 保存
graph_export(r"G:\out\g06.png", graph="G06")
project_save(r"G:\out\carbon_827.opju", backup=True)
```

`import_log.json` 落盘后即可复盘：每个工作簿 ← 哪个源 CSV ← 什么处理说明（provenance）。

---

## 8. 速查表

### 8.1 颜色索引（COM 通道只认索引，实测渲染）

| 索引 | 颜色 | | 索引 | 颜色 |
|---|---|---|---|---|
| 1 | 黑 | | 6 | 品红 |
| 2 | 红 | | 7 | 黄（白底慎用）|
| 3 | 绿 | | 8 | 深黄/橙 |
| 4 | 蓝 | | 14 | 深蓝 |
| 5 | 青 | | 16 | 紫 |

> 灰色的索引值未确认，参考线建议用黑色(1)。
> 7 条曲线推荐配色：`[4, 2, 3, 6, 5, 8, 1]`（蓝红绿品红青橙黑）。

### 8.2 绘图类型（plot_type）

| 名称 | ID | 效果 |
|---|---|---|
| `line` | 200 | 折线 |
| `scatter` | 201 | 散点 |
| `line_symbol` | 202 | 点线 |
| `column` | 203 | 柱状 |

### 8.3 列类型（worksheet_set_columns 的 type）

| 名称 | 代码 | 说明 |
|---|---|---|
| `x` | 4 | X 列 |
| `y` | 1 | Y 列（默认）|
| `yerr` | 3 | Y 误差列（误差棒必需，实测）|
| `xerr` | 5 | X 误差列（文档推测，未实测）|

### 8.4 导出格式

`png`（默认推荐）· `pdf`（矢量，投稿）· `tif`（期刊常要 300dpi+）·
`eps`（LaTeX）· `emf` · `jpg`。**svg 不可用**（2021 COM 限制）。

---

## 9. 实测坑点全集

这些结论决定了服务端实现，**自己写 LabTalk 前必读**：

1. **expGraph 静默失败**：绝对路径或 width/height/tr1.* 尺寸子句 → 返回 True
   却无文件。方案 = 裸文件名导出到用户文件夹（UFF）再复制（工具已内置）。
2. **app.Save 只认反斜杠绝对路径**；LabTalk `save` 可能抛 COM 异常。
3. **newbook name:= 不生效** → newbook 后 `win -r %H 名` 重命名并验证。
4. **iy 语法**：`(1,(2:3))` 内层括号 = Y 对自己作图（错）；正确 `(1,2:3)`。
5. **Execute 只有 bool**：读状态用 `labtalk_evaluate`。
6. **COM 集合 Item() 从 0 开始**；枚举统一走 Pages 集合 + TypeName。
7. **LabTalk `set` 样式命令完全无效**（返回 True 不生效）→ 样式走 COM
   `DataPlot.SetNumProp`；color 只认索引，RGB 渲染为黑。
8. **误差棒**：yerr designation=3 + 选区绘图（`worksheet -s` + 无 iy 的
   plotxy）；每对绘制前必须重新激活工作簿；X 列也要设 designation。
9. **图例懒重建**：裁剪图例必须放在所有操作之后，且先关
   `LegendsAutoUpdate`（工具已封装）。
10. **data_put 行数只扩不缩**（无脑赋值会截掉已有数据，已修复）。
11. **参考线**：`draw -l` 坐标在 COM 下不可控（命名错乱、坐标丢失）；
    两点数据图 hack 会进图例 + 触发轴缩放。现已改用原生 `layer.{axis}.refline#`
    对象（不进图例、不缩放、不写数据表），见 §6.10。
12. svg 经 COM 导出不可用。

完整细节与更多案例：`../origin-auto-skill/references/com_pitfalls.md`

---

## 10. 故障排查

| 现象 | 原因与处理 |
|---|---|
| COM 连接失败 | 手动启动一次 Origin 完成首次运行/许可验证后关闭，重试 `diagnose.py --com` |
| 检测不到安装路径 | 服务端从 CLSID LocalServer32 反查（含 WOW6432Node）；确认 Origin 未被卸载残留 |
| 导出/保存"成功但无文件" | 不要自己写 LabTalk 导出/保存，一律走 `graph_export`/`project_save` |
| 曲线数据错乱 | 列设了 `type=x` 又用显式 iy 指定 X —— 二选一 |
| 误差棒不出现 | 误差列没设 `type="yerr"`，或没用 pairs+err_cols 模式 |
| 画出了多余的线 | X 列未设 designation，被当成数据线；给 X 列设 `type="x"` |
| 图例删了又出现 | 裁剪时机不对；`legend_remove_last` 必须在所有操作后最后调用 |
| 曲线样式不生效 | 用 `series_style` 的 `y_col` 参数定位；不要自己写 `set` 命令 |
| MCP 启动超时 | 首次 gencache 生成类型库缓存较慢，第二次启动即恢复 |
| 返回 `busy=true, not_executed=true` | 前一有状态调用仍在执行；本次未执行，等待后再调用 |
| 调用超时或“连接关闭”但 Origin 仍在 | 重连同一 `ApplicationSI`，先 `pages_list` 核对现状，只补缺失结果，勿盲目重试或杀进程 |
| Agent 找不到工具 | 确认注册后重启了客户端；`opencode` 里用 `/mcp` 查看状态 |

---

## 11. 安全约定

- 默认**不 kill** 已运行的 Origin（防止丢失用户未保存工作）；
  `origin_connect` 使用 SI 单实例类，直接附着到已开实例
- `quit_origin` 默认拒绝执行，须 `force=true` 且征得用户同意
- `project_new` 会清空当前工程——Agent 调用前应先提醒用户保存
- `project_save` 覆盖已有工程前默认创建临时 `.bak`；若当前工程为空或保存后体积
  异常缩水则拒绝/报警并保留备份。正常覆盖仍应先取得用户授权
- 所有有状态 MCP 工具串行调用；收到 `busy=true` 时不得并发重试

---

## 配套 Skill

`../origin-auto-skill/` 是自然语言触发层：告诉 Agent "用Origin画图" 即自动
选择 MCP 工具链；MCP 未注册时回退到独立脚本 `quick_plot.py`。
其 `references/com_pitfalls.md` 是本文档 §9 的完整版。

## 实战案例

`G:\USTB2024\0_Project\CarbonFiber\figures\` —— 15 个试件的力-位移曲线
（分组对比图 + 均值±SD 误差棒图 + 参考虚线 + 边框），全部由本 MCP 工具链
自动生成，OPJU 工程已保存可直接继续编辑。
