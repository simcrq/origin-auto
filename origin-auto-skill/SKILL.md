---
name: origin-auto-skill
description: >-
  Drive OriginLab Origin via COM automation to create publication-grade
  scientific figures from CSV/Excel data. Use when the user asks for Origin
  plots (.opju source files), says "用Origin画图", "Origin绘图",
  "科研绘图", "出版级图表", or needs journal-submission figures that must NOT
  be made with matplotlib. Supports line/scatter/line_symbol/column plots,
  log axes, multi-Y series, Chinese/English labels, reference lines,
  and exports PNG/PDF/TIF/EPS plus .opju project saving.
---

# origin-auto-skill — Agent 自动化操作 Origin 科研绘图

> 让 AI Agent 通过自然语言驱动本机 Origin：导入数据 → 绘图 → 格式化 → 导出/存档，
> 全程文件真实存在性校验。已在 Origin 2021 (9.8002) 全链路实测；
> 脚本回退通道由 `scripts/test_session.py` 覆盖。

## 何时使用

- 用户明确要求 Origin 图表（投稿需要 .opju 源文件或审稿人要求 Origin 原生图）
- 用户说"用Origin画"、"Origin绘图"、"科研绘图"、"出版级图表"
- 数据在 CSV/Excel 中，需要折线/散点/点线/柱状图，含中文标签
- **不要**用于：示意草图、快速探索（matplotlib 更快）；3D/等高线等复杂图先用 labtalk_execute 尝试

## 执行前决策树

```
1. 本会话是否注册了 origin-mcp？
   ├─ 是 → 直接按【MCP 工作流】调用工具（首选，长驻会话最快）
   └─ 否/不确定 → 先试 origin_status 工具；
        失败则走【脚本回退】：python scripts/quick_plot.py
2. 脚本也失败（无 pywin32/无 Origin）？
   → 告知用户无法生成 Origin 原生图，询问是否接受 matplotlib 替代
```

`scripts/quick_plot.py` 与 MCP 走同一套 COM 封装（`scripts/origin_session.py`），
已用 12977 行真实 CSV 实测通过，不是未验证的占位代码。

## 并发与超时（必须串行）

Origin 是单实例 STA，COM 调用不会并行。**同一会话一次只调用一个有状态工具**，
不要并发发送绘图、激活窗口、参考线、导出或保存请求，也不要把 `win -a` 与后续
`layer.*` 读取拆成可交错的并行工具调用。

- MCP 服务端实行单飞：并发请求返回 `busy=true, not_executed=true`，表示本次没有
  入队、没有副作用。等待当前调用完成后再继续。
- 排队任务超时会在执行前取消；已经开始的 COM 调用无法安全取消，服务端会等待
  真实结果，不会先报失败再在后台继续修改 Origin。
- 若客户端/传输先超时或报“连接关闭”，**不代表 Origin 进程或工程已丢失**。
  重连 `ApplicationSI`，先调用 `pages_list` 枚举现状，再只补建缺失结果；不要盲目
  重试 `plot_create`、`add_ref_line`、`graph_export`、`project_save` 等副作用操作。
- 不要用 `taskkill Origin64.exe` 处理并发超时；这会丢失未保存工程。

故障机理与恢复细节见 `references/com_pitfalls.md` 的“8.8 并发、超时与恢复”。

## MCP 工作流（标准 8 步）

```
0. origin_status          —— 确认已安装（不启动）
1. origin_connect         —— 连接/启动（首次约 10~40 秒）
2. workbook_new           —— 必须用返回的【实际】窗口名（Origin 会净化）
3. worksheet_set_columns  —— 列长名/单位（可选但推荐；勿设 type 与 plotxy 混用）
4. data_put               —— 二维数组一次写入（COM PutWorksheet，快）
5. plot_create            —— 返回的 graph 名已验证存在（同样可能是净化名）
   axis_get / axis_set / series_style / text_label
6. graph_frame / add_ref_line
7. legend_remove_last     —— 仅有多余图例条目时，最后调用（参考线不进图例）
8. graph_export + project_save —— 返回值 = 文件真实生成的凭证
```

### 工具速查

| 意图 | 工具 | 要点 |
|---|---|---|
| 查环境 | `origin_status` | 不启动 Origin |
| 连接 | `origin_connect` | 长驻会话，后续调用秒级 |
| 读状态 | `labtalk_evaluate` | 读 `layer.*` **必须传 window**，否则读到别的窗口 |
| **读轴范围** | **`axis_get`** | **首选通道**（COM GetNumProp），不依赖活动窗口 |
| 写数据 | `data_put` | start_row/start_col 从 1 起；大表自动分块 |
| 读数据 | `worksheet_get_data` | 核对写入结果 |
| 绘图 | `plot_create` | 多 Y 列需相邻；pairs+err_cols 可带误差棒 |
| 曲线样式 | `series_style` | 带误差棒时用 y_col 定位（索引会错位） |
| 设轴 | `axis_set` | 给了 vmin/vmax 会**自动锁轴**（rescale=0） |
| 边框 | `graph_frame` | 补全上/右轴线（封闭边框） |
| 参考线 | `add_ref_line` | x=竖线/y=横线，可虚线/标签；**原生 refline**（不进图例） |
| 图例清理 | `legend_remove_last` | 仅清理多余曲线/误差棒条目；**参考线不进图例** |
| 导出 | `graph_export` | png/pdf/tif/eps/emf；分辨率取 Origin 默认设置 |
| 存档 | `project_save` | 必须绝对路径；空工程覆盖已有大文件会被**拒绝** |

### 误差棒 / 边框 / 参考线 工作流

```
data_put（布局 [X,Y,err,X2,Y2,err2...]）
→ worksheet_set_columns（err 列 type=yerr，X 列 type=x）
→ plot_create(pairs=[[1,2],[4,5]], err_cols=[3,6])
→ series_style(y_col=2/5, ...)          # 用 y_col 定位
→ graph_frame(boxed=True)
→ add_ref_line(axis="x", pos=...)        # 原生参考线，不进图例
→ axis_set / text_label
→ graph_export / project_save
```

参考线走 Origin 原生 `layer.{axis}.refline#` 对象，**不进图例**（无需
`legend_remove_last`）、不触发轴缩放、不写数据表；`legend_remove_last`
仅在有多余数据曲线/误差棒分组条目时才用。

### 窗口命名：必须做的事

Origin 会把窗口短名**净化**：删掉所有非 ASCII 字母数字，再**截断到 13 字符**。

| 请求名 | 实际窗口名 |
|---|---|
| `Data_0.6` | `Data06` |
| `Data_0.9` | `Data09` |
| `Force_Displacement_0.6` | `ForceDisplace` |
| `My Book` | `MyBook` |
| `中文名` | `A`（不可用，勿用中文命名） |

- `workbook_new` / `plot_create` 的返回值已是**回读到的实际名**，后续一律用它。
- 万一传了未净化的名字（如 `Data_0.6`），工具会自动解析到 `Data06`，不会静默失败。
- **不要用会净化后重名的名字**：`Force_Displacement_0.6` 与 `Force_Displacement_0.9`
  净化后都是 `ForceDisplace`。重名会让 `win -r` 弹模态框并**挂死 COM 180~366 秒**。

## 脚本回退（未注册 MCP 时）

```powershell
python <skill目录>/scripts/quick_plot.py `
  --data "D:\data\spectrum.xlsx" --sheet "Sheet1" `
  --x-col 1 --y-cols 2,3 --plot-type line_symbol `
  --x-title "时间 (s)" --y-title "强度 (a.u.)" `
  --log-y --export "D:\out\fig1.png" --save "D:\out\fig1.opju"
```

脚本结束码 0=全部成功；输出中每步带 PASS/FAIL 与真实路径。
依赖：`pip install pywin32 openpyxl`。

底层封装 `scripts/origin_session.py` 与 MCP 服务端同源，同样具备：
窗口名净化回读、未净化名自动解析、大表分块写入、COM 轴回读、
空工程覆盖保护。自检脚本：`python scripts/test_session.py`。

## 关键坑点（实测 Origin 2021，务必遵守）

| # | 坑 | 正确做法 |
|---|---|---|
| 1 | `expGraph` 带绝对路径或 width/height/tr1.* 参数 → 返回 True 但**不生成文件** | 只用裸文件名导出到用户文件夹(UFF)，再复制到目标路径 |
| 2 | `app.Save("D:/x.opju")` 正斜杠 → 静默失败 | 反斜杠绝对路径 `D:\x\y.opju` |
| 3 | `newbook name:=X` 的 name 参数经 COM 不生效 | newbook 后立即 `win -r %H X` 重命名并**回读实际名** |
| 4 | Execute 只返回 bool，拿不到 LabTalk 输出 | 用 Evaluate(expr)/LTStr/LTVar 回读 |
| 5 | `iy:=(1,(2:3))` 内层括号被解析成 Y-vs-Y（A-vs-A、C-vs-C）；plot:=200=折线、201=散点、202=点线（渲染确认） | 多 Y 写 `(x,a:b)` 无内层括号；勿将列 designation(type=x) 与 plotxy 显式 iy 混用 |
| 6 | COM 集合 Item() 从 **0** 开始 | 枚举写 range(0, Count)，否则永远漏第一个元素 |
| 7 | WorksheetPages/GraphPages 子集合偶发 Item 访问失败 | 枚举统一走 Pages 集合 + TypeName 分类 |
| 8 | LabTalk `save` 可能触发 COM 崩溃(-2147417851) | 优先 app.Save()，LabTalk save 仅兜底且捕获 com_error |
| 9 | svg 经 COM 导出不可用（2021 实测） | 改用 png/pdf/emf/eps/tif，或告知用户手动导 SVG |
| 10 | 杀 Origin 进程会丢失用户未保存工作 | 默认不 kill；ApplicationSI 直接连接已有实例 |
| **11** | **窗口名被净化**（去非字母数字 + 截断 13 字符），用请求名 `win -a` 静默失败 | 始终使用返回值里的实际名；工具会自动解析别名 |
| **12** | **`win -r` 重名 → 弹模态框挂死 COM 180~366s** | 唯一化必须在净化后的名字上做（工具已实现） |
| **13** | **`layer.x.from` 读到垃圾值/NANUM**：它永远作用于【当前活动窗口】的图层，活动窗口是工作簿时读到工作表图层（实测 0/1/8） | 用 `axis_get`（COM `GraphLayer.GetNumProp`）；LabTalk 读 layer.* 必须传 window |
| **14** | **`axis_set(vmin=0)` 后仍出现自动边距**：Origin 会自动重缩放把范围改回去 | 设 vmin/vmax 时工具自动加 `layer.*.rescale=0` 锁定 |
| **15** | **参考线**：`draw -l` 在 COM 下命名错乱/坐标丢失；两点数据图 hack 会进图例 + 触发轴缩放 | 用原生 `add_ref_line`（`layer.{axis}.refline#`），不进图例、不缩放、不写数据表 |
| **16** | **会话假死导致空工程覆盖**：pages_list 返回 [] 时保存会用 0.6KB 空工程覆盖 6.9MB 成果 | `project_save` 默认拒绝；覆盖前留 `.bak`，体积骤降会告警 |
| **17** | **并发是假并发**：排队超时后重试会重复副作用，活动窗口交错会静默读错图层 | 一次只调用一个有状态工具；`busy=true` 表示未执行；外部超时后重连并先 pages_list 核对，勿盲目重放 |

完整版见 `references/com_pitfalls.md`。

## 安全约定

- `quit_origin` / 强制重启前必须征得用户同意
- `project_new` 会丢弃当前工程，先提醒用户保存
- 所有输出写到用户指定或当前工作目录，不覆盖已有文件名不加 `-f` 类参数
- `project_save` 遇到"当前工程 0 页面 + 目标文件已存在且较大"时会拒绝；
  确需覆盖才传 `force=True`，且原文件会自动备份为 `<name>.opju.bak`
