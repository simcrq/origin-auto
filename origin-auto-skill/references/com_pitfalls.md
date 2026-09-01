# Origin COM 自动化坑点大全（实测版）

> 来源：origin-mcp 在 Origin 2021 (9.85) 上的系统实测，并融合
> codex-origin-plot 项目的历史实战记录。写任何 Origin 自动化代码前先读一遍。

## 一图速查

| 操作 | 结论 | 说明 |
|------|------|------|
| 连接 | EnsureDispatch("Origin.ApplicationSI") | SI=单实例；gencache 后才有完整方法集 |
| 新建工作簿 | newbook; + win -r %H 名字 | name:= 参数经 COM 不生效 |
| 批量写数据 | app.PutWorksheet(书名, 二维元组, r1, c1) | r1/c1 从 0 起，比逐格快几个量级 |
| 读数据 | app.GetWorksheet(书名) | 返回二维元组，可整表读回 |
| 连续多Y绘图 | plotxy iy:=(1,(2:4)) plot:=202 ogl:=[<new template:=LineSymb>] | 区间列号必须连续 |
| 追加曲线到已有图层 | 无可靠命令 | ogl:=[G]1、layer -i 均无效或另建新图 |
| 导出图片 | 裸文件名 expGraph → UFF → Python 复制 | 见下文【导出】 |
| 保存工程 | app.Save(反斜杠绝对路径) | 正斜杠静默失败 |
| 回读状态 | Evaluate / LTStr / LTVar | Execute 只有 bool |

## 详细说明

### 1. 静默失败三兄弟（最浪费时间的问题）

以下写法 Execute 返回 True 但什么都没发生：

- expGraph + 绝对路径（正反斜杠都一样）
- expGraph + 尺寸参数（width/height/tr1.unit/tr1.dpi 任一子句）
- LabTalk save + 正斜杠路径

唯一稳定方案：

- 导出：激活目标图后用裸文件名导出（落在用户文件夹 UFF，
  用 app.LTStr("%Y") 取路径），Python 端复制到目标位置。
  默认分辨率约 3216x2461，出版级够用。
  实测可用格式：png / pdf / eps / tif / emf；svg 经 COM 导出不可用。
- 保存：COM 方法 app.Save() + 反斜杠绝对路径。
- 每次操作后用 os.path.exists() 校验，不要相信返回值。

### 2. 输出捕获

Execute() 永远只给 bool。读状态只有三条路：

```python
app.Evaluate("layer.x.from")   # 数值表达式 -> float
app.LTStr("%Y")                # 字符串寄存器/变量
app.LTVar("某数值变量")
```

LabTalk 的 type 重定向到文件在 COM Execute 下不工作。

### 3. COM 结构细节

- 所有 COM 调用必须在同一线程：专用 worker 线程 + 任务队列，
  线程内一次性 CoInitialize。跨线程直调代理会报"接口为另一线程整理"。
- 页面枚举统一走 app.Pages（TypeName 区分 WorksheetPage/GraphPage）；
  WorksheetPages / GraphPages 子集合偶发 Item 访问失败。
- 集合 Item(i) 从 0 开始数，Count 是个数：循环 range(0, Count)。
- FindWorksheet(name) 仅对工作簿有效；图形窗口验证用
  win -a 名字; 之后比对 ActivePage.Name。
- gencache.EnsureDispatch 会把类型库缓存到 %TEMP%\gen_py；
  换 Python 版本后首次调用会自动重建。

### 4. 绘图语法

- plot 类型 ID（2021 实测渲染确认）：200=折线, 201=散点, 202=点线, 203=柱状。
  不同版本可能漂移，可传整数覆盖映射表。
- **iy 区间语法是重灾区**：
  - `iy:=(1,2)` 正确 → B vs A 散点
  - `iy:=(1,2:3)` 正确 → B、C 都对 A 作图（多 Y 必须这样写）
  - `iy:=(1,(2:3))` 错误！内层括号会被解析成 A-vs-A、C-vs-C（Y 对自己作图）
- 列设计ation（type=x）与 plotxy 显式 iy 混用会改变语义导致曲线错乱，
  二选一：要么设 designation 后用 `plotxy iy:=((B),(C))` 类范围，
  要么不设 designation 直接 `iy:=(1,2)`。服务端默认后者。
- 模板名 LineSymb / Column 在 2021 可用；不确定时不传 template 更稳
  （默认 Graph 模板即可正常渲染）。

### 5. 会话生命周期

- Origin.ApplicationSI 复用已开实例；用户手动打开的 Origin 不会被破坏。
- 不要随意 taskkill Origin64.exe——可能丢用户未保存的工作。
- COM 启动的 Origin 在最后一个引用释放后可能保持运行（SI 特性），
  彻底退出需 app.Exit()（须征得用户同意）。

### 6. 历史记录补充（codex-origin-plot）

- merge_graph 可能返回 False → 用独立导出代替。
- doc -e 循环 + type 重定向枚举窗口不可行 → 用 Pages 集合。
- echo %H 恒为 False → 用 ActivePage.Name。

### 7. 样式与装饰（第二轮实测补充）

- **LabTalk set 命令完全无效**：`set %C -c` / `set 名字 -c` / range 后 set，
  全部返回 True 但不改任何样式。唯一可靠通道：
  COM `DataPlot.SetNumProp(属性名, 值)`：
  - `color`：只认 Origin 颜色索引 1-24（1黑 2红 3绿 4蓝 5青 6品红 7黄
    8深黄 14深蓝 16紫）；RGB 复合值写入成功但渲染为黑
  - `line.width`：单位 pt，默认 0.6
  - 误差棒模式下 DataPlots 会混入误差子图，按索引定位会错位，
    应枚举 `GetDatasetName()` 匹配 `书名_列字母`（服务端 series_style
    的 y_col 参数已封装）
- **边框**：`GraphLayer.SetNumProp('showframe', 1)` 补全上/右轴线。
- **参考线（原生，2018 SR1+）**：`draw -l` 坐标在 COM 下不可控（实测命名错乱、
  坐标丢失、样式设不上），两点数据图 hack 也已弃用。正确做法是用 Origin 原生
  `layer.{axis}.refline#` 对象（见 8.7 节），不进图例、不触发轴缩放、不写数据表。
- **误差棒**：列 designation **yerr=3**（实测确认），且必须用选区方式绘图：
  `win -a book; worksheet -s c1 1 c3 nrows; plotxy plot:=201 ogl:=[G]1;`
  - plotxy 显式 iy 语法不会自动关联误差列
  - 选区内未设 designation 的列会被画成普通 Y 线
  - 每对绘制前必须重新 `win -a book`——上一次 plotxy 会把活动窗口
    切到新图，导致后续 worksheet -s 选区失效（静默画错列）
  - X 列也要设 designation=x，否则 X 列会被当成数据线画出
- **图例**：
  - Legend 对象 COM `Text` 属性可读写（格式 `\l(n) %(n)` 每行一条）
  - 误差棒 plot 不占图例行
  - **Origin 会懒重建图例**（axis_set/导出等操作触发），提前裁剪会被
    恢复 → 裁剪必须是所有操作后的最后一步；且先关
    `GraphPage.LegendsAutoUpdate = 0`
- **data_put 的 wks.nrows 只能扩不能缩**：无脑赋值会把已有数据行截掉
  （服务端已修复为 if 只扩不缩）。

### 8. 第三轮实测补充（命名净化 / 图层读取 / 写入 / 稳定性）

> 全部结论来自 `origin-mcp/scripts/probe_origin_caps*.py` 在 Origin 2021
> (9.8002) 上的实测输出，不是文档推测。

#### 8.1 窗口名净化规则（实测表）

Origin 短窗口名 = **删掉所有非 ASCII 字母数字** + **截断到 13 字符**。

| 请求名 | 实际名 | 说明 |
|---|---|---|
| `Data_0.6` | `Data06` | 点被删 |
| `Force_Displacement_0.6` | `ForceDisplace` | 删符号后 17 字符 → 截断 13 |
| `My Book` | `MyBook` | 空格被删 |
| `A.B.C` | `ABC` | |
| `Book-1` | `Book1` | 连字符被删 |
| `Graph_0.6_Force` | `Graph06Force` | |
| `测试窗口名称很长很长` | `A` | 中文不可用，净化为空后回退成 `A` |

后果与对策：

- 用请求名去 `win -a` / `FindGraphLayer` 会**静默失败**。必须以回读值为准。
- `Force_Displacement_0.6` 与 `Force_Displacement_0.9` 净化后**同名**
  （都是 `ForceDisplace`）——净化才引入的重名风险。
- **重名会挂死**：`win -r %H X` 当 X 与已有窗口重名时，Origin 弹模态框，
  COM 调用阻塞 **366 秒**后返回失败，ActivePage 仍是旧名。
  因此唯一化必须在【净化后】的名字上做，且追加序号要给序号留位置
  （`s[:13-len(suffix)] + suffix`），否则截断后再次撞名。

#### 8.2 图层坐标轴：LabTalk vs COM

| 读法 | 活动窗口=图 | 活动窗口=工作簿 |
|---|---|---|
| `Evaluate("layer.x.from")` | 正确（-2.5） | **垃圾值**（实测 0.0 / 1.0 / 8.0 / 15.0） |
| `Evaluate("[G]1!layer.x.from")` | 正确 | 正确（跨窗口语法，无需激活） |
| `Evaluate("[G]layer.x.from")` | 失败 → NANUM `-1.23456789e-300` | 同 |
| **`gl.GetNumProp("x.from")`** | **正确** | **正确（首选）** |

- `GraphLayer.GetNumProp` 支持 `x./y.` + `from/to/inc/type`，大小写不敏感，
  **完全不依赖活动窗口**——这是唯一干净的通道。
- 类型库里 `GraphLayer` 没有 Axis 对象（已用 `Origin8.tlb` 枚举确认），
  只有 `GetNumProp/SetNumProp/Execute/DataPlots/GraphObjects`。
- `x.rescale` **读回不可靠**（设成 0 后读回是 8.0），但**写入有效**：
  设过 `rescale=0` 后 from/to 就不再被自动缩放改回去。
- `Evaluate()` **只接受表达式**：把 `win -a X;` 拼进去当语句求值会返回 `None`。
  必须先 `Execute("win -a X;")` 再 `Evaluate("layer.x.from")`。

#### 8.3 PutWorksheet

- **会自动扩行扩列**：向默认 32 行 × 2 列的空表写 120×3，之后
  `wks.nrows=120 / wks.ncols=3`。所以 `wks.nrows` 预检只是保险不是前提。
- **start_row 分块数据完全正确**：0/100/200 起始行分三次写，回读三个位置都对。
- 性能实测：`68804×20` 整块 **0.42s**；按 20000 行分块 **0.51s**。
  慢的从来不是 PutWorksheet，而是逐格兜底路径。
- 因此回退顺序固定为：整块 → （重连后）整块 → 分块 → 逐格（仅 ≤20000 单元格）。
  137 万单元格走逐格是不可接受的，应当直接报错并提示重连。

#### 8.4 追加数据图会触发轴重新缩放

`plotxy ... ogl:=[G]1` 往图层**追加数据图**（如第二条曲线）后，实测轴范围：

```
绘后        x=[-2.5, 12.5]  y=[-10, 60]
追加数据图   x=[-0.5,  5.5]  y=[-20, 70]   <-- 被改掉了
```

对策：追加后立刻写回原范围并锁定——

```labtalk
layer.x.from = x0; layer.x.to = x1; layer.x.rescale = 0;
layer.y.from = y0; layer.y.to = y1; layer.y.rescale = 0;
```

实测能完整锁回 `x=[-2.5,12.5] y=[-10,60]`。
（注：**参考线已改用原生 refline，不涉及此坑**；这里说的是追加数据曲线。）

#### 8.5 图例

- 追加的**数据图**会进入图例，条目文本是列名（表现为 `T, V, X...`）。
  实测追加 1 条后图例从 `\l(1) %(1)` 变成 `\l(1) %(1)\r\n\l(2) %(2)`。
- 这是 Origin 默认行为，不是 bug；要清掉只能在**所有操作完成后**
  调 `legend_remove_last(count=N)`，并先 `GraphPage.LegendsAutoUpdate = 0`。
- 想让图例显示自定义文字：给该列写长名（`col(C)[L]$ = "..."`），
  因为图例取的就是列长名。
- **原生 refline 参考线不进图例**（DataPlots 不变），无需清理。

#### 8.6 会话 / 工程稳定性

- 长 `PutWorksheet` 后 `Pages` 枚举偶发返回空（`pages_list` → `ok:False []`），
  但 Origin 里的工程其实还在 —— 这是 **COM 代理假死**，不是数据丢了。
  对策：`pages_list` 在"此前枚举到过页面 + 现在为空"时重连同一 Origin
  实例（SI 保证是同一个）后重试，10 秒内限流一次。
- **空工程覆盖是最贵的坑**：会话假死后 `project_save` 会用 0.6KB 空工程
  覆盖掉 6.9MB 的已有成果。三重防线：
  1. 当前 0 页面 + 目标文件 > 100KB → 默认拒绝（除非 `force=True`）；
  2. 覆盖前复制 `<name>.opju.bak`；
  3. 保存后体积 < 原体积一半 → 保留 `.bak` 并在返回值里告警。
  体积正常时 `.bak` 会自动删除，不留下垃圾文件。

#### 8.7 原生参考线（Origin 2018 SR1+，替换 hack 的最终方案）

`scripts/probe_refline_native.py` 实测确认，Origin 原生参考线对象在 COM 下**完全可用**：

```labtalk
layer.x.reflines.count = 1;       // 条数
layer.x.refline1.value = 5;       // 在 x=5 加竖线
layer.x.reflines.lineshow = 1;    // 显示
layer.x.refline1.linecolor = 2;   // 颜色索引（2=red, 4=blue）
layer.x.refline1.linestyle = 2;   // 线型（1=实线, 2=虚线）
layer.x.refline1.linethickness = 1.5;   // 线宽 pt
layer.x.refline1.labeltext$ = "yield";  // 标签文本
layer.x.refline1.labelshow = 1;
```

关键实测结论：

- `value` / `linecolor` / `linestyle` / `linethickness` / `labelshow` 写入后
  **读回正确**，可做写后校验（`layer.x.refline1.value` 用 Evaluate 读）。
- **追加多条**：先读 `layer.{axis}.reflines.count` 得 N，再 `count=N+1`、
  设 `refline{N+1}.value`，前面的线不会被覆盖。空槽位 `value` 读回是 `0.0`
  （不是 NANUM），所以**必须用 count 定位新索引，不能用 value 判空**。
- 相比「两点数据图」hack 的三胜：不进图例、不触发轴缩放、不写数据表。
- 前提：Origin 2018 SR1+（2021 满足）。`draw -l`（COM 下命名错乱、坐标丢失）
  与两点数据图 hack 均已弃用。
