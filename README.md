# origin-auto

用 AI Agent 自动化驱动 OriginLab Origin 做科研绘图 —— 把 Origin 暴露成
MCP 工具，再配一层 WorkBuddy 自然语言技能，让 "用 Origin 画 XX 曲线，出版级"
这种话直接变成可复现的出图。

本仓库含两个强耦合的组件：

| 目录 | 是什么 | 给谁用 |
|---|---|---|
| [`origin-mcp/`](./origin-mcp) | MCP 服务端（`origin_mcp_server.py`，28 个工具），通过 COM 自动化本机 Origin | 任意支持 MCP 的 Agent（Claude / opencode / Codex） |
| [`origin-auto-skill/`](./origin-auto-skill) | WorkBuddy 技能（自然语言触发层），未注册 MCP 时回退到独立脚本 `quick_plot.py` | WorkBuddy 用户 |

两者共享同一套 COM 自动化经验，完整坑点记录在
[`origin-auto-skill/references/com_pitfalls.md`](./origin-auto-skill/references/com_pitfalls.md)。

---

## 工作原理

```
你："用 Origin 画 CarbonFiber 的力-位移曲线，出版级"
        │ 自然语言
        ▼
Agent ── 选 origin-auto-skill 触发 MCP 工具链
        │ MCP 协议 (stdio)
        ▼
origin_mcp_server.py  ←—— 28 个工具（连接/数据/绘图/坐标轴/坐标轴原点/样式/参考线/导出/保存）
        │ COM 自动化
        ▼
Origin64.exe（本机 Origin，图形界面可见）
        └─→ PNG / PDF / OPJU（全部经过真实存在性校验）
```

---

## 组件一：origin-mcp（MCP 服务端）

把 Origin 暴露为 25 个 MCP 工具，端到端实测 30/30 通过，并完成真实科研数据
（CarbonFiber 15 试件力-位移曲线）的完整出图验证。

**环境要求**

- Windows 10/11（COM 自动化仅限 Windows）
- OriginLab Origin 2017+（2021 实测），需完成过一次手动启动与许可验证
- Python 3.9+（3.12 实测），包：`mcp`、`pywin32`（可选 `openpyxl`）

**安装与注册**

```powershell
cd origin-mcp
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
python scripts\diagnose.py --com      # 真实 COM 连接自检（会启动 Origin）
```

注册到 Agent 客户端（`command`/`args` 用绝对路径，指向 venv 内 python.exe）：

```json
{
  "mcpServers": {
    "origin": {
      "command": "G:\\path\\to\\origin-auto\\origin-mcp\\.venv\\Scripts\\python.exe",
      "args": ["G:\\path\\to\\origin-auto\\origin-mcp\\origin_mcp_server.py"]
    }
  }
}
```

完整工具表、参数、典型工作流、实测坑点见 [`origin-mcp/README.md`](./origin-mcp/README.md)。

---

## 组件二：origin-auto-skill（WorkBuddy 技能）

WorkBuddy 技能。自然语言 "用 Origin 画图" 即自动选择上面的 MCP 工具链；
MCP 未注册时回退到独立脚本 `scripts/quick_plot.py`（同一套 COM 封装，可直接 CLI 跑）。

**安装到 WorkBuddy**

把 `origin-auto-skill/` 整个目录复制到 WorkBuddy 的技能目录：

```powershell
# 用户级（所有项目可用），推荐
Copy-Item -Recurse origin-auto-skill C:\Users\<你>\.workbuddy\skills\origin-auto
```

安装后在 WorkBuddy 里直接说 "用 Origin 画 XX"，技能会自动编排 MCP 调用；
或在没有 Origin/MCP 的环境用回退脚本：

```powershell
python origin-auto-skill\scripts\quick_plot.py --csv data.csv --x 1 --y 2 --out out.png
```

技能内部知识库（颜色/线型/列类型/绘图类型速查、完整坑点）见
[`origin-auto-skill/references/com_pitfalls.md`](./origin-auto-skill/references/com_pitfalls.md)
与 [`origin-auto-skill/SKILL.md`](./origin-auto-skill/SKILL.md)。

---

## 安全约定

- 默认**不强制关闭**已运行的 Origin（`origin_connect` 用单实例类附着到已有实例）
- `quit_origin` 默认拒绝，须显式 `force=true` 且征得用户同意
- `project_new` 会清空当前工程，调用前应先提醒用户保存
- 所有导出/保存走工具内置校验，不轻信 Origin 的布尔返回值

---

## 许可证

[MIT](./LICENSE) © 2026 simcrq
