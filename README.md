# coreldraw-layout

面向 Windows 本地 CorelDRAW 的 AI 排版技能：从源文件清点、图文对应与版式执行，到可编辑 CDR 保存、重新打开与视觉验收。

**推荐与 [CorelDRAW MCP](https://github.com/caoshuo594/coreldraw-mcp) 配合使用。** MCP 提供控制 CorelDRAW 的执行工具；本技能指导 AI 何时、怎样使用这些工具，以及如何判断交付是否合格。

## 两个仓库如何配合

| 项目 | 负责什么 | 是否可单独使用 |
| --- | --- | --- |
| **本仓库：coreldraw-layout** | 理解排版要求、图文映射、比例与坐标、异常恢复和保存重开验收 | 可以，通过 Windows 本地 COM 路径执行 |
| **[CorelDRAW MCP](https://github.com/caoshuo594/coreldraw-mcp)** | 连接应用、检查对象、会话编辑、模板批量、保存与导出 | 可以，调用者需自行安排工作流和验收 |

推荐流程：**先安装 MCP → 再安装本技能 → 描述源文件和版式要求 → AI 按技能调用 MCP → 检查保存后的 CDR 与预览。**

配套 MCP 由 [CorelDRAW Live MCP](https://github.com/CubicDev1/corel-mcp) 与 [CorelDRAW Signage MCP](https://github.com/wuguirongsg/coreldraw-mcp) 的本地适配版本整合而来；详细功能继承与改动见其[项目来源记录](https://github.com/caoshuo594/coreldraw-mcp/blob/main/docs/provenance.md)。

## 适用场景

- 展板、展墙与多个墙面的单元排版。
- 图片、标题、正文和图说的对应与组合。
- 按平面图绘制墙面、编号与尺寸标注。
- 继续编辑已有 CDR / CDT，保留原图和可编辑文字。

以 CorelDRAW X8 实操为基础；其他版本先核实接口。尺寸、比例、字号和图片数量由当前任务决定，历史案例数值不是默认值。

默认在单页工作区中分区、分组组织设计，只有任务明确需要时才分页。内容可以有意放在页面框外；页面边界、实际成品区与工作区范围分开处理，不为方便导出而改变源设计。

配套 MCP 的 `workspace` 预览范围会复制目标工作区到临时文档，包含可见桌面对象而不改源桌面层的打印开关；PDF等矢量输出另行确定实际范围。X8 桌面层、保存重开及导出注意事项见 [统一 MCP](references/unified-mcp.md) 和 [COM 操作](references/com-workflow.md)。

## 安装

先按 [MCP 安装说明](https://github.com/caoshuo594/coreldraw-mcp#安装-mcp) 安装执行服务，再把本仓库放到 AI 工具的技能目录。例如在支持该目录的 Codex 环境中：

```powershell
git clone https://github.com/caoshuo594/coreldraw-layout.git "$env:USERPROFILE\.codex\skills\coreldraw-layout"
```

目标已存在时先比较并备份，不覆盖本地修改。也可从 [Releases](https://github.com/caoshuo594/coreldraw-layout/releases) 下载技能包，将完整 `coreldraw-layout` 文件夹放入工具支持的技能目录。

支持 `SKILL.md` 的其他 AI 工具可读取同一目录；不支持自动发现时，直接让它读取 [SKILL.md](SKILL.md)。仅有云端聊天能力还需要本地执行桥接。

单独使用直接 COM 路径时，可跳过 MCP 安装，要求 Windows、匹配位数的 CorelDRAW 和本地执行环境。技能中的 PowerShell 预检脚本使用 Windows PowerShell 5.1、64 位、STA。

## 搭配调用示例

```text
请使用 coreldraw-layout 技能，通过 coreldraw MCP 操作本地 CorelDRAW。
源文件：D:\项目\素材.cdr
参考版式：D:\项目\参考.cdr
输出：D:\项目\交付\展板.cdr
实际尺寸：7000 × 3700 mm，比例请先读取模板确认。
保留指定全文与原图比例，正文和图说保持可编辑。
先核对完整路径和图文对应，再建立工作副本修改。
保存后重新打开验收，交付 CDR、预览与核对记录。
```

支持技能名称调用的工具也可使用 `$coreldraw-layout`。MCP 的精确参数及执行能力见 [工具说明](https://github.com/caoshuo594/coreldraw-mcp/blob/main/docs/tools.md)。

## 工作过程

1. 核实来源、目标文档、单位与比例，建立可恢复工作副本。
2. 清点图文，建立来源对应关系，抽查复杂项。
3. 根据参考排版，保持原图比例、全文范围、图说和阅读顺序。
4. 串行操作，按阶段保存；超时先确认实际状态。
5. 从磁盘重开的文件检查结构、字体、溢出与视觉预览。

MCP 路径按完整路径绑定文档，以会话和修订号修改。常规编辑为页面左上角毫米坐标；高级工具先查询各自参数与坐标约定。

## 文件导航

| 文件 | 适用场景 |
| --- | --- |
| [SKILL.md](SKILL.md) | AI 加载的主入口与核心约束 |
| [统一 MCP](references/unified-mcp.md) | 会话、工具查询、批量、超时与磁盘验收 |
| [COM 操作](references/com-workflow.md) | 版本、连接、原图复制、单位与比例 |
| [版式与验收](references/layout-and-qa.md) | 图片比例、字体、溢出、保存重开 |
| [多墙图文映射](references/full-wall-layout.md) | 全文/图片映射、转角、游线与背景分配 |
| [墙面尺寸标注](references/dimension-annotations.md) | 平面图尺寸解读与标注 |
| [跨工具使用](references/portable-usage.md) | 安装、调用与已有副本同步 |
| [问题处理](references/troubleshooting.md) | 连接、字体、复制、失败与恢复 |
| [只读检查脚本](scripts/get-coreldraw-state.ps1) | 环境、已运行应用、文档与普通群组清单 |
| [界面元数据](agents/openai.yaml) | 技能名称、简介与默认调用提示 |

## 只读预检

在仓库根目录执行：

```powershell
# 检查环境，不连接或启动 CorelDRAW
powershell.exe -NoProfile -STA -File .\scripts\get-coreldraw-state.ps1

# 检查已运行的 CorelDRAW 和文档列表
powershell.exe -NoProfile -STA -File .\scripts\get-coreldraw-state.ps1 -InspectRunning
```

脚本不是一键排版器，不创建、保存或关闭文档。普通群组遍历完成不能证明 PowerClip、符号、文本链和视觉版式通过验收。

## 验证边界

技能结构与相对引用已经校验，配套 MCP 提供自动测试与 X8 真实链路验证。历史案例说明已经观察到的行为，不代表新项目已经验收。

交付仍需对照本次来源清单，检查完整文字、配图、裁切、比例、字体、溢出和视觉效果。只有 PNG、文件存在或脚本没有报错，均不足以证明可编辑 CDR 已完成。

## 更新与反馈

更新前备份本地改动，保留整个目录的相对结构。同步多个工具时只更新已存在且确认的技能副本，保留 Junction / 符号链接。文件同步后按工具要求重新加载技能。

问题可通过 [Issues](https://github.com/caoshuo594/coreldraw-layout/issues) 提交，请附 CorelDRAW 版本、任务条件、复现步骤、报错及最后可恢复状态。
