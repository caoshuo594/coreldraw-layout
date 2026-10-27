# 统一 CorelDRAW MCP

服务名 `coreldraw`，源码与安装方法见 [CorelDRAW MCP 仓库](https://github.com/caoshuo594/coreldraw-mcp)。由 Live 的会话编辑层与 Signage 的模板/数据合并功能整合，底层仍是 Windows COM，非 Corel 官方接口。按仓库说明建立独立 Python 环境，不依赖原电脑安装路径。MCP 启动不启动应用；CorelDRAW 先打开。用当前工具发现结果确认功能，不依据旧工具数量猜参数。

## 选择入口

- 已打开文件：`list_open_documents` → 核实完整路径 → `attach_open_document(name_or_path=完整路径)`。`attach_active_document` 仅用于用户明确指定当前文档。
- 磁盘文件：`open_document(path=完整路径)`；已打开则附加，未打开则打开。模板打开不代表已建立工作副本，修改前另存带时间戳的工作文件。
- 新文件：`create_new_document(width_mm, height_mm)`；单位为文档毫米，实际尺寸还要核对 WorldScale。
- 常规编辑：`inspect_document` 取得 `session_id`、`revision` 和 `shape_id` → `apply_operations`。每批不超过 250 项，短事务、单次撤销；创建操作可用唯一名称供后续操作引用。
- 高级功能：`list_feature_tools(category=...)` 查名称；`list_feature_tools(name=...)` 查该项 JSON 参数结构；再调用 `run_feature_tool(name, arguments, session_id, expected_revision)`。不要臆造入参。
- 常用保存与导出：`save_current_document` 保存已有路径，`save_document_as` 使用绝对新路径，`render_preview` 导出预览，`export_document` 导出生产文件。默认不覆盖；源文件和已验收版保留。
- 磁盘验收：`verify_saved_document(path, preview_path=可选新PNG路径)`，打开磁盘文件的字节一致临时副本，回报原文件 SHA256、页数据、毫米、比例、可编辑文字、字体、字号和溢出，并从该副本导出首页预览。它不保存或关闭用户正在编辑的文件。多页预览及全文/图片匹配仍需按任务补查。

## 操作结构

```json
{
  "session_id": "从附加结果读取",
  "expected_revision": 0,
  "label": "修改图说",
  "operations": [
    {"op": "text.set_content", "shape_id": "从检查结果读取", "value": "完整中文图说"},
    {"op": "text.set_style", "shape_id": "同一对象ID", "font_family": "微软雅黑", "font_size_pt": 18}
  ]
}
```

`expected_revision` 必须使用最近返回值，不固定为 0。常用操作：`text.create_frame`、`text.set_content`、`text.set_style`、`shape.create_rectangle`、`shape.set_bounds`、`shape.move`、`shape.resize`、`shape.rotate`、`shape.delete`、`fill.set_rgb`、`fill.set_cmyk`、`page.append`。创建文本框/矩形使用 `bounds: {x,y,width,height}`，新增页用 `width_mm,height_mm`。完整字段以当前服务代码和返回错误为准。

## 坐标和状态

会话编辑的几何以页面左上角为原点、毫米、Y 向下。高级工具沿用各自函数约定：矩形、椭圆、直线、文本框默认 `top_left=true`；导入、位置、二维码等部分功能仍用原生坐标。先查结构、用已知对象验证换算；不把高级工具一律当左上角坐标。直接 COM 使用原生坐标，并在切页后激活目标、设置并回读单位。

同一进程的工具调用串行，共用一个 STA 队列；统一服务的多个实例通过 Windows 命名互斥体串行执行 COM 任务。该锁不能约束旧 MCP、手动操作或外部 COM 脚本，因此不要同时让其他控制器写同一实例。多步跨工具流程也不构成跨聊天原子事务。

遇到修订冲突，重新检查目标文件、页面与对象后再决定如何继续；不要只把修订号加一。指纹检查只覆盖部分内容，不能发现所有字体、颜色或特殊容器变化。超时任务可能仍在执行；不自动重放，不立刻转用外部 COM 重做，等状态明确后核查对象和输出目录。

## 模板批量

查询 `batch_merge` / `merge_record` 的参数。统一入口使用模板临时副本，按文本对象名称匹配数据列，遍历普通群组与 PowerClip。调用前清点命名占位符；返回 `matched_fields/unmatched_fields` 后核查未匹配项。它不是任意正文全文替换。

CSV 为 UTF-8（可带 BOM）；XLSX/XLS 使用相应读取依赖。批量名称必须是普通文件名，不能产生目录、重复名或已有输出。整个批次先检查目标路径，再执行；后续失败会保留已经完成的文件，并停止，不把半批当完成，也不自动覆盖重跑。

## 验收边界与恢复

`verify_saved_document` 成功表示文件可重新打开并提取结构，不等于版式合格。对照源清单检查全文、图片数量/比例、字体回读、未知溢出状态、单位/比例、裁切、越界和视觉预览。`overflow=null` 表示接口未知，需另查；不能按无溢出通过。

迁移时停用旧 `coreldraw-live` 和 `coreldraw-signage` 配置，再启用统一服务；先保留旧代码、环境和配置备份。新版源码仓库可通过自己的 `requirements.txt` 建立独立 `.venv`。升级时备份本地补丁，重跑握手、会话编辑/冲突、模板批量和保存重开验收。回滚仅恢复相关 MCP 配置段，避免用旧整份配置覆盖后续其他设置。

MCP 不可用时可使用本技能的 COM 路径。先确认统一服务无进行中任务，再串行接管；不要并行控制。
