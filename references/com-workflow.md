# 本地 COM 操作细节

## 连接与接口

先核实注册入口、安装路径和版本。以下仅为本机 X8 历史有效模式：

```powershell
# Windows PowerShell 5.1 / .NET Framework；优先连接已运行实例
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application.18')
$app.Version
for ($i=1; $i -le $app.Documents.Count; $i++) {
    $d=$app.Documents.Item($i)
    [pscustomobject]@{Index=$i; Name=$d.Name; FullFileName=$d.FullFileName}
}
# 只在确需启动时使用，不保证连接原来的实例
# $app = New-Object -ComObject CorelDRAW.Application.18
# $app.Visible = $true
```

多实例中 ROT 连接可能取不到目标实例；找不到完整路径时先定位实例，不新开同名文件冒充原目标。不调用 Quit 来关闭用户应用，不强制释放共享对象。

脚本保持 UTF-8；PowerShell 5.1 含中文脚本宜用 UTF-8 BOM。优先 64 位 STA 主机。连接失败区分未注册、主机不兼容、无 ROT 实例、权限不同及应用正忙。

优先查 `$app | Get-Member`、`$doc.ActiveLayer | Get-Member` 和本机类型库，不猜签名/枚举。本机历史程序集：

```powershell
$assembly='C:\Program Files\Corel\CorelDRAW Graphics Suite X8\Programs64\Assemblies\Corel.Interop.VGCore.dll'
# 路径存在且适用后：
# Add-Type -Path $assembly
# [Corel.Interop.VGCore.IVGDocument].GetMethods()
# [enum]::GetValues([Corel.Interop.VGCore.cdrFilter])
```

本机信息不足时再查对应版本的 Corel 官方文档。示例不等于已验证的新任务脚本。

## 对象清单

记录页、图层、群组路径、StaticID（若支持）、Type、Name、LeftX、TopY、BottomY、SizeWidth、SizeHeight，文字另记全文、字体、字号及溢出。

历史 X8 类型值：矩形1、位图5、文字6、群组7；其他版本查类型库。COM 集合通常从1开始；区分顶层数量和递归叶对象数量。页面外对象仍可能是有效素材。

普通 Shapes 递归不能保证覆盖 PowerClip、符号、文本链；按接口补查。锁定/隐藏内容不应全部解锁/显示。缓存清单减少 COM 遍历；索引在增删、分组后失效，修改前用 ID、名称、类型、文字片段、几何共同断言身份，集合变化后重建映射。跨文档粘贴不假定 ID 保持不变。

## 几何

```powershell
# 仅用于已确认的新建输出文档；3 是历史 X8 毫米值，优先核实枚举
$doc.Unit=3
$doc.WorldScale=10
$doc.ActivePage.SetSize(700,370)
```

不要为套示例改既有文档原点。先回读原点、页面边界，用已知矩形验证坐标方向和基准。先缩放再定位，避免选择参考点干扰。LeftX/TopY 曾遇 Double/Decimal 转换异常，显式 `[decimal]` 曾有效，但不是通用规则。

### 重开后的单位与跨页定位

2026-10 第七单元实操中，文件曾按毫米排版并存盘，但 `OpenDocument()` 后 `Document.Unit` 回读为 `1`（本机 X8 的英寸）。未经确认就加 `4`，实际位移变成 101.6 mm。`WorldScale=10` 仍不能证明当前单位是毫米；不能据此断言 `Move()` 永远使用英寸。

- 打开文件、激活目标页后，显式设置所需测量单位并回读，再读取边界和计算位移。比例沿用已确认值，不为修复单位错误重设整页比例。
- 找到对象后也要先激活其所属文档和页面。不要以非活动页的 `LeftX/RightX` 算目标位移；错误的大位移曾使图文落入桌面层，随后逐墙计数失配。
- 图片、图框、图说一起移动。短事务中保存修改前坐标，逐项回读实际位移及所属页/图层；异常立即停下，查检查点，不叠加重跑相对位移。

本机 Python COM 的毫米定位示例（目标及对象身份需事先核实；这里的位移是经当前比例换算后的 API 图上毫米）：

```python
doc.Activate()
page.Activate()
doc.Unit = C.cdrMillimeter
assert doc.Unit == C.cdrMillimeter
x, y = float(shape.LeftX), float(shape.TopY)
shape.LeftX = x + dx_api_mm
shape.TopY = y + dy_api_mm
assert abs(float(shape.LeftX) - x - dx_api_mm) < 0.001
assert abs(float(shape.TopY) - y - dy_api_mm) < 0.001
```

以上回读只证明当前事务的坐标，不代替保存后重开检查、目标页配图集合核对及桌面层检查。

## 复制和导入

```powershell
# 已断言所有对象和归属后，串行执行
$sourceDoc.Activate()
$sourceShape.Copy()
$targetDoc.Activate()
$targetPage.Activate()
$pasted=$targetLayer.Paste()
```

核实 Paste 返回类型，不假定永远单个 Shape；比较粘贴前后集合，核对新增数量、尺寸、类型和归属。不依赖 ActiveShape。跨文档 CopyToLayer 曾报“指定对象来自另一个文档”，剪贴板流程成功。

ImportEx 曾需要显式过滤器0和 CreateStructImportOptions()，再调用返回过滤器 Finish()；实际签名先查。空暂存图层有助识别新增对象，非空图层不能假定 Item(1) 就是新图。

本机 X8 实测不支持本地 MCP 旧实现的 `Document.Import(path)`；SVG和位图使用 `page.ActiveLayer.ImportEx(path, 0, app.CreateStructImportOptions())`，随后 `Finish()`。导入前后按 StaticID 差集识别新对象，不用 `Shapes.Last` 猜测新增对象。多个新对象用 `app.CreateShapeRange()` 逐个 `Add` 后组合；只有一个对象时直接使用它，不调用单对象 `Group()` 并假定返回新群组。变换带描边的 SVG 后须检查线宽是否随图形缩放，必要时在SVG中把描边构造成填充路径。

切页或分组异常先用独立文档复现，区分工具适配、脚本对象处理和应用状态问题；不要把一次异常直接写成X8普遍缺陷。跨调用后回读目标页，创建/移动对象后核对实际所属页、图层和StaticID。

本机X8实测：把群组移到页面框外时，可能自动转到共享的“桌面”主图层，该层默认不可打印；`page.Shapes` 也可能包含共享桌面对象。搬移多页内容到同一工作页时，源页范围要按目标页本地图层或明确StaticID筛选，不能直接把每次 `page.Shapes` 全部打包；否则可能把已搬到桌面层的内容再次纳入后续群组。`MoveToLayer(指定本地层)` 可临时恢复归属，但保存/重开后仍可能转回桌面层，必须以磁盘重开结果核实，不能据内存状态保证。框外内容可以保留在桌面层；完整预览使用统一MCP的 `workspace` 范围或临时导出文档，不为输出而把整个桌面层设为可打印，也不擅自改其他桌面素材。跨本地图层和桌面层的混合 `ShapeRange.Copy()` 本机曾报错；临时导出按根对象逐个 Copy/Paste，放到临时页内的本地图层，并在调用导出前再次显式激活临时文档。

## 文字、事务与保存

```powershell
$textShape.Text.Story.Font='微软雅黑' # 先确认本机可用
$textShape.Text.Story.Size=[single]11 # 仅示例
$actualFont=$textShape.Text.Story.Font
$overflow=$textShape.Text.Overflow
```

检查混合文字范围及导出中文、符号和生僻字。保留段落、标点及必要空白；全文核对可统一 CRLF/LF，不能删实质文本后称一致。

使用命令组或关闭重绘/事件时先核实接口，try/finally 恢复原状态。命令组不代替备份。写入结果不明时先检查对象，不能直接重试粘贴。

保存/导出查本机选项结构。SaveAsCopy 曾拒绝 Range=1，不假定支持当前页另存；可以新建独立单页输出。ActiveWindow 为空时跳过视图控制，保存不依赖窗口缩放成功。
