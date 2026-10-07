# Windows Origin 自动化实测要点

以下来自本机 Origin 10.100178 的格式修改验证。换版本或机器时先在一个复制工程上确认，不把本机字体编号写死到通用流程。

## 字体初始化与启动

本机通过 `DispatchEx` 新建的 COM 会话中，`font(Arial)` 曾返回 `0`，导出出现替代字体和文字异常。正常启动的 Origin 会话中返回 `74`，同图正常导出。优先使用正常应用启动加载字体；字体编号由当前会话动态查询。COM 可用于属性读取和工程核验。

本机可把包含 `def AutoExec { ... };` 的 UTF-8 BOM `.cnf` 放在 Origin 用户文件目录，再用 Origin 命令行 `-a <cnf文件名，不含扩展名>` 启动。后台启动在 Windows 使用隐藏窗口。仅管理本次创建的进程，保留用户已有会话。

修改前先枚举文字对象及其原文，包括相对的轴长名称和存在时的 Z 轴，不只处理默认的 `xb`、`yl`。`doc -e G` 遍历当前图层的图形对象，`GD` 包括组内对象；按对象类型识别文字。对所有文本设 Arial 并加粗后，分别设置长名称、刻度和图例字号。

## 字体主题与瀑布图文字

Origin 的字体索引与主题中持久保存的字体 DWORD 编码不同。正常启动、加载字体后，动态查询：

```text
int idx=font(Arial);
double code=font($(idx));
```

`idx` 是当前字体列表索引，`code` 是主题所用 DWORD；这种转换自 Origin 2021 起受支持，见 [Font 官方文档](https://docs.originlab.com/labtalk/ref/font-func/)。不能把本机实测索引或 DWORD 常量写死。

本机瀑布图的样本文字另存字体，普通文字属性修改未覆盖它们。可复制安装自带 `Times New Roman Font.oth` 的全局字体主题结构，生成独立 Arial 主题：把根 `Font` 的 `NodeID 262` 设置为动态取得的 `code`，`263=1` 表示加粗，`264=0` 表示非斜体，再通过 `themeApply2g` 应用至当前图页。先在单个复制图上验证主题效果和导出文字，再推广；保持主题仅含字体属性，避免主题意外覆盖颜色、字号、线宽或数据。

## 已验证的 LabTalk 属性

```text
xb.font=font(Arial);
xb.fsize=36;
yl.font=font(Arial);
yl.fsize=36;
legend.font=font(Arial);
legend.fsize=20;
layer.x.label.font=font(Arial);
layer.x.label.pt=30;
layer.x.label.bold=1;
layer.x.thickness=3;
layer.x.tickthickness=3;
layer.x.mtickthickness=3;
```

Y 轴及实际存在的其他轴使用对应属性。三维图的字号、线宽属性需按其图层类型确认。内嵌字体、常规字重和斜体覆盖可能使整体加粗失效；保留括号内容、数学上下标、希腊字母与数据替换表达式，规范化后重设 Arial 加粗。不要把“移除 \b(...)”当作取消加粗的标准。图形文字对象没有通用的具名 `.bold` 属性，仅含字体属性的加粗主题可处理瀑布图另存样本文字，但具名文字对象中的局部常规格式可能仍覆盖主题。长名称、图例和注释均须覆盖：保留具名文字对象的原内容，逐行以 `\b(该行内容)` 包裹，空行保持为空行。多行图例在本机不能用一个跨换行的 `\b(...)` 包住整段，否则后续行可能恢复常规字重并显示多余右括号。刻度标签另显式设置 `.label.bold=1`。必须检查 EMF 实际字重后确认成功。

曲线的 `set %C -w 1500` 表示 **3 pt**，因为 `-w` 使用 pt × 500。遍历每条目标曲线执行并读回，避免只修改当前曲线。符号边缘宽度使用另一套属性和单位，不把 `-w` 的数值直接用于符号。

图例多行文本可使用 `%(CRLF)`，字符串中的字面 `\n` 在本机曾显示成字母。保留图例的曲线绑定与科学名称。

## 保存与导出

本机 `save -dix "路径"` 曾停在保存对话框；对**不含空格**的完整输出路径，`save -di 完整路径.opju` 已验证能保存。此写法只适用于已确认无空格的路径；其他路径应使用经验证的接口或另选临时输出目录，不能直接照搬去掉引号。

原生导出可用：

```text
expGraph type:=emf export:=page filename:="输出文件名" path:="输出目录" overwrite:=replace;
```

仅对本任务授权的输出文件使用覆盖模式。保持页面与图层的既有几何尺寸，检查重新打开的工程和导出文件。OpenGL 三维图必要时在隔离会话中测试 `@OGLFAE=0; @OGLFA=0;`；其 EMF 可能嵌入位图，不能仅凭扩展名认定为完整矢量图。

## EMF 检查

用标准 EMF 渲染器生成白底预览，核验的是导出文件的实际外观。字体记录可能包含未使用的备用字体；检查字体对象的选择状态与文字绘制记录，判断实际文字是否使用 Arial，以及 LOGFONT `weight=700` 是否为加粗；逐条检查实际文字记录，包括所有刻度和上下标。数学上下标会有小于基准字号的字体记录，这是正常缩放。

## 官方参考

- [图形文字对象属性](https://docs.originlab.com/labtalk/ref/graphic-objs/)
- [坐标刻度标签对象](https://docs.originlab.com/labtalk/ref/layer-axis-label-obj/)
- [轴长名称与图例对象](https://docs.originlab.com/quick-help/set-special-text-bylt/)
- [Set 命令](https://docs.originlab.com/labtalk/ref/set-cmd/)
- [Document 遍历选项](https://docs.originlab.com/labtalk/ref/document_options_others/)
- [Save 命令](https://docs.originlab.com/labtalk/ref/save-cmd/)
