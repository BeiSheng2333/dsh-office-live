# dsh-office-live

给 DeepSeek Harness 的「实时 Office」插件：读取、截屏、修改**用户当前正打开着**的
Word / Excel / PowerPoint 文档——不落地转文件、不关闭重开，改动立刻出现在用户眼前的窗口里。

与 `dsh-office-tools`（读写磁盘上的 `.docx`/`.xlsx`/`.pptx` 文件）不同，本插件通过 COM
附着到**正在运行的 Office 进程**，操作的就是那个"活"文档——包括用户保存之后又敲进去的内容。

## 实际效果

Word：`office_live_edit` 追加了两节正文和一张表格，改动已经出现在窗口里（未保存）：

![Word 实时编辑](docs/images/word-live.png)

Excel：`office_live_read` 读到的是公式**求值后**的结果，随后 `excel.set_cell` 往打开的
工作簿里写入"备注"列：

![Excel 实时编辑](docs/images/excel-live.png)

PowerPoint：`ppt.set_text` 重写了项目符号列表，`ppt.add_slide` 增加了第二页：

![PowerPoint 实时编辑](docs/images/powerpoint-live.png)

> 实现逻辑（三层结构、JSON 桥接、为什么结果走 UTF-8 旁路文件、为什么默认不保存）见
> **[docs/HOW-IT-WORKS.md](docs/HOW-IT-WORKS.md)**。

## 工具

| 工具 | 作用 |
| --- | --- |
| `office_live_status` | 查看 Word / Excel / PowerPoint 是否在运行、各自打开了哪些文档 |
| `office_live_read` | 读取实时内容（Word 段落+样式、Excel 单元格+已求值公式、PPT 形状），返回 JSON |
| `office_live_shot` | 截取应用窗口画面为 PNG（含功能区与页面） |
| `office_live_snapshot` | 一次完成 read + shot，改之前先用它 |
| `office_live_edit` | 批量实时修改（段落、查找替换、表格、单元格、幻灯片） |
| `office_live_open` | 用真 Office 打开文件，让修改可见 |
| `office_live_close` | 关闭窗口（可选择保存） |

`office_live_edit` 支持的操作：

- Word —— `word.append_paragraph`、`word.insert_paragraph_before`、`word.set_paragraph`、
  `word.replace`、`word.insert_table`、`word.insert_page_break`
- Excel —— `excel.set_cell`、`excel.set_range`、`excel.add_sheet`
- PowerPoint —— `ppt.set_text`、`ppt.add_slide`

## 依赖

- Windows + 桌面版 Microsoft Office（COM 组件 `Word.Application` / `Excel.Application` /
  `PowerPoint.Application`）。已在 Office LTSC 2024 上验证。
- Windows PowerShell 5.1 及以上（`powershell.exe`）；可用环境变量 `OFFICE_LIVE_POWERSHELL`
  指定别的解释器（如 `pwsh.exe`）。
- 必须是真 Office：Microsoft Store 版/网页版没有 COM 自动化，WPS 的 ProgID 也不同。

## 安全约定

- 只附着**已运行**的实例，不会偷偷新建文档副本。
- 除非显式 `save: true`，修改**不保存**——用户可以先在窗口里看效果、Ctrl+Z 撤销、或继续提要求。
- Word 只读前 400 段；Excel 只读活动工作表 used range 的前 200 行 × 40 列。
- 截图会把窗口提到前台（会短暂抢焦点）。

## 安装

```bash
# 市场界面里添加： dsh-office-live
# 或者从源码/压缩包：
dsh plugin --profile desktop add ./dsh-office-live-0.1.0.tgz
```

重启 DSH 后，新会话即可使用这七个工具。

## 实现方式

`lib/index.js` 用 `@deepseek-ai/dsh-tools` 的 `defineTool` 注册七个工具，调用
`assets/office-live.ps1` 并传 `-ResultFile`（UTF-8 结果文件，避免 Windows 控制台代码页把中文
弄乱）。脚本负责 COM 与截图，每次调用输出一个 JSON 对象。

## 许可

MIT
