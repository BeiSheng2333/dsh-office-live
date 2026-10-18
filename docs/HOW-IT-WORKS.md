# 实现逻辑说明 · How dsh-office-live works

这份文档解释插件的分层结构、一次工具调用的完整数据流、以及每个设计选择背后的原因。
想直接看代码可以从 `lib/index.js` 和 `assets/office-live.ps1` 开始。

---

## 1. 它解决什么问题

`dsh-office-tools` 这类插件操作的是**磁盘上的文件**：读 `.docx`、生成新的 `.xlsx`。但用户说
"把这段改一下"时，指的往往是**眼前正开着的那个窗口**——里面可能还有没保存的改动、刚插入的图片、
正在排版的表格。文件工具看不到这些，改完还得让用户自己重开文件才能看到效果。

本插件的目标很具体：

> 让模型能**看见**用户眼前的 Office 窗口，并把改动**实时写回同一个窗口**，用户不用离开文档。

三个直接结果：

- 读取的是活文档（含未保存内容），不是上次保存的副本；
- 修改立即出现在窗口里，用户能 Ctrl+Z、能说"再改改"；
- 模型能通过**截图**看到真实版面（标题居没居中、表格宽不宽），而不只是文本。

---

## 2. 三层结构

```
┌───────────────────────────────────────────────────────────────┐
│ DSH 宿主（Electron，cordis 组合）                              │
│                                                               │
│  ┌─────────────────────────────────────────────────────────┐  │
│  │ dsh-office-live（本插件，ESM）                            │  │
│  │  cordis.patch.yml   → insert 插件节点                     │  │
│  │  lib/index.js       → defineTool() 注册 7 个工具          │  │
│  └───────────────┬─────────────────────────────────────────┘  │
│                  │ child_process.execFile(powershell.exe)     │
│  ┌───────────────▼─────────────────────────────────────────┐  │
│  │ assets/office-live.ps1（Windows PowerShell 5.1+）        │  │
│  │  COM: Word.Application / Excel.Application /            │  │
│  │       PowerPoint.Application                            │  │
│  │  GDI: user32 GetWindowRect + Graphics.CopyFromScreen     │  │
│  └───────────────┬─────────────────────────────────────────┘  │
└──────────────────┼────────────────────────────────────────────┘
                   │ COM / 窗口像素
        ┌──────────▼───────────┐
        │ 用户正在看的 Word /  │
        │ Excel / PowerPoint   │
        └──────────────────────┘
```

- **插件层**负责"在 DSH 里成为一等公民"：报告工具清单、跟随宿主生命周期、出错时只让该工具失败。
- **桥接层**负责所有需要 Windows 原生能力的事：COM 对象调用与屏幕像素抓取。
- 两层的契约极窄：**一个动作名 + 若干字符串参数 → 一个 JSON 对象**。

---

## 3. 一次调用的数据流

以"把第二段改短一点"为例：

1. 模型调用 `office_live_snapshot({ target: "word" })`。
2. `lib/index.js` 生成临时文件路径，`execFile` 启动
   `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File assets/office-live.ps1`
   并追加 `-Action snapshot -Target word -Out <png> -ResultFile <json>`。
3. 脚本用 `Marshal.GetActiveObject("Word.Application")` 附着到**已运行**的 Word 进程
   （凭 ROT 里的运行对象表，不需要知道 PID，也不会新建实例）。
4. 读取：遍历 `ActiveDocument.Paragraphs`（上限 400 段）拿文本/样式/字体，`ComputeStatistics`
   拿页数与字数，`Tables` 拿表格尺寸。
5. 截图：找到 Word 主窗口句柄 → 若最小化则还原 → `SetForegroundWindow` → `GetWindowRect`
   → `Bitmap(w,h)` + `CopyFromScreen` → 存 PNG。
6. 结果通过 `-ResultFile` 以 **UTF-8** 写入临时文件（见 §4.3），进程退出。
7. `lib/index.js` 读文件、`JSON.parse`、把对象作为工具结果返回；`output.render` 把它序列化成
   文本块给模型。临时文件在 `finally` 里删除。
8. 模型据此调用 `office_live_edit({ target: "word", ops: [...] })`：
   同样的桥接方式，但 `-Ops <json 文件>` 传入操作数组，脚本逐条执行并返回每条的结果。
9. 需要"改完的样子"时再 `office_live_shot` 一次，把新 PNG 路径交回模型用 `read_image` 查看。

---

## 4. 关键设计决策

### 4.1 为什么用 COM 附着，而不是读写文件

文件级操作必须"打开→改→关闭"，用户看不到中间状态，且会丢掉未保存改动。COM 的
`GetActiveObject` 给出的是**同一个进程里那个文档对象**，改动即时渲染到窗口。
代价是只能操作"已经打开"的实例——这恰好是我们想要的语义（用户面前那一份）。

### 4.2 为什么工具层 shell 出去调 PowerShell，而不是 node 里直接 COM

Node 没有内置 COM 互操作；要么引入 native 依赖（`win32ole`/`node-ffi` 之类，装个插件要编译、
版本一换就崩），要么走外部进程。PowerShell 自带 COM 与 GDI+ 支持，且用户机器上一定有，
**零原生依赖**是这个选择的核心收益。代价是每次调用多一个进程启动开销（约 0.3–1s），
对这个使用场景（人机来回对话）完全可以接受。

### 4.3 为什么结果走 `-ResultFile` 而不是 stdout

Windows PowerShell 5.1 在被管道重定向时按**控制台代码页**（简体中文系统是 GBK）输出，
插件进程用 UTF-8 解码就会得到乱码；而 `ConvertTo-Json` 不会转义中文。第一次实现就踩了这个坑。
现在脚本把结果用 `System.IO.File::WriteAllText(..., UTF8Encoding(false))` 写成无 BOM 的
UTF-8 文件，Node 端直接 `readFile(..., 'utf8')`，中文彻底不受代码页影响。

### 4.4 为什么默认不保存

`office_live_edit` 只在 `save: true` 时调 `Document.Save()`。默认不保存意味着：

- 用户先在窗口里看效果，不满意可以直接 Ctrl+Z；
- 模型可以连续多次修改，最后由用户决定保存；
- 万一改错，磁盘上的文件还是上一版。

安全优先于便利，这一条是有意的。

### 4.5 为什么只附着已打开实例、不新建隐藏副本

`GetApp(AllowCreate: false)` 在找不到运行实例时返回 `app-not-running` 错误，而不是偷偷
`New-Object -ComObject Word.Application` 起一个不可见实例去改文档——后者会让用户
完全不知道模型动了什么，也会产生幽灵进程。

### 4.6 为什么截图要先 `SetForegroundWindow`

`CopyFromScreen` 抓的是屏幕像素，窗口被别的窗口遮挡时抓到的是遮挡物。所以顺序固定为
还原最小化 → 提到前台 → 等 500ms（等重绘）→ 抓图。副作用是模型截图时会短暂抢焦点，
这是已知且可接受的代价。

### 4.7 为什么用 `-ExecutionPolicy Bypass`

本机（以及不少企业机）`CurrentUser` 策略是 `RemoteSigned`，且某些位置的 `.ps1` 会被判为
未签名而拒绝执行（实测 `D:\...` 下的脚本直接 `& script.ps1` 报"无法在当前系统上运行该脚本"）。
插件自己拉起解释器时显式传 `-ExecutionPolicy Bypass` 只影响这一次子进程，不改系统策略。

### 4.8 错误形状统一为 `{ ok, error: { code, message } }`

脚本层所有失败（没有实例、没打开文档、参数非法、COM 抛异常）都走同一个出口，
JSON 里 `error.code` 是短横线小写词（`app-not-running`、`no-open-document`、`bad-target`…），
`message` 保留原始异常文本。Node 层看到 `error` 字段就抛出 `Error("office-live <code>: <message>")`，
DSH 会把它当作一次正常的工具失败展示给模型，模型可以据此改参数重试。

`edit` 是唯一"部分成功"允许存在的动作：每条 op 独立 try/catch，失败的那条带 `error`，
其余继续执行，最终 `ok` 只表示"是否全部成功"。

---

## 5. 工具契约

| 工具 | 参数 | 成功返回（要点） |
| --- | --- | --- |
| `office_live_status` | — | `apps.{word,excel,powerpoint}.running`，各自打开的文档列表与 active 路径 |
| `office_live_read` | `target?` | Word：`paragraphs[]`（i/text/style/bold/font/size）、`tables[]`、`pageCount`、`wordCount`、`selection`；Excel：`sheets[]`、`usedRange`、`cells[][]`；PPT：`slideCount`、`activeSlide`、`shapes[]` |
| `office_live_shot` | `target?`, `out?` | `path`、`width`、`height`（PNG 绝对路径） |
| `office_live_snapshot` | `target?`, `out?` | `read` + `shot` 两段 |
| `office_live_edit` | `target?`, `ops[]`, `save?` | `results[]`，每条 `{op, ok, ...}` |
| `office_live_open` | `target?`, `file` | `opened`（绝对路径） |
| `office_live_close` | `target?`, `save?` | `closed`、`saved` |

`target` 缺省为 `word`，取值 `word | excel | powerpoint`。

### ops 一览

```
Word
  word.append_paragraph        { text, style?, bold?, size?, font?, align? }
  word.insert_paragraph_before { index, text, style? }
  word.set_paragraph           { index, text?, style?, bold?, size?, font?, align?, color? }
  word.replace                 { find, replace, all? }
  word.insert_table            { data:[[...]], header? }
  word.insert_page_break       {}
Excel
  excel.set_cell               { cell, value?|formula?, bold?, numberFormat?, sheet? }
  excel.set_range              { startCell, data:[[...]], sheet? }
  excel.add_sheet              { name }
PowerPoint
  ppt.set_text                 { slide?, shape?, text, size?, bold? }
  ppt.add_slide                { layout?, title? }
```

`ppt.layout` 用 `PpSlideLayout` 枚举值（`12` = 仅标题之类），默认 `2`（标题+内容）。

---

## 6. 限制与边界（刻意不做的事）

- **只支持桌面版 Microsoft Office**。Microsoft Store 版与网页版没有 COM 自动化接口，WPS 的
  ProgID 也不同（`KWPS.Application` 等）——需要另写一套适配，不在本插件范围内。
- 读取有上限：Word 400 段、Excel 200 行 × 40 列。目的是避免一次工具结果塞爆模型上下文；
  更大的需求应改成分页参数，而不是无限放大。
- 截图是**整窗口像素**（含功能区、标尺、任务窗格），不是文档的纯内容渲染。这正是"用户看到什么
  就抓什么"的语义，但也就不能当"文档缩略图"用。
- 不支持宏、修订痕迹的逐条读写、图表逐元素编辑、SmartArt 内容编辑。
- 没有"监听文档变化"的推送机制：模型是按需 snapshot，而不是常驻监视。

---

## 7. 本地开发与验证

### 7.1 不装进 DSH 也能测桥接层

```powershell
$s = '.\assets\office-live.ps1'
& $s -Action status -ResultFile "$env:TEMP\status.json"   # 需要先 -ExecutionPolicy Bypass
```

### 7.2 不重启 DSH 也能测工具层

插件只依赖 `@deepseek-ai/dsh-tools` 的 `defineTool`。在 `node_modules/@deepseek-ai/dsh-tools/`
放一个桩（导出 `defineTool = (o) => o`），然后用假 ctx 跑 `apply`：

```js
const names = [];
const ctx = {
  tools: { register: (def) => { names.push(def.name); defs.set(def.name, def); return () => {}; } },
  effect: (fn) => fn(),
};
const mod = await import('file:///.../lib/index.js');
mod.apply(ctx, {});
const result = await defs.get('office_live_status').execute({}, {});
```

（真实环境里 `defineTool` 会额外做参数 JSON Schema 编译与运行时校验；桩测试只验证模块能加载、
工具能注册、桥接能跑通。）

### 7.3 打包与本地安装

```powershell
npm pack                                   # 生成 dsh-office-live-0.1.0.tgz
# 把 tgz 放进 profile 目录，然后：
node <pnpm.cjs> install --dir <profile> --config.confirmModulesPurge=false
```

profile 的 `package.json` 需要两处登记：

```json
{ "dependencies": { "dsh-office-live": "file:./dsh-office-live-0.1.0.tgz" },
  "dsh": { "profile": { "bundles": ["...", "dsh-office-tools", "dsh-office-live", "..."] } } }
```

`bundles` 里的顺序决定 cordis 组合顺序；`dsh.bundle.patch` 指向的 `cordis.patch.yml`
在组合时被展开，把插件节点 insert 进来。

### 7.4 激活是否成功

重启后查市场接口即可，不必看界面：

```powershell
(Invoke-WebRequest http://127.0.0.1:<port>/dsh-market/installed).Content
# activation['dsh-office-live'] → { state: 'live', bundle: true, hot: true }
```

---

## 8. 文件职责

| 文件 | 职责 |
| --- | --- |
| `package.json` | npm 元数据 + `dsh.bundle.patch`（bundle 补丁入口）+ `dsh.compatibility`（宿主版本范围） |
| `dsh.plugin.json` | DSH 插件清单：id、入口、`engines.dsh`、`contributes.tools`（7 个工具名） |
| `cordis.patch.yml` | 组合补丁：`insert` 一个 `dsh-office-live` 节点，name 按包名从 node_modules 解析 |
| `lib/index.js` | `defineTool` 注册 7 个工具；参数校验由宿主完成；负责临时文件、子进程、JSON 解析、错误映射 |
| `assets/office-live.ps1` | 全部原生能力：COM 附着与读写、GDI 截图、ops 解释执行、统一 JSON 出口 |
| `README.md` / `README.zh.md` | 用户向说明：能力、依赖、安全约定、安装 |
| `docs/HOW-IT-WORKS.md` | 本文件：实现逻辑与决策依据 |

---

## 9. 后续可做的方向

- 分页读取（`offset/limit`）替代硬上限，支持长文档；
- Excel 图表/条件格式、Word 样式批量套用；
- `office_live_watch`：按间隔出帧 + 文本 diff，做"边改边看"的连续反馈；
- WPS / 仅网页版 Office 的替代适配（不同 ProgID 或 Office.js 加载项路线）。
