# =====================================================================
# Office Live Bridge — 让 DSH 智能体读写"用户正打开着的" Office 文档
#
# 能力：
#   status    列出 Word / Excel / PowerPoint 是否在运行、打开了哪些文件
#   open      用真 Office 打开文件（弹出窗口）
#   read      读取当前活动文档的内容结构（段落/单元格/形状，JSON）
#   shot      截取该应用窗口画面为 PNG（智能体可"看版面"）
#   snapshot  read + shot 一次完成
#   edit      批量实时修改（Word / Excel / PowerPoint），返回每个操作结果
#   close     关闭当前文档（默认不保存）
#
# 设计约定：
#   * 只附着到"已经打开"的 Office 实例，绝不新建/覆盖用户文档，除非调用 open / close
#   * 默认不保存：改动即时显示在窗口里，保存需显式 -Save
#   * 通过 COM 调用本机 Microsoft Office；窗口截图用 user32 + System.Drawing
#
# 输出：默认向 stdout 打印一行 JSON；传 -ResultFile <path> 时改为写入 UTF-8 文件
#       （插件场景用后者，避免中文编码问题）
# =====================================================================
param(
  [Parameter(Mandatory = $true)][string]$Action,
  [string]$Target = 'word',
  [string]$File,
  [string]$Out,
  [string]$Ops,
  [string]$ResultFile,
  [switch]$Save
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}
try { $OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}

Add-Type -AssemblyName System.Drawing

if (-not ('OfficeLiveNative' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public struct OLRECT { public int Left; public int Top; public int Right; public int Bottom; }
public static class OfficeLiveNative {
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out OLRECT r);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
}
'@
}

function Out-Json($obj) {
  $json = $obj | ConvertTo-Json -Depth 8 -Compress
  if ($ResultFile) {
    [System.IO.File]::WriteAllText($ResultFile, $json, (New-Object System.Text.UTF8Encoding($false)))
  } else {
    Write-Output $json
  }
}

function Fail($code, $msg) {
  Out-Json @{ ok = $false; error = @{ code = $code; message = $msg } }
  exit 0
}

$progIds = @{ word = 'Word.Application'; excel = 'Excel.Application'; powerpoint = 'PowerPoint.Application' }
$processNames = @{ word = 'WINWORD'; excel = 'EXCEL'; powerpoint = 'POWERPNT' }
if (-not $progIds.ContainsKey($Target)) { Fail 'bad-target' "unknown -Target '$Target' (word|excel|powerpoint)" }

# 附着到已运行的实例；只有 -Action open 允许新建实例
function Get-App([bool]$AllowCreate) {
  $progId = $progIds[$Target]
  try { $existing = [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) } catch { $existing = $null }
  if ($existing) { return $existing }
  if (-not $AllowCreate) { return $null }
  return New-Object -ComObject $progId
}

function Get-MainWindow {
  return Get-Process -Name $processNames[$Target] -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
}

function Save-WindowShot([string]$path) {
  $p = Get-MainWindow
  if (-not $p) { return @{ ok = $false; error = 'no-window' } }
  $h = $p.MainWindowHandle
  if ([OfficeLiveNative]::IsIconic($h)) { [OfficeLiveNative]::ShowWindow($h, 9) | Out-Null }
  [OfficeLiveNative]::SetForegroundWindow($h) | Out-Null
  Start-Sleep -Milliseconds 500
  $r = [OLRECT]::new()
  [OfficeLiveNative]::GetWindowRect($h, [ref]$r) | Out-Null
  $w = $r.Right - $r.Left
  $hh = $r.Bottom - $r.Top
  if ($w -le 0 -or $hh -le 0) { return @{ ok = $false; error = 'bad-rect' } }
  $dir = Split-Path -Parent $path
  if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  $bmp = [System.Drawing.Bitmap]::new($w, $hh)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($r.Left, $r.Top, 0, 0, [System.Drawing.Size]::new($w, $hh))
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose()
  $bmp.Dispose()
  return @{ ok = $true; path = $path; width = $w; height = $hh }
}

function Read-Word($app) {
  if ($app.Documents.Count -lt 1) { return @{ ok = $false; error = 'no-open-document' } }
  $doc = $app.ActiveDocument
  $paras = @()
  $i = 0
  foreach ($p in $doc.Paragraphs) {
    $i++
    if ($i -gt 400) { break }
    $rg = $p.Range
    $t = $rg.Text
    if ($t) { $t = $t.TrimEnd("`r", "`n") }
    $paras += [ordered]@{
      i     = $i
      text  = $t
      style = $p.Style.NameLocal
      bold  = [bool]$rg.Font.Bold
      font  = $rg.Font.Name
      size  = $rg.Font.Size
    }
  }
  $tables = @()
  foreach ($tb in $doc.Tables) {
    $tables += [ordered]@{ rows = $tb.Rows.Count; cols = $tb.Columns.Count; rangeStart = $tb.Range.Start; rangeEnd = $tb.Range.End }
  }
  $sel = $null
  try { $sel = @{ start = $app.Selection.Start; end = $app.Selection.End; text = $app.Selection.Text } } catch {}
  return [ordered]@{
    ok             = $true
    target         = 'word'
    file           = $doc.FullName
    saved          = ($doc.Saved -eq -1)
    pageCount      = $doc.ComputeStatistics(2)
    wordCount      = $doc.ComputeStatistics(0)
    paragraphCount = $doc.Paragraphs.Count
    tables         = $tables
    selection      = $sel
    paragraphs     = $paras
  }
}

function Read-Excel($app) {
  if ($app.Workbooks.Count -lt 1) { return @{ ok = $false; error = 'no-open-workbook' } }
  $wb = $app.ActiveWorkbook
  $ws = $app.ActiveSheet
  $used = $ws.UsedRange
  $rows = $used.Rows.Count
  $cols = $used.Columns.Count
  $vals = $used.Value2
  $grid = @()
  if ($null -ne $vals) {
    if ($rows -eq 1 -and $cols -eq 1) { $grid = @(, @("$vals")) }
    else {
      for ($r = 1; $r -le [Math]::Min($rows, 200); $r++) {
        $line = @()
        for ($c = 1; $c -le [Math]::Min($cols, 40); $c++) {
          $v = $vals.GetValue($r, $c)
          $line += ($(if ($null -eq $v) { '' } else { "$v" }))
        }
        $grid += , $line
      }
    }
  }
  $sheets = @()
  foreach ($s in $wb.Worksheets) { $sheets += $s.Name }
  return [ordered]@{
    ok         = $true
    target     = 'excel'
    file       = $wb.FullName
    saved      = [bool]$wb.Saved
    sheets     = $sheets
    activeSheet = $ws.Name
    usedRange  = $used.Address($false, $false)
    rows       = $rows
    cols       = $cols
    cells      = $grid
  }
}

function Read-PowerPoint($app) {
  if ($app.Presentations.Count -lt 1) { return @{ ok = $false; error = 'no-open-presentation' } }
  $pres = $app.ActivePresentation
  $slide = $app.ActiveWindow.View.Slide
  $shapes = @()
  foreach ($sh in $slide.Shapes) {
    $txt = $null
    try { if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) { $txt = $sh.TextFrame.TextRange.Text } } catch {}
    $shapes += [ordered]@{ name = $sh.Name; type = $sh.Type; left = $sh.Left; top = $sh.Top; width = $sh.Width; height = $sh.Height; text = $txt }
  }
  return [ordered]@{
    ok          = $true
    target      = 'powerpoint'
    file        = $pres.FullName
    slideCount  = $pres.Slides.Count
    activeSlide = $slide.SlideIndex
    shapes      = $shapes
  }
}

function Invoke-Edit($app, $opList) {
  $results = @()
  foreach ($op in $opList) {
    $kind = [string]$op.op
    try {
      switch ($kind) {

        # ---------------- Word ----------------
        'word.append_paragraph' {
          $doc = $app.ActiveDocument
          $doc.Content.InsertParagraphAfter()
          $rg = $doc.Paragraphs.Last.Range
          $rg.Text = [string]$op.text
          try { if ($op.style) { $rg.Style = $op.style } } catch {}
          try { if ($null -ne $op.bold) { $rg.Font.Bold = [int][bool]$op.bold } } catch {}
          try { if ($null -ne $op.size) { $rg.Font.Size = [double]$op.size } } catch {}
          try { if ($op.font) { $rg.Font.Name = [string]$op.font } } catch {}
          try { if ($null -ne $op.align) { $rg.ParagraphFormat.Alignment = [int]$op.align } } catch {}
          $results += [ordered]@{ op = $kind; ok = $true; at = $doc.Paragraphs.Count }
        }
        'word.insert_paragraph_before' {
          $doc = $app.ActiveDocument
          $rg = $doc.Paragraphs.Item([int]$op.index).Range
          $rg.InsertParagraphBefore()
          $rg2 = $doc.Paragraphs.Item([int]$op.index).Range
          $rg2.Text = [string]$op.text
          try { if ($op.style) { $rg2.Style = $op.style } } catch {}
          $results += [ordered]@{ op = $kind; ok = $true }
        }
        'word.set_paragraph' {
          $doc = $app.ActiveDocument
          $rg = $doc.Paragraphs.Item([int]$op.index).Range
          if ($null -ne $op.text) { $rg.Text = [string]$op.text }
          try { if ($op.style) { $rg.Style = $op.style } } catch {}
          try { if ($null -ne $op.bold) { $rg.Font.Bold = [int][bool]$op.bold } } catch {}
          try { if ($null -ne $op.size) { $rg.Font.Size = [double]$op.size } } catch {}
          try { if ($op.font) { $rg.Font.Name = [string]$op.font } } catch {}
          try { if ($null -ne $op.align) { $rg.ParagraphFormat.Alignment = [int]$op.align } } catch {}
          try { if ($null -ne $op.color) { $rg.Font.Color = [int]$op.color } } catch {}
          $results += [ordered]@{ op = $kind; ok = $true }
        }
        'word.replace' {
          $doc = $app.ActiveDocument
          $find = $doc.Content.Find
          $find.ClearFormatting()
          $find.Replacement.ClearFormatting()
          $find.Text = [string]$op.find
          $find.Replacement.Text = [string]$op.replace
          $replaceMode = if ($op.all -ne $false) { 2 } else { 1 }
          $found = $find.Execute([string]$op.find, $false, $false, $false, $false, $false, $true, 1, $false, [string]$op.replace, $replaceMode)
          $results += [ordered]@{ op = $kind; ok = [bool]$found; found = [bool]$found }
        }
        'word.insert_table' {
          $doc = $app.ActiveDocument
          $doc.Content.InsertParagraphAfter()
          $rg = $doc.Paragraphs.Last.Range
          $data = @($op.data)
          $rows = [int]$op.rows
          $cols = [int]$op.cols
          if ($data.Count -gt 0) { $rows = $data.Count; $cols = @($data[0]).Count }
          $tb = $doc.Tables.Add($rg, $rows, $cols)
          if ($op.style) { try { $tb.Style = $op.style } catch {} }
          if ($op.header -eq $true) { try { $tb.Rows.Item(1).Range.Font.Bold = -1 } catch {} }
          $r = 0
          foreach ($row in $data) {
            $r++
            $c = 0
            foreach ($cell in @($row)) {
              $c++
              try { $tb.Cell($r, $c).Range.Text = [string]$cell } catch {}
            }
          }
          $results += [ordered]@{ op = $kind; ok = $true; rows = $rows; cols = $cols }
        }
        'word.insert_page_break' {
          $app.Selection.InsertBreak(7)
          $results += [ordered]@{ op = $kind; ok = $true }
        }

        # ---------------- Excel ----------------
        'excel.set_cell' {
          $ws = if ($op.sheet) { $app.Workbooks.Item(1).Worksheets.Item([string]$op.sheet) } else { $app.ActiveSheet }
          $rg = $ws.Range([string]$op.cell)
          if ([string]$op.formula) { $rg.Formula = [string]$op.formula } else { $rg.Value2 = $op.value }
          try { if ($null -ne $op.bold) { $rg.Font.Bold = [int][bool]$op.bold } } catch {}
          try { if ($null -ne $op.numberFormat) { $rg.NumberFormat = [string]$op.numberFormat } } catch {}
          $results += [ordered]@{ op = $kind; ok = $true; cell = [string]$op.cell }
        }
        'excel.set_range' {
          $ws = if ($op.sheet) { $app.Workbooks.Item(1).Worksheets.Item([string]$op.sheet) } else { $app.ActiveSheet }
          $start = [string]$op.startCell
          $data = @($op.data)
          $r0 = $ws.Range($start).Row
          $c0 = $ws.Range($start).Column
          $r = 0
          foreach ($row in $data) {
            $c = 0
            foreach ($cell in @($row)) {
              $ws.Cells.Item($r0 + $r, $c0 + $c).Value2 = $cell
              $c++
            }
            $r++
          }
          $results += [ordered]@{ op = $kind; ok = $true; rows = $data.Count }
        }
        'excel.add_sheet' {
          $wb = $app.ActiveWorkbook
          $ws = $wb.Worksheets.Add()
          $ws.Name = [string]$op.name
          $results += [ordered]@{ op = $kind; ok = $true; name = $ws.Name }
        }

        # ---------------- PowerPoint ----------------
        'ppt.set_text' {
          $pres = $app.ActivePresentation
          $slide = if ($op.slide) { $pres.Slides.Item([int]$op.slide) } else { $app.ActiveWindow.View.Slide }
          $sh = if ($op.shape) { $slide.Shapes.Item([string]$op.shape) } else { $slide.Shapes.Item(1) }
          if ($sh.HasTextFrame -eq -1) {
            $sh.TextFrame.TextRange.Text = [string]$op.text
          } else { throw "shape '$($sh.Name)' has no text frame" }
          try { if ($null -ne $op.size) { $sh.TextFrame.TextRange.Font.Size = [double]$op.size } } catch {}
          try { if ($null -ne $op.bold) { $sh.TextFrame.TextRange.Font.Bold = [int][bool]$op.bold } } catch {}
          $results += [ordered]@{ op = $kind; ok = $true; shape = $sh.Name }
        }
        'ppt.add_slide' {
          $pres = $app.ActivePresentation
          $layout = if ($null -ne $op.layout) { [int]$op.layout } else { 2 }
          $idx = $pres.Slides.Add($pres.Slides.Count + 1, $layout)
          try { if ($op.title) { $idx.Shapes.Item(1).TextFrame.TextRange.Text = [string]$op.title } } catch {}
          $results += [ordered]@{ op = $kind; ok = $true; slide = $idx.SlideIndex }
        }

        default { throw "unknown op '$kind'" }
      }
    } catch {
      $results += [ordered]@{ op = $kind; ok = $false; error = $_.Exception.Message }
    }
  }
  return $results
}

# ---------------------------------------------------------------------
try {
  switch ($Action) {

    'status' {
      $state = @{}
      foreach ($t in 'word', 'excel', 'powerpoint') {
        $Target = $t
        $app = Get-App $false
        if ($app) {
          $info = @{ running = $true }
          try {
            switch ($t) {
              'word' {
                $info.documents = @()
                foreach ($d in $app.Documents) { $info.documents += @{ name = $d.Name; path = $d.FullName; saved = ($d.Saved -eq -1) } }
                try { $info.active = $app.ActiveDocument.FullName } catch {}
              }
              'excel' {
                $info.workbooks = @()
                foreach ($w in $app.Workbooks) { $info.workbooks += @{ name = $w.Name; path = $w.FullName; saved = [bool]$w.Saved } }
                try { $info.active = $app.ActiveWorkbook.FullName } catch {}
              }
              'powerpoint' {
                $info.presentations = @()
                foreach ($pz in $app.Presentations) { $info.presentations += @{ name = $pz.Name; path = $pz.FullName; saved = [bool]$pz.Saved } }
                try { $info.active = $app.ActivePresentation.FullName } catch {}
              }
            }
          } catch { $info.error = $_.Exception.Message }
          $state[$t] = $info
        } else { $state[$t] = @{ running = $false } }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
      }
      Out-Json @{ ok = $true; apps = $state }
    }

    'open' {
      if (-not $File) { Fail 'missing-file' '-Action open needs -File' }
      if (-not (Test-Path $File)) { Fail 'file-not-found' $File }
      $File = (Resolve-Path $File).Path
      $app = Get-App $true
      switch ($Target) {
        'word' { $app.Visible = $true; $doc = $app.Documents.Open($File) }
        'excel' { $app.Visible = $true; $wb = $app.Workbooks.Open($File) }
        'powerpoint' { $app.Visible = -1; $pres = $app.Presentations.Open($File, 0, 0, -1) }
      }
      Start-Sleep -Milliseconds 800
      Out-Json @{ ok = $true; target = $Target; opened = $File }
    }

    'read' {
      $app = Get-App $false
      if (-not $app) { Fail 'app-not-running' "no running $Target instance" }
      $r = switch ($Target) {
        'word' { Read-Word $app }
        'excel' { Read-Excel $app }
        'powerpoint' { Read-PowerPoint $app }
      }
      Out-Json $r
    }

    'shot' {
      if (-not $Out) { Fail 'missing-out' '-Action shot needs -Out <png path>' }
      $r = Save-WindowShot $Out
      Out-Json ($r + @{ target = $Target })
    }

    'snapshot' {
      if (-not $Out) { Fail 'missing-out' '-Action snapshot needs -Out <png path>' }
      $app = Get-App $false
      if (-not $app) { Fail 'app-not-running' "no running $Target instance" }
      $read = switch ($Target) {
        'word' { Read-Word $app }
        'excel' { Read-Excel $app }
        'powerpoint' { Read-PowerPoint $app }
      }
      $shot = Save-WindowShot $Out
      Out-Json @{ ok = $true; target = $Target; read = $read; shot = $shot }
    }

    'edit' {
      if (-not $Ops) { Fail 'missing-ops' '-Action edit needs -Ops <json file>' }
      if (-not (Test-Path $Ops)) { Fail 'ops-not-found' $Ops }
      $opList = Get-Content $Ops -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($opList -isnot [array]) { $opList = @($opList) }
      $app = Get-App $false
      if (-not $app) { Fail 'app-not-running' "no running $Target instance" }
      $results = Invoke-Edit $app $opList
      $anyFail = @($results | Where-Object { -not $_.ok }).Count -gt 0
      if ($Save -and -not $anyFail) {
        switch ($Target) {
          'word' { $app.ActiveDocument.Save() }
          'excel' { $app.ActiveWorkbook.Save() }
          'powerpoint' { $app.ActivePresentation.Save() }
        }
      }
      Out-Json @{ ok = (-not $anyFail); target = $Target; saved = [bool]$Save; results = $results }
    }

    'close' {
      $app = Get-App $false
      if (-not $app) { Fail 'app-not-running' "no running $Target instance" }
      switch ($Target) {
        'word' { if ($Save) { $app.ActiveDocument.Save() }; $app.ActiveDocument.Close(0) }
        'excel' { if ($Save) { $app.ActiveWorkbook.Save() }; $app.ActiveWorkbook.Close(0) }
        'powerpoint' { if ($Save) { $app.ActivePresentation.Save() }; $app.ActivePresentation.Close() }
      }
      Out-Json @{ ok = $true; target = $Target; closed = $true; saved = [bool]$Save }
    }

    default { Fail 'bad-action' "unknown -Action '$Action'" }
  }
} catch {
  Fail 'exception' ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
