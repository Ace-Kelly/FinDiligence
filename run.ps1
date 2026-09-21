<#
.SYNOPSIS
  尽职调查网络核查截图助手 —— 驱动你自己的浏览器，逐站逐公司截图归档。

.DESCRIPTION
  以「网站优先」顺序循环：打开一个网站 → 把所有公司搜一遍并截图 → 换下一个网站。
  脚本只负责三件事：复制搜索词到剪贴板、在你确认后全屏截图、按规则命名归档并生成
  Markdown / Word。搜索、翻页、验证码由你在浏览器里手动完成。

  为什么不做浏览器自动化，见 README.md。

.EXAMPLE
  .\run.ps1
  .\run.ps1 -Site 04
  .\run.ps1 -Company 示例集团 -FromSite 02 -ToSite 05
  .\run.ps1 -Sites config\sites.gov33.json -Companies config\companies.example.txt
#>

param(
  [string]$Sites = "",
  [string]$Companies = "",
  [string]$Out = "",
  [string]$Company = "",
  [string]$Site = "",
  [string]$FromSite = "",
  [string]$ToSite = "",
  [string]$Group = "",
  [switch]$Overwrite,
  [switch]$NewWindow,
  [string]$Browser = "msedge",
  [string]$Python = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ---------------- paths & defaults ----------------
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path

if (!$Sites)     { $Sites     = Join-Path $Here "config\sites.example.json" }
if (!$Companies) { $Companies = Join-Path $Here "config\companies.example.txt" }
if (!$Out)       { $Out       = Join-Path $Here ("截图输出_" + (Get-Date -Format "yyyyMMdd")) }

if (!(Test-Path -LiteralPath $Sites))     { throw "找不到网站配置: $Sites" }
if (!(Test-Path -LiteralPath $Companies)) { throw "找不到公司清单: $Companies" }

if (!$Python) {
  $Python = if (Get-Command python -ErrorAction SilentlyContinue) { "python" }
            elseif (Get-Command py -ErrorAction SilentlyContinue) { "py" }
            else { throw "找不到 Python。请安装 Python 3 并确保 python 命令可用，或用 -Python 指定路径。" }
}

# ---------------- load config ----------------
$siteList = Get-Content -LiteralPath $Sites -Raw -Encoding UTF8 | ConvertFrom-Json

<#
  公司清单支持两种格式：

  1) .txt —— 一行一个公司名。每个网站给每家公司截 1 张图，最省事。
  2) .json —— 结构化清单，一家公司在一个网站下可以有多个核查对象
     （发行人、子公司、董监高……），文件夹名和图片名完全自定义：

     {
       "示例集团": {
         "name": "示例集团有限公司",
         "group": "发行人",
         "sites": {
           "01": {
             "folderName": "工商信息",
             "paragraphs": [
               { "text": "发行人：示例集团有限公司",
                 "imgName": "发行人：示例集团有限公司.jpg",
                 "searchName": "示例集团有限公司" }
             ]
           }
         }
       }
     }
#>
function ConvertFrom-CompanyList {
  param([string]$Path, [array]$SiteList)

  if ([System.IO.Path]::GetExtension($Path).ToLower() -eq ".json") {
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $result = @()
    foreach ($prop in $raw.PSObject.Properties) {
      $c = $prop.Value
      $result += [PSCustomObject]@{
        key   = $prop.Name
        name  = $c.name
        group = if ($c.group) { $c.group } else { "" }
        sites = $c.sites
      }
    }
    return $result
  }

  # 纯文本清单：自动展开成「每站 1 张图」的结构
  $names = Get-Content -LiteralPath $Path -Encoding UTF8 |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith("#") }

  $result = @()
  foreach ($n in $names) {
    $sites = [ordered]@{}
    foreach ($s in $SiteList) {
      $sites[$s.id] = [PSCustomObject]@{
        folderName = "$($s.id)-$($s.name)"
        paragraphs = @([PSCustomObject]@{
          text       = $n
          imgName    = "$n.jpg"
          searchName = $n
        })
      }
    }
    $result += [PSCustomObject]@{
      key = $n; name = $n; group = ""
      sites = [PSCustomObject]$sites
    }
  }
  return $result
}

$allCompanies = ConvertFrom-CompanyList -Path $Companies -SiteList $siteList

# ---------------- filters ----------------
if ($Site)     { $siteList = $siteList | Where-Object { $_.id -eq $Site -or $_.name -match [regex]::Escape($Site) } }
if ($FromSite) { $siteList = $siteList | Where-Object { $_.id -ge $FromSite } }
if ($ToSite)   { $siteList = $siteList | Where-Object { $_.id -le $ToSite } }
if (!$siteList -or @($siteList).Count -eq 0) { throw "没有匹配的网站。" }

$targetCompanies = $allCompanies |
  Where-Object { [string]::IsNullOrEmpty($Group)   -or $_.group -eq $Group } |
  Where-Object { [string]::IsNullOrEmpty($Company) -or $_.name  -match [regex]::Escape($Company) -or $_.key -match [regex]::Escape($Company) }
if (@($targetCompanies).Count -eq 0) { throw "没有匹配的公司。" }
$targetCompanies = @($targetCompanies | Sort-Object group, name)

# ---------------- helpers ----------------
function Get-SafeName([string]$Name) {
  return (($Name -replace '[\\/:*?"<>|]', '_') -replace '\s+', ' ').Trim()
}

# 去掉「发行人：」「1、」这类前缀，只留下真正要搜的名字
function Get-SearchTerm([string]$Name, [array]$Prefixes) {
  if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
  $pure = $Name.Trim() -replace '^\d+[.、]\s*', ''
  foreach ($p in $Prefixes) {
    $pure = $pure -replace ("^" + [regex]::Escape($p) + '[：:]\s*'), ''
  }
  return $pure.Trim()
}

function Get-BrowserPath([string]$Name) {
  $exe = if ($Name -match 'chrome') { "chrome.exe" } else { "msedge.exe" }
  $dir = if ($Name -match 'chrome') { "Google\Chrome" } else { "Microsoft\Edge" }
  $candidates = @(
    "$env:ProgramFiles\$dir\Application\$exe",
    "${env:ProgramFiles(x86)}\$dir\Application\$exe",
    "$env:LOCALAPPDATA\$dir\Application\$exe"
  )
  foreach ($item in $candidates) { if ($item -and (Test-Path $item)) { return $item } }
  return $exe
}

if (-not ('DdCaptureWin' -as [type])) {
  Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DdCaptureWin {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
  [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  public const int SW_MINIMIZE = 6;
  public const int SW_RESTORE  = 9;
  public const int SW_SHOW     = 3;
}
"@
}

# 返回浏览器主窗口句柄，截图脚本靠它判断该截哪一块显示器
function Enable-BrowserWindow {
  $procName = if ($Browser -match 'chrome') { "chrome" } else { "msedge" }
  $proc = Get-Process $procName -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } |
    Sort-Object StartTime -Descending | Select-Object -First 1
  if ($proc) {
    (New-Object -ComObject WScript.Shell).AppActivate($proc.Id) | Out-Null
    Start-Sleep -Milliseconds 300
    [DdCaptureWin]::ShowWindowAsync($proc.MainWindowHandle, [DdCaptureWin]::SW_SHOW) | Out-Null
    Start-Sleep -Milliseconds 300
    [DdCaptureWin]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 600
    return $proc.MainWindowHandle
  }
  return [IntPtr]::Zero
}

function Open-Url([string]$Url) {
  $exe = Get-BrowserPath $Browser
  $argList = if ($NewWindow) { @("--new-window", $Url) } else { @($Url) }
  Start-Process -FilePath $exe -ArgumentList $argList
  Start-Sleep -Seconds 2
  Enable-BrowserWindow | Out-Null
}

function Save-Screenshot([string]$OutFile) {
  $console = [DdCaptureWin]::GetConsoleWindow()
  try {
    # 截图前最小化本终端窗口，免得终端被拍进图里
    if ($console -ne [IntPtr]::Zero) {
      [DdCaptureWin]::ShowWindowAsync($console, [DdCaptureWin]::SW_MINIMIZE) | Out-Null
      Start-Sleep -Milliseconds 400
    }
    $hwnd = Enable-BrowserWindow

    $capArgs = @((Join-Path $Here "capture_screen.py"), $OutFile)
    if ($hwnd -and $hwnd -ne [IntPtr]::Zero) { $capArgs += @("--hwnd", [string]([int64]$hwnd)) }
    & $Python $capArgs
    if ($LASTEXITCODE -ne 0)                    { throw "截图失败，退出码 $LASTEXITCODE" }
    if (!(Test-Path -LiteralPath $OutFile))     { throw "没有生成截图文件: $OutFile" }
    if ((Get-Item -LiteralPath $OutFile).Length -le 0) { throw "截图文件为空: $OutFile" }
  }
  finally {
    if ($console -ne [IntPtr]::Zero) {
      [DdCaptureWin]::ShowWindowAsync($console, [DdCaptureWin]::SW_RESTORE) | Out-Null
      Start-Sleep -Milliseconds 200
    }
  }
}

function New-Markdown([string]$Path, [array]$Items) {
  $sb = [System.Text.StringBuilder]::new()
  $num = 1
  foreach ($p in $Items) {
    [void]$sb.AppendLine("$num. $($p.text)")
    [void]$sb.AppendLine("")
    if (Test-Path -LiteralPath $p.imgPath) {
      [void]$sb.AppendLine("![]($($p.imgName))")
    } else {
      Write-Warning "缺少截图，Markdown 中只保留序号和名称: $($p.imgPath)"
    }
    [void]$sb.AppendLine("")
    $num++
  }
  New-Item -ItemType Directory -Force -Path (Split-Path $Path) | Out-Null
  Set-Content -LiteralPath $Path -Value $sb.ToString() -Encoding UTF8
}

# ---------------- header ----------------
Write-Host "===== 尽调网络核查截图助手 ====="
Write-Host "公司数: $(@($targetCompanies).Count)   网站数: $(@($siteList).Count)"
Write-Host "输出目录: $Out"
Write-Host ""
Write-Host "逻辑：打开一个网站 → 依次把所有公司搜一遍并截图 → 换下一个网站。"
Write-Host ""
Write-Host "操作键：[回车]=截图并下一个  r=重截  o=重开网址  s=跳过  n=换下家  q=退出"
Write-Host ""

# ---------------- main loop ----------------
$done = 0
$todo = 0

:siteLoop foreach ($s in $siteList) {
  $workItems = @()
  foreach ($c in $targetCompanies) {
    $sd = $c.sites.$($s.id)
    if (!$sd -or !$sd.paragraphs -or @($sd.paragraphs).Count -eq 0) { continue }

    $cDir = if ($c.group) { Join-Path (Join-Path $Out (Get-SafeName $c.group)) (Get-SafeName $c.name) }
            else          { Join-Path $Out (Get-SafeName $c.name) }
    $sDir = Join-Path $cDir (Get-SafeName $sd.folderName)

    $images = @()
    foreach ($p in $sd.paragraphs) {
      $images += [PSCustomObject]@{
        text = $p.text; imgName = $p.imgName; searchName = $p.searchName
        imgPath = Join-Path $sDir $p.imgName
      }
    }
    $workItems += [PSCustomObject]@{
      company = $c.name; folderName = (Get-SafeName $sd.folderName)
      siteDir = $sDir; images = $images
    }
  }

  if ($workItems.Count -eq 0) { Write-Host "[$($s.id)/$($s.name)] 无待查项，跳过。"; continue }

  Write-Host ""
  Write-Host "########## [$($s.id)] $($s.name) ##########"
  Write-Host "网址: $($s.url)"
  if ($s.note) { Write-Host "提示: $($s.note)" }
  $totalImgs = ($workItems | ForEach-Object { $_.images.Count } | Measure-Object -Sum).Sum
  Write-Host "待查: $($workItems.Count) 家公司, 共 $totalImgs 张图"

  Open-Url $s.url

  $ci = 0
  :companyLoop foreach ($wi in $workItems) {
    $ci++
    Write-Host ""
    Write-Host "=========================================="
    Write-Host "[$ci/$($workItems.Count)] $($wi.company)"
    Write-Host "  文件夹: $($wi.folderName)   图片数: $($wi.images.Count)"
    Write-Host "------------------------------------------"

    New-Item -ItemType Directory -Force -Path $wi.siteDir | Out-Null

    # 有些网站一家公司只搜公司名本身、但要截多张不同页面的图（例如股权穿透图 + 工商变更图），
    # 这时在 sites.json 里给该站加 "searchByCompanyOnly": true，搜索词就只复制一次。
    if ($s.searchByCompanyOnly) {
      $term = Get-SearchTerm $wi.company $s.stripPrefixes
      Set-Clipboard -Value $term
      Write-Host "  （已复制搜索词: $term）"
    }

    $mdPath = Join-Path $wi.siteDir "$($wi.folderName).md"
    $allExist = -not ($wi.images | Where-Object { -not (Test-Path -LiteralPath $_.imgPath) })
    if ($allExist -and (Test-Path -LiteralPath $mdPath) -and !$Overwrite) {
      Write-Host "  [跳过] 图片和 md 都已存在"
      $done += $wi.images.Count; $todo += $wi.images.Count
      continue
    }

    $ei = 0
    foreach ($img in $wi.images) {
      $ei++; $todo++

      if ((Test-Path -LiteralPath $img.imgPath) -and !$Overwrite) {
        Write-Host "  [$ei/$($wi.images.Count)] 已存在: $($img.imgName)"
        $done++
        continue
      }

      if (-not $s.searchByCompanyOnly) {
        $term = Get-SearchTerm $img.searchName $s.stripPrefixes
        Set-Clipboard -Value $term
      }

      Write-Host ""
      Write-Host "  [$ei/$($wi.images.Count)] $($img.text)"
      Write-Host "  → 文件名: $($img.imgName)"
      if (-not $s.searchByCompanyOnly) { Write-Host "  （搜索词已复制: $term）" }

      $ok = $false
      while (-not $ok) {
        $key = Read-Host "  [回车]=截 / r=重截 / o=重开网址 / s=跳过 / n=换下家 / q=退出"
        switch ($key.Trim().ToLower()) {
          "q" { Write-Host "已退出。"; break siteLoop }
          "n" { Write-Host "跳过本站剩余公司。"; break companyLoop }
          "s" { Write-Host "  已跳过"; $ok = $true }
          "o" { Open-Url $s.url }
          "r" { Save-Screenshot $img.imgPath; Write-Host "  已重截: $($img.imgName)" }
          default {
            Save-Screenshot $img.imgPath
            Write-Host "  已保存: $($img.imgName)"
            $ok = $true; $done++
          }
        }
      }
    }

    New-Markdown -Path $mdPath -Items $wi.images
    Write-Host "  -> 已生成: $($wi.folderName).md"

    $exporter = Join-Path $Here "export_word.ps1"
    if (Test-Path -LiteralPath $exporter) {
      & $exporter -MarkdownPath $mdPath -Force -Quiet
    }
  }
}

Write-Host ""
Write-Host "========== 完成 =========="
Write-Host "已截图: $done / 预计: $todo"
Write-Host "输出目录: $Out"
