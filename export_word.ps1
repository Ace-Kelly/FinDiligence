<#
.SYNOPSIS
  把截图目录里的 Markdown 通过 Pandoc 导出为同名 Word。

.DESCRIPTION
  run.ps1 每截完一家公司会自动调用本脚本。也可以单独运行，批量补导整个目录。
  没装 Pandoc 不影响截图和 Markdown，只是不会生成 docx。

.EXAMPLE
  .\export_word.ps1 -Root .\截图输出_20260721
  .\export_word.ps1 -Root .\截图输出_20260721 -Force
  .\export_word.ps1 -MarkdownPath ".\截图输出_20260721\示例集团\01-信用中国\01-信用中国.md"
#>

[CmdletBinding()]
param(
  [string]$MarkdownPath = "",
  [string]$Root = "",
  [switch]$Force,
  [switch]$Quiet
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Write-Info([string]$Message) {
  if (-not $Quiet) { Write-Host $Message }
}

$pandoc = Get-Command pandoc -ErrorAction SilentlyContinue
if (!$pandoc) {
  $msg = "未找到 Pandoc，跳过 Word 导出。安装：https://pandoc.org/installing.html"
  if ($Quiet) { Write-Warning $msg; return }
  throw $msg
}

if ($MarkdownPath) {
  if (!(Test-Path -LiteralPath $MarkdownPath -PathType Leaf)) { throw "找不到 Markdown 文件: $MarkdownPath" }
  $files = @(Get-Item -LiteralPath $MarkdownPath)
} else {
  if (!$Root) { throw "请用 -Root 指定截图输出目录，或用 -MarkdownPath 指定单个 md 文件。" }
  if (!(Test-Path -LiteralPath $Root -PathType Container)) { throw "找不到目录: $Root" }
  $files = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Filter "*.md")
}

if ($files.Count -eq 0) { Write-Info "没有需要导出的 Markdown。"; return }

foreach ($md in $files) {
  $docx = [System.IO.Path]::ChangeExtension($md.FullName, ".docx")

  if ((Test-Path -LiteralPath $docx) -and !$Force) {
    if ((Get-Item -LiteralPath $docx).LastWriteTimeUtc -ge $md.LastWriteTimeUtc) {
      Write-Info "已是最新，跳过: $([System.IO.Path]::GetFileName($docx))"
      continue
    }
  }

  # --resource-path 指到 md 所在目录，图片才找得到
  & $pandoc.Source "--from=gfm" "--resource-path=$($md.Directory.FullName)" `
    $md.FullName "--output=$docx"

  if ($LASTEXITCODE -ne 0) { throw "Pandoc 导出失败（退出码 $LASTEXITCODE）: $($md.FullName)" }
  if (!(Test-Path -LiteralPath $docx -PathType Leaf)) { throw "Pandoc 没有生成文件: $docx" }
  if ((Get-Item -LiteralPath $docx).Length -le 0) { throw "生成的 Word 文件为空: $docx" }

  Write-Info "已导出: $docx"
}
