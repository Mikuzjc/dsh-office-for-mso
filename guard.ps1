# guard.ps1 — 编辑/提交前守卫：检测工作区相对 git HEAD 的未提交改动
# 背景：DSH 的 edit 工具自带 re-read 强制，但 pwsh/node 等脚本直接改文件时不经过该保护。
# 本脚本让 AI/人在编辑或提交前先跑一次，发现"可能被外部进程/其他 AI/脚本修改过"的文件。
# 用法：
#   powershell -ExecutionPolicy Bypass -File guard.ps1             # 检查整个工作区
#   powershell -ExecutionPolicy Bypass -File guard.ps1 server.js    # 只检查指定文件（可多个）
param([string[]]$Paths)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (-not (Test-Path '.git')) {
  Write-Host '[guard] 当前目录不是 git 仓库，跳过' -ForegroundColor Yellow
  exit 0
}

$changed = git status --porcelain 2>$null
if (-not $changed) {
  Write-Host '[guard] 工作区干净（无未提交改动）✅' -ForegroundColor Green
  exit 0
}

$lines = @($changed | Where-Object {
  if ($Paths.Count -eq 0) { $true }
  else {
    $p = (($_ -replace '^.. ', '')).Trim()
    [bool]($Paths | Where-Object { $p -like "*$_*" })
  }
})

if ($lines.Count -eq 0) {
  Write-Host '[guard] 指定文件无未提交改动 ✅' -ForegroundColor Green
  exit 0
}

Write-Host '[guard] ⚠️ 以下文件相对 git HEAD 有未提交改动（可能被外部进程/其他 AI/脚本修改过）：' -ForegroundColor Yellow
foreach ($l in $lines) {
  $status = $l.Substring(0, 2).Trim()
  $file = ($l -replace '^.. ', '').Trim()
  $full = Join-Path $PSScriptRoot $file
  $mtime = if (Test-Path $full) { (Get-Item $full).LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') } else { '(已删除)' }
  Write-Host ("  [{0}] {1}  (mtime {2})" -f $status, $file, $mtime) -ForegroundColor Red
}
Write-Host ''
Write-Host '→ 编辑这些文件前，请先 read 最新内容（re-read），确认改动来源后再继续。' -ForegroundColor Cyan
Write-Host '→ 提交前用 git diff / git status 核对改动。' -ForegroundColor Cyan
exit 1
