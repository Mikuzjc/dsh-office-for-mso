# sideload.ps1 — 一键注册/移除 DSH Office 加载项（WEF 注册表方式，跨机器通用，无需手动点 Office 菜单）
# 用法：
#   powershell -ExecutionPolicy Bypass -File sideload.ps1            # 注册加载项
#   powershell -ExecutionPolicy Bypass -File sideload.ps1 -Remove     # 移除加载项
param([switch]$Remove)

$ErrorActionPreference = 'Stop'
$manifest = Join-Path $PSScriptRoot 'manifest.xml'
if (-not (Test-Path $manifest)) {
  Write-Host "[FAIL] 找不到 manifest.xml: $manifest" -ForegroundColor Red
  exit 1
}

# 从 manifest.xml 提取加载项 GUID（<Id>）
$content = Get-Content -Raw $manifest
if ($content -match '<Id>\s*([0-9a-fA-F-]{36})\s*</Id>') {
  $guid = $matches[1]
} else {
  Write-Host "[FAIL] manifest.xml 中未找到 <Id> GUID" -ForegroundColor Red
  exit 1
}

$devKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (-not (Test-Path $devKey)) { New-Item -Path $devKey -Force | Out-Null }

if ($Remove) {
  Remove-ItemProperty -Path $devKey -Name $guid -ErrorAction SilentlyContinue
  Write-Host "[OK] 已移除加载项注册: $guid" -ForegroundColor Green
  Write-Host "     （请关闭并重新打开 Office 文档后生效）" -ForegroundColor Yellow
  exit 0
}

New-ItemProperty -Path $devKey -Name $guid -Value $manifest -PropertyType String -Force | Out-Null
$check = (Get-ItemProperty -Path $devKey).$guid
if ($check -eq $manifest) {
  Write-Host "[OK] 加载项已注册: $guid" -ForegroundColor Green
  Write-Host "     manifest: $manifest" -ForegroundColor Gray
  Write-Host "请关闭并重新打开 Office 文档（或完全退出 Office 后重开），" -ForegroundColor Yellow
  Write-Host "侧边栏将自动出现「DSH Office 执行器」窗格（显示 已连接：等待 DSH 指令）。" -ForegroundColor Yellow
} else {
  Write-Host "[FAIL] 注册校验失败" -ForegroundColor Red
  exit 1
}
