# update.ps1 — 更新到最新版：git pull + 重启桥接服务
# 用法：powershell -ExecutionPolicy Bypass -File update.ps1
# 说明：actions.js 改动重启服务后窗格自动热更新（无需重开窗格）；
#       外壳（taskpane.js/html）改动需重开窗格；server.js 改动本脚本已处理（重启服务）。
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

Write-Host '[update] git pull...' -ForegroundColor Cyan
git pull --rebase 2>&1 | ForEach-Object { Write-Host $_ }
if ($LASTEXITCODE -ne 0) { Write-Host '[update] git pull 失败（可能有本地未提交改动或冲突），请检查' -ForegroundColor Red; exit 1 }

Write-Host '[update] 重启桥接服务...' -ForegroundColor Cyan
$l = netstat -ano | Select-String '127.0.0.1:3000' | Select-String 'LISTENING'
if ($l) {
  $parts = ($l.ToString() -split '\s+' | Where-Object { $_ })
  $pid3000 = [int]$parts[$parts.Count - 1]
  Stop-Process -Id $pid3000 -Force
  Start-Sleep -Seconds 2
}
Start-ScheduledTask -TaskName 'DSH Office Bridge' 2>&1 | Out-Null
Start-Sleep -Seconds 3

try {
  $s = Invoke-RestMethod 'http://127.0.0.1:3000/office/status' -TimeoutSec 5
  Write-Host "[update] ✅ 服务已重启，在线文档: $($s.hosts.PSObject.Properties.Name -join ', ')" -ForegroundColor Green
} catch {
  Write-Host '[update] 服务状态未确认（请检查 http://127.0.0.1:3000/office/status）' -ForegroundColor Yellow
}
Write-Host '[update] 完成。若本机 Office 文档开着，actions 改动已自动热更新；外壳改动请重开窗格。' -ForegroundColor Gray
