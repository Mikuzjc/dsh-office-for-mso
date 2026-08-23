# install.ps1 — 一键安装 DSH Office Bridge 服务（自动按当前目录配置，跨机器通用）
# 用法：powershell -ExecutionPolicy Bypass -File install.ps1
# 效果：注册计划任务「DSH Office Bridge」（用户登录时静默自启）并立即启动服务
# 注意：注册计划任务需要管理员权限，脚本会自动弹出 UAC 提权，请点击「是」
$ErrorActionPreference = 'Stop'

# 管理员权限检测 + 自动提权（UAC）
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
  Write-Host '需要管理员权限，正在请求提升（请在弹出的 UAC 对话框中点击「是」）...' -ForegroundColor Yellow
  Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`""
  exit
}

$taskName = 'DSH Office Bridge'
$projectDir = $PSScriptRoot
$nodePath = (Get-Command node -ErrorAction Stop).Source
$serverPath = Join-Path $projectDir 'server.js'

# 生成 run-hidden.vbs（静默启动用；node 用完整路径 + cmd /c 包装，避免计划任务环境 PATH 缺失导致"找不到文件"）
$vbsPath = Join-Path $projectDir 'run-hidden.vbs'
if (-not (Test-Path $vbsPath)) {
  Set-Content -Path $vbsPath -Value "Set sh = CreateObject(`"WScript.Shell`")`r`nsh.Run `"cmd /c `" & WScript.Arguments(0), 0, False" -Encoding ASCII
}

$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$vbsPath`" `"`"$nodePath`" `"$serverPath`"`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Write-Host "[OK] 计划任务已注册: $taskName" -ForegroundColor Green
Write-Host "     启动命令: wscript `"$vbsPath`" `"$nodePath $serverPath`"" -ForegroundColor Gray

Start-ScheduledTask -TaskName $taskName
Start-Sleep -Seconds 2

try {
  $s = Invoke-RestMethod 'http://127.0.0.1:3000/office/status' -TimeoutSec 5
  Write-Host "[OK] 桥接服务已启动，在线文档: $($s.hosts.PSObject.Properties.Name -join ', ')" -ForegroundColor Green
} catch {
  Write-Host "[WARN] 服务已启动但状态未确认（请稍后检查 http://127.0.0.1:3000/office/status）" -ForegroundColor Yellow
}
