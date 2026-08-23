# start.ps1 — 前台启动 DSH Office Bridge 桥接服务
# 生产环境建议用计划任务（登录自启）托管：见 README「服务托管」
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
Write-Host '启动 DSH Office Bridge (http://127.0.0.1:3000) ...' -ForegroundColor Cyan
node server.js
