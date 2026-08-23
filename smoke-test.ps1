# smoke-test.ps1 — 冒烟测试：检查桥接服务与关键端点
# 用法：powershell -ExecutionPolicy Bypass -File smoke-test.ps1
$base = 'http://127.0.0.1:3000'
$fail = 0

try {
  $s = Invoke-RestMethod "$base/office/status" -TimeoutSec 5
  $online = $s.hosts.PSObject.Properties.Name -join ','
  if ($online) { Write-Host "[OK] 在线文档: $online" -ForegroundColor Green }
  else { Write-Host "[WARN] 无文档在线（确认 Word/Excel/PPT 已打开且窗格已加载）" -ForegroundColor Yellow }
} catch { Write-Host "[FAIL] /office/status: $($_.Exception.Message)" -ForegroundColor Red; $fail++ }

try {
  $c = Invoke-RestMethod "$base/office/capabilities" -TimeoutSec 5
  Write-Host "[OK] capabilities: $($c.actions.Count) 个 action" -ForegroundColor Green
} catch { Write-Host "[FAIL] /office/capabilities: $($_.Exception.Message)" -ForegroundColor Red; $fail++ }

try {
  $v = Invoke-RestMethod "$base/office/actions-version" -TimeoutSec 5
  Write-Host "[OK] actions-version: $($v.version)" -ForegroundColor Green
} catch { Write-Host "[FAIL] /office/actions-version: $($_.Exception.Message)" -ForegroundColor Red; $fail++ }

try {
  $j = Invoke-WebRequest "$base/taskpane.html" -UseBasicParsing -TimeoutSec 5
  Write-Host "[OK] taskpane.html: HTTP $($j.StatusCode)" -ForegroundColor Green
} catch { Write-Host "[FAIL] taskpane.html: $($_.Exception.Message)" -ForegroundColor Red; $fail++ }

if ($fail -eq 0) { Write-Host "`n冒烟测试全部通过 ✅" -ForegroundColor Green; exit 0 }
else { Write-Host "`n冒烟测试 $fail 项失败 ❌" -ForegroundColor Red; exit 1 }
