# vybe-test: powershell/math/math_trig
$pi = [Math]::PI
$sin90 = [Math]::Round([Math]::Sin($pi / 2), 10)
$cos0  = [Math]::Round([Math]::Cos(0), 10)
if ($sin90 -ne 1.0) { Write-Host "FAIL: sin(90°) should be 1, got $sin90"; exit 1 }
if ($cos0  -ne 1.0) { Write-Host "FAIL: cos(0) should be 1, got $cos0";   exit 1 }
$tan45 = [Math]::Round([Math]::Tan($pi / 4), 10)
if ($tan45 -ne 1.0) { Write-Host "FAIL: tan(45°) should be 1, got $tan45"; exit 1 }
Write-Host "PASS"
exit 0
