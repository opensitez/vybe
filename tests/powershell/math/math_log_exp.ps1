# vybe-test: powershell/math/math_log_exp
$e  = [Math]::E
$ln = [Math]::Log($e)
$rounded = [Math]::Round($ln, 5)
if ($rounded -ne 1.0) { Write-Host "FAIL: ln(e) should be 1, got $rounded"; exit 1 }
$log10 = [Math]::Log10(1000)
if ($log10 -ne 3.0) { Write-Host "FAIL: log10(1000) should be 3"; exit 1 }
Write-Host "PASS"
exit 0
