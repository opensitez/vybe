# vybe-test: powershell/math/math_clamp
# [Math]::Clamp available in .NET 5+
$low  = [Math]::Clamp(-5, 0, 100)
$high = [Math]::Clamp(200, 0, 100)
$mid  = [Math]::Clamp(42, 0, 100)
if ($low  -ne 0)   { Write-Host "FAIL: clamp low"; exit 1 }
if ($high -ne 100) { Write-Host "FAIL: clamp high"; exit 1 }
if ($mid  -ne 42)  { Write-Host "FAIL: clamp mid"; exit 1 }
Write-Host "PASS"
exit 0
