# vybe-test: powershell/arrays/array_unary_comma_single
# Unary comma creates a single-element array
$single = ,42
if ($single -isnot [array])  { Write-Host "FAIL: should be array"; exit 1 }
if ($single.Count -ne 1)     { Write-Host "FAIL: count should be 1"; exit 1 }
if ($single[0] -ne 42)       { Write-Host "FAIL: value"; exit 1 }
Write-Host "PASS"
exit 0
