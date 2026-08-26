# vybe-test: powershell/type_half_precision_floats/half_negative_zero_check
$negZero = [System.Half]::NegativeZero
if (-not [System.Half]::IsNegative($negZero)) { Write-Host "FAIL: Half NegativeZero failed"; exit 1 }
Write-Host "PASS"; exit 0
