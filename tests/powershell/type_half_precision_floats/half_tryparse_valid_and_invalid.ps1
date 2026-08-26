# vybe-test: powershell/type_half_precision_floats/half_tryparse_valid_and_invalid
$val = [System.Half]::MinValue
$ok1 = [System.Half]::TryParse("1.5", [ref]$val)
$ok2 = [System.Half]::TryParse("not_a_num", [ref]$val)
if (-not $ok1 -or $ok2) { Write-Host "FAIL: Half TryParse failed"; exit 1 }
Write-Host "PASS"; exit 0
