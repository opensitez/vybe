# vybe-test: powershell/type_conversion/parse_int_tryparse
$ok = [int]::TryParse("42", [ref]$null)
if (-not $ok) { Write-Host "FAIL: '42' should parse"; exit 1 }
$fail = [int]::TryParse("abc", [ref]$null)
if ($fail) { Write-Host "FAIL: 'abc' should not parse"; exit 1 }
Write-Host "PASS"
exit 0
