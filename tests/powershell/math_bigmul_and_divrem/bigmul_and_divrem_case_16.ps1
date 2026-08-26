# vybe-test: powershell/math_bigmul_and_divrem/bigmul_and_divrem_case_16
$big = [math]::BigMul([int32]10000, [int32]20000)
$t = [math]::DivRem([int64]25, [int64]4)
$quotient = if ($t.Quotient -ne $null) { $t.Quotient } else { $t.Item1 }
if ($big -ne 200000000 -or $quotient -ne 6) { Write-Host "FAIL: BigMul/DivRem failed"; exit 1 }
Write-Host "PASS"; exit 0
