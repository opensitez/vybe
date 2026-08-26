# vybe-test: powershell/math_fused_multiply_add/fused_multiply_add_case_2
$res = [math]::FusedMultiplyAdd([double]2, [double]2, [double]5)
if ($res -ne ([double]2 * 2.0 + 5.0)) { Write-Host "FAIL: FusedMultiplyAdd failed"; exit 1 }
Write-Host "PASS"; exit 0
