# vybe-test: powershell/math_fused_multiply_add/fused_multiply_add_case_19
$res = [math]::FusedMultiplyAdd([double]19, [double]2, [double]5)
if ($res -ne ([double]19 * 2.0 + 5.0)) { Write-Host "FAIL: FusedMultiplyAdd failed"; exit 1 }
Write-Host "PASS"; exit 0
