# vybe-test: powershell/math_clamp_and_range_boundaries/sign_function_positive_negative_zero
$sPos = [math]::Sign(50)
$sNeg = [math]::Sign(-50)
$sZero = [math]::Sign(0)
if ($sPos -ne 1 -or $sNeg -ne -1 -or $sZero -ne 0) {
    Write-Host "FAIL: Sign function failed"
    exit 1
}
Write-Host "PASS"
exit 0
