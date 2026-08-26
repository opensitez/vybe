# vybe-test: powershell/math_clamp_and_range_boundaries/copy_sign_method
$val = [math]::CopySign(3.0, -1.0)
if ($val -ne -3.0) {
    Write-Host "FAIL: CopySign expected -3.0, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
