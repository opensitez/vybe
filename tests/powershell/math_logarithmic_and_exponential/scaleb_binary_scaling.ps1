# vybe-test: powershell/math_logarithmic_and_exponential/scaleb_binary_scaling
$scaled = [math]::ScaleB(1.5, 3) # 1.5 * 2^3 = 12.0
if ($scaled -ne 12.0) {
    Write-Host "FAIL: ScaleB(1.5, 3) expected 12.0, got $scaled"
    exit 1
}
Write-Host "PASS"
exit 0
