# vybe-test: powershell/math_logarithmic_and_exponential/cbrt_cube_root_function
$cb1 = [math]::Cbrt(27.0)
$cb2 = [math]::Cbrt(-8.0)
if ([math]::Abs($cb1 - 3.0) -gt 1e-12 -or [math]::Abs($cb2 - (-2.0)) -gt 1e-12) {
    Write-Host "FAIL: Cbrt cube root failed"
    exit 1
}
Write-Host "PASS"
exit 0
