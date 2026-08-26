# vybe-test: powershell/math_hypot_and_cbrt/cbrt_and_hypot_case_7
$cbrt = [math]::Cbrt(27.0)
$cbrt8 = [math]::Cbrt(8.0)
if ([math]::Abs($cbrt - 3.0) -gt 1e-6 -or [math]::Abs($cbrt8 - 2.0) -gt 1e-6) { Write-Host "FAIL: Cbrt failed"; exit 1 }
Write-Host "PASS"; exit 0
