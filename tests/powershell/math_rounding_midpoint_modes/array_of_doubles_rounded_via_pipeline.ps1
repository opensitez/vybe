# vybe-test: powershell/math_rounding_midpoint_modes/array_of_doubles_rounded_via_pipeline
$arr = @(1.5, 2.5, 3.5, 4.5)
$rounded = @($arr | ForEach-Object { [math]::Round($_, [System.MidpointRounding]::AwayFromZero) })
if ($rounded[0] -ne 2.0 -or $rounded[1] -ne 3.0 -or $rounded[2] -ne 4.0 -or $rounded[3] -ne 5.0) {
    Write-Host "FAIL: Pipeline array rounding failed"
    exit 1
}
Write-Host "PASS"
exit 0
