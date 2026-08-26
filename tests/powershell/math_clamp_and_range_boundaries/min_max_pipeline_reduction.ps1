# vybe-test: powershell/math_clamp_and_range_boundaries/min_max_pipeline_reduction
$arr = @(45, 12, 89, 34, 67)
$min = $arr[0]
$max = $arr[0]
foreach ($x in $arr) {
    $min = [math]::Min($min, $x)
    $max = [math]::Max($max, $x)
}
if ($min -ne 12 -or $max -ne 89) {
    Write-Host "FAIL: Min/Max loop reduction failed"
    exit 1
}
Write-Host "PASS"
exit 0
