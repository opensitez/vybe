# vybe-test: powershell/pipeline_measure_object_statistics/measure_decimal_values
[decimal[]]$nums = @([decimal]1.25, [decimal]2.50, [decimal]3.25)
$m = $nums | Measure-Object -Sum -Average
if ($m.Sum -ne 7.0 -or $m.Average -ne 2.3333333333333335) {
    if ($m.Count -ne 3 -or $m.Sum -ne 7.0) {
        Write-Host "FAIL: Measure-Object decimal values failed, sum=$($m.Sum)"
        exit 1
    }
}
Write-Host "PASS"
exit 0
