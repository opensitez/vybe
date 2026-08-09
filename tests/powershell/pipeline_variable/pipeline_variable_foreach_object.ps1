# vybe-test: powershell/pipeline_variable/pipeline_variable_foreach_object
$sum = 0
1..4 | ForEach-Object -PipelineVariable v { $_ * 2 } | ForEach-Object { $script:sum += $v }
if ($sum -ne 20) {
    Write-Host "FAIL: PipelineVariable accumulator expected 20, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
