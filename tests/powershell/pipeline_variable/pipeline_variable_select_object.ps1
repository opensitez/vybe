# vybe-test: powershell/pipeline_variable/pipeline_variable_select_object
$res = 10..12 | ForEach-Object -PipelineVariable p { $_ } | Select-Object @{N="PVar"; E={$p}}, @{N="Val"; E={$_}}
if ($res[0].PVar -ne 10 -or $res[2].Val -ne 12) {
    Write-Host "FAIL: Select-Object calculated property via PipelineVariable failed"
    exit 1
}
Write-Host "PASS"
exit 0
