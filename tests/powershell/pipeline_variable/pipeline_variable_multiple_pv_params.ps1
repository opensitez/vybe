# vybe-test: powershell/pipeline_variable/pipeline_variable_multiple_pv_params
$list = [System.Collections.Generic.List[string]]::new()
1..3 | ForEach-Object -PipelineVariable num { $_ } | ForEach-Object {
    $list.Add("Item:$num")
}
if ($list.Count -ne 3 -or $list[0] -ne "Item:1") {
    Write-Host "FAIL: PipelineVariable failed"
    exit 1
}
Write-Host "PASS"
exit 0
