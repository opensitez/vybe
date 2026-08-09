# vybe-test: powershell/pipeline_variable/pipeline_variable_type_retention
$res = @([int]10, [double]20.5) | ForEach-Object -PipelineVariable item { $_ } | ForEach-Object { $item.GetType().Name }
if ($res[0] -ne "Int32" -or $res[1] -ne "Double") {
    Write-Host "FAIL: PipelineVariable type retention expected Int32, Double"
    exit 1
}
Write-Host "PASS"
exit 0
