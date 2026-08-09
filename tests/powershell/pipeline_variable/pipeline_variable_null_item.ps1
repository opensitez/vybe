# vybe-test: powershell/pipeline_variable/pipeline_variable_null_item
$res = @($null, "Value") | ForEach-Object -PipelineVariable nv { $_ } | ForEach-Object { if ($nv -eq $null) { "WAS_NULL" } else { $nv } }
if ($res[0] -ne "WAS_NULL" -or $res[1] -ne "Value") {
    Write-Host "FAIL: null PipelineVariable item expected WAS_NULL, Value"
    exit 1
}
Write-Host "PASS"
exit 0
