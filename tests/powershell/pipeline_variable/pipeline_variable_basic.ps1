# vybe-test: powershell/pipeline_variable/pipeline_variable_basic
$res = 1..3 | ForEach-Object -PipelineVariable item { $_ * 10 } | ForEach-Object { "$item:$_" }
if ($res[0] -ne "1:10" -or $res[1] -ne "2:20" -or $res[2] -ne "3:30") {
    Write-Host "FAIL: PipelineVariable basic binding expected 1:10, 2:20, 3:30, got $($res -join ', ')"
    exit 1
}
Write-Host "PASS"
exit 0
