# vybe-test: powershell/pipeline_variable/pipeline_variable_nested
$res = @("X", "Y") | ForEach-Object -PipelineVariable parent {
    1..2 | ForEach-Object { "$parent-$_" }
}
if ($res.Count -ne 4 -or $res[0] -ne "X-1" -or $res[3] -ne "Y-2") {
    Write-Host "FAIL: nested PipelineVariable expansion expected 4 items, got $($res -join ', ')"
    exit 1
}
Write-Host "PASS"
exit 0
