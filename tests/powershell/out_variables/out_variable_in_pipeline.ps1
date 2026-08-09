# vybe-test: powershell/out_variables/out_variable_in_pipeline
$pipelineResult = 1..3 | ForEach-Object { $_ * 2 } -OutVariable mid | ForEach-Object { $_ + 1 }
if ($mid.Count -ne 3 -or $mid[0] -ne 2 -or $pipelineResult[0] -ne 3) {
    Write-Host "FAIL: mid-pipeline OutVariable capture failed"
    exit 1
}
Write-Host "PASS"
exit 0
