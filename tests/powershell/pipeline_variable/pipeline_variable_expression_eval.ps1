# vybe-test: powershell/pipeline_variable/pipeline_variable_expression_eval
$res = 5..6 | ForEach-Object -PipelineVariable n { $n * $n } | ForEach-Object { "$n->$($_)" }
if ($res[0] -ne "5->25" -or $res[1] -ne "6->36") {
    Write-Host "FAIL: PipelineVariable expression evaluation expected 5->25, 6->36"
    exit 1
}
Write-Host "PASS"
exit 0
