# vybe-test: powershell/pipeline_variable/pipeline_variable_multi_stage
$out = 1..2 | ForEach-Object -PipelineVariable a { $_ } | ForEach-Object -PipelineVariable b { $_ * 5 } | ForEach-Object { "$a-$b-$_" }
if ($out[0] -ne "1-5-5" -or $out[1] -ne "2-10-10") {
    Write-Host "FAIL: multi-stage PipelineVariable expected 1-5-5, 2-10-10, got $($out -join ', ')"
    exit 1
}
Write-Host "PASS"
exit 0
