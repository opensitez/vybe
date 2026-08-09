# vybe-test: powershell/pipeline_variable/pipeline_variable_array_unrolling
$groups = @(@(1, 2), @(3, 4))
$res = $groups | ForEach-Object -PipelineVariable grp { $_ } | ForEach-Object { $grp.Count }
if ($res[0] -ne 2 -or $res[1] -ne 2) {
    Write-Host "FAIL: PipelineVariable array item Count expected 2, 2"
    exit 1
}
Write-Host "PASS"
exit 0
