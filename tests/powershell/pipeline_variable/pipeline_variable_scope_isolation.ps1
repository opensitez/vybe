# vybe-test: powershell/pipeline_variable/pipeline_variable_scope_isolation
$pv = "Outer"
1..1 | ForEach-Object -PipelineVariable pv { "Inner" } | Out-Null
if ($pv -ne "Outer") {
    Write-Host "FAIL: PipelineVariable leaked into caller scope"
    exit 1
}
Write-Host "PASS"
exit 0
