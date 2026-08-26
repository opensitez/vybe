# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_in_pipeline_scriptblock
$caught = $false
try {
    1..5 | ForEach-Object { if ($_ -eq 3) { throw "PipelineError" } }
} catch {
    $caught = $_.Exception.Message.Contains("PipelineError")
}
if (-not $caught) {
    Write-Host "FAIL: Throw in pipeline scriptblock failed"
    exit 1
}
Write-Host "PASS"
exit 0
