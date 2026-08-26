# vybe-test: powershell/exceptions_throw_expression_vs_statement/throw_expression_in_function_pipeline_chain
$caught = $false
try {
    1..5 | ForEach-Object { $_ } | ForEach-Object { if ($_ -eq 3) { throw "MidPipeline" } }
} catch {
    $caught = $_.Exception.Message.Contains("MidPipeline")
}
if (-not $caught) {
    Write-Host "FAIL: Throw in multi-stage pipeline failed"
    exit 1
}
Write-Host "PASS"
exit 0
