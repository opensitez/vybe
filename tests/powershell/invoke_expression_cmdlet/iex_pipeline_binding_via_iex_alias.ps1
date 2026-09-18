# vybe-test: powershell/invoke_expression_cmdlet/iex_pipeline_binding_via_iex_alias
# The 'iex' alias passes a string through the pipeline into Invoke-Expression for evaluation
$result = "2 * 5" | Invoke-Expression

if ($result -ne 10) {
    Write-Host "FAIL: expected 10 from pipeline-bound Invoke-Expression, got: $result"
    exit 1
}

Write-Host "PASS"
exit 0
