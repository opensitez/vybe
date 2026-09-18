# vybe-test: powershell/invoke_expression_cmdlet/iex_evaluates_pipeline_expression
# Invoke-Expression handles pipeline expressions embedded within the command string
$result = Invoke-Expression '@("alpha", "beta", "gamma") | Where-Object { $_ -ne "beta" }'

if ($result.Count -ne 2 -or ($result -join ",") -ne "alpha,gamma") {
    Write-Host "FAIL: pipeline result mismatch: @($($result -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
