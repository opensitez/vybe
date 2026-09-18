# vybe-test: powershell/invoke_expression_cmdlet/iex_preserves_result_type
# Invoke-Expression preserves the .NET type of the evaluated expression result
$result = Invoke-Expression "2 + 2"

if ($result.GetType().Name -ne "Int32") {
    Write-Host "FAIL: expected Int32 result type, got: $($result.GetType().Name)"
    exit 1
}

Write-Host "PASS"
exit 0
