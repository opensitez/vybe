# vybe-test: powershell/invoke_expression_cmdlet/iex_range_operator_expression
# Invoke-Expression evaluates the PowerShell range operator and returns an integer array
$result = Invoke-Expression "1..5"

if ($result.Count -ne 5 -or $result[0] -ne 1 -or $result[4] -ne 5) {
    Write-Host "FAIL: range expression mismatch, Count=$($result.Count), [0]=$($result[0]), [4]=$($result[4])"
    exit 1
}

Write-Host "PASS"
exit 0
