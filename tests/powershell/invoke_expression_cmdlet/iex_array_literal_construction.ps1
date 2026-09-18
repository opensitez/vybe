# vybe-test: powershell/invoke_expression_cmdlet/iex_array_literal_construction
# Invoke-Expression evaluates array literal syntax and returns an array object
$result = Invoke-Expression "@(1, 2, 3)"

if ($result.Count -ne 3 -or $result[2] -ne 3) {
    Write-Host "FAIL: array literal mismatch, Count=$($result.Count), [2]=$($result[2])"
    exit 1
}

Write-Host "PASS"
exit 0
