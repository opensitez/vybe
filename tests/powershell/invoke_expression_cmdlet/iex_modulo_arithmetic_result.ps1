# vybe-test: powershell/invoke_expression_cmdlet/iex_modulo_arithmetic_result
# Invoke-Expression correctly computes the modulo (%) operator in an expression string
$result = Invoke-Expression "17 % 5"

if ($result -ne 2) {
    Write-Host "FAIL: expected 2 from '17 % 5', got: $result"
    exit 1
}

Write-Host "PASS"
exit 0
