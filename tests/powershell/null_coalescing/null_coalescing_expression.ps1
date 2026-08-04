# vybe-test: powershell/null_coalescing/null_coalescing_expression
$value = $null
$result = $value ?? (1 + 1)
if ($result -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
