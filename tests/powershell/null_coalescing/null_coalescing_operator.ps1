# vybe-test: powershell/null_coalescing/null_coalescing_operator
$value = $null
$result = $value ?? 'fallback'
if ($result -ne 'fallback') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
