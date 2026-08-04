# vybe-test: powershell/null_coalescing/null_coalescing_with_value
$value = 'x'
$result = $value ?? 'fallback'
if ($result -ne 'x') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
