# vybe-test: powershell/null_coalescing/null_coalescing_to_null
$value = $null
$result = $value ?? $null
if ($result -ne $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
