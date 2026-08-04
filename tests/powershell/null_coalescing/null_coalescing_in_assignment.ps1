# vybe-test: powershell/null_coalescing/null_coalescing_in_assignment
$value = $null
$value = $value ?? 'default'
if ($value -ne 'default') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
