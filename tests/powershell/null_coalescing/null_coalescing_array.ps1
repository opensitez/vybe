# vybe-test: powershell/null_coalescing/null_coalescing_array
$value = $null
$result = $value ?? @(1,2)
if ($result.Count -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
