# vybe-test: powershell/null_coalescing/null_coalescing_variable
$val = 'ok'
$result = $val ?? 'fallback'
if ($result -ne 'ok') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
