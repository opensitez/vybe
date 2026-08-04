# vybe-test: powershell/null_coalescing/null_coalescing_hashtable
$value = $null
$result = $value ?? @{ a=1 }
if ($result.a -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
