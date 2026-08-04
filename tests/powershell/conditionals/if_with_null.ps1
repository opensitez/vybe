# vybe-test: powershell/conditionals/if_with_null
$value = $null
if ($value -eq $null) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
