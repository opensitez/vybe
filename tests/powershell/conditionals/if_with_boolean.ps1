# vybe-test: powershell/conditionals/if_with_boolean
$flag = $true
if ($flag) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
