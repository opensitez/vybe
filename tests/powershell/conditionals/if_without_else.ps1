# vybe-test: powershell/conditionals/if_without_else
if (1 -eq 2) {
    $result = 'a'
}
if ($result -ne $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
