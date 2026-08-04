# vybe-test: powershell/conditionals/if_else_basic
if (1 -eq 1) {
    $result = 'yes'
} else {
    $result = 'no'
}
if ($result -ne 'yes') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
