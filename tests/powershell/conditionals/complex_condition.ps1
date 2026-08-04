# vybe-test: powershell/conditionals/complex_condition
if ((1 -eq 1) -and (2 -eq 2)) {
    $result = 'ok'
} else {
    $result = 'fail'
}
if ($result -ne 'ok') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
