# vybe-test: powershell/in_scope/local_scope_in_if
if ($true) {
    $x = 1
}
if ($x -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
