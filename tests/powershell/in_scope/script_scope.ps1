# vybe-test: powershell/in_scope/script_scope
$script:x = 1
if ($x -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
