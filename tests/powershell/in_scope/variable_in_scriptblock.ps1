# vybe-test: powershell/in_scope/variable_in_scriptblock
& { $x = 3 }
if ($x -ne $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
