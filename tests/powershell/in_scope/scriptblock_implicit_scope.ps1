# vybe-test: powershell/in_scope/scriptblock_implicit_scope
& { $x = 5 }
if ($x -ne $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
