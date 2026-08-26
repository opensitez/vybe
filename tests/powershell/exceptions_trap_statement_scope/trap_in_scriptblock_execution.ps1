# vybe-test: powershell/exceptions_trap_statement_scope/trap_in_scriptblock_execution
$trapped = $false
$sb = {
    trap { $script:trapped = $true; continue }
    throw "SbTrap"
}
& $sb
if (-not $trapped) {
    Write-Host "FAIL: Trap in scriptblock failed"
    exit 1
}
Write-Host "PASS"
exit 0
