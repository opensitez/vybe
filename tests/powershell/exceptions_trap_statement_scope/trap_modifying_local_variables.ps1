# vybe-test: powershell/exceptions_trap_statement_scope/trap_modifying_local_variables
$trapped = $false
function Test-TrapModifyLocal {
    trap {
        $script:trapped = $true
        continue
    }
    1 / 0
}
Test-TrapModifyLocal
if (-not $trapped) {
    Write-Host "FAIL: Trap modifying local variable failed"
    exit 1
}
Write-Host "PASS"
exit 0
