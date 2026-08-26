# vybe-test: powershell/exceptions_trap_statement_scope/trap_basic_execution_with_continue
$trapped = $false
function Test-TrapBasic {
    trap {
        $script:trapped = $true
        continue
    }
    1 / 0
    return "AfterTrap"
}
$res = Test-TrapBasic
if (-not $trapped -or $res -ne "AfterTrap") {
    Write-Host "FAIL: Trap basic execution with continue failed, trapped=$trapped, res='$res'"
    exit 1
}
Write-Host "PASS"
exit 0
