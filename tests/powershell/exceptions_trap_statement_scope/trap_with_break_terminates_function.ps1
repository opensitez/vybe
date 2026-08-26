# vybe-test: powershell/exceptions_trap_statement_scope/trap_with_break_terminates_function
$trapped = $false
function Test-TrapBreak {
    trap {
        $script:trapped = $true
        break
    }
    1 / 0
    return "ShouldNotReach"
}
$caught = $false
try {
    $res = Test-TrapBreak
} catch {
    $caught = $true
}
if (-not $trapped -or -not $caught) {
    Write-Host "FAIL: Trap with break should rethrow error out of scope"
    exit 1
}
Write-Host "PASS"
exit 0
