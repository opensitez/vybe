# vybe-test: powershell/exceptions_trap_statement_scope/trap_rethrow_via_throw_statement_inside_trap
$outerCaught = $false
function Test-TrapRethrow {
    trap {
        throw "RethrownFromTrap"
    }
    1 / 0
}
try {
    Test-TrapRethrow
} catch {
    $outerCaught = $_.Exception.Message.Contains("RethrownFromTrap")
}
if (-not $outerCaught) {
    Write-Host "FAIL: Throw statement inside trap failed"
    exit 1
}
Write-Host "PASS"
exit 0
