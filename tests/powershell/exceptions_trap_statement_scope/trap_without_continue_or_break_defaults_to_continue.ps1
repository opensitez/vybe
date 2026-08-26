# vybe-test: powershell/exceptions_trap_statement_scope/trap_without_continue_or_break_defaults_to_continue
$trapped = $false
function Test-DefaultTrap {
    trap {
        $script:trapped = $true
        # implicit continue
    }
    1 / 0
    return "Done"
}
$res = Test-DefaultTrap
if (-not $trapped) {
    Write-Host "FAIL: Default trap execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
