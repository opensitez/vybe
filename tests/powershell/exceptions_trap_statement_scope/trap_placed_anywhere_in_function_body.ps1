# vybe-test: powershell/exceptions_trap_statement_scope/trap_placed_anywhere_in_function_body
$trapped = $false
function Test-TrapAtBottom {
    1 / 0
    return "After"
    trap {
        $script:trapped = $true
        continue
    }
}
$res = Test-TrapAtBottom
if (-not $trapped -or $res -ne "After") {
    Write-Host "FAIL: Trap at bottom of function failed"
    exit 1
}
Write-Host "PASS"
exit 0
