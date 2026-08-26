# vybe-test: powershell/exceptions_trap_statement_scope/trap_in_while_loop_continues_iteration
$trappedTimes = 0
function Test-TrapInWhile {
    trap { $script:trappedTimes++; continue }
    $i = 0
    while ($i -lt 3) {
        $i++
        if ($i -eq 2) { 1 / 0 }
    }
}
Test-TrapInWhile
if ($trappedTimes -ne 1) {
    Write-Host "FAIL: Trap in while loop failed, got $trappedTimes"
    exit 1
}
Write-Host "PASS"
exit 0
