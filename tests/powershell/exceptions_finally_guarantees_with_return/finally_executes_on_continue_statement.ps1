# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_executes_on_continue_statement
$finallyCount = 0
for ($i = 0; $i -lt 3; $i++) {
    try {
        continue
    } finally {
        $finallyCount++
    }
}
if ($finallyCount -ne 3) {
    Write-Host "FAIL: Finally block must run on continue statement, count=$finallyCount"
    exit 1
}
Write-Host "PASS"
exit 0
