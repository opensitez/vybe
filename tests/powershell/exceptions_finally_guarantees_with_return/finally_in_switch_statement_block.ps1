# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_in_switch_statement_block
$finallyRan = $false
switch (1) {
    1 {
        try {
            "Match1"
        } finally {
            $finallyRan = $true
        }
    }
}
if (-not $finallyRan) {
    Write-Host "FAIL: Finally inside switch statement clause failed"
    exit 1
}
Write-Host "PASS"
exit 0
