# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_executes_on_break_statement
$finallyRan = $false
function Test-FinallyBreak {
    while ($true) {
        try {
            break
        } finally {
            $script:finallyRan = $true
        }
    }
}
Test-FinallyBreak
if (-not $finallyRan) {
    Write-Host "FAIL: Finally block must run on break statement"
    exit 1
}
Write-Host "PASS"
exit 0
