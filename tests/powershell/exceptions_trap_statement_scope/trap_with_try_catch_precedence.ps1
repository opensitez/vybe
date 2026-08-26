# vybe-test: powershell/exceptions_trap_statement_scope/trap_with_try_catch_precedence
$catchRan = $false
$trapRan = $false
function Test-TryCatchTrapPrecedence {
    trap { $script:trapRan = $true; continue }
    try {
        1 / 0
    } catch {
        $script:catchRan = $true
    }
}
Test-TryCatchTrapPrecedence
if (-not $catchRan -or $trapRan) {
    Write-Host "FAIL: try-catch should take precedence over trap"
    exit 1
}
Write-Host "PASS"
exit 0
