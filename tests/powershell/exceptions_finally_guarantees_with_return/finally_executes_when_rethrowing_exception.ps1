# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_executes_when_rethrowing_exception
$finallyRan = $false
function Test-FinallyRethrow {
    try {
        throw "RethrowMe"
    } catch {
        throw $_
    } finally {
        $script:finallyRan = $true
    }
}
$caught = $false
try {
    Test-FinallyRethrow
} catch {
    $caught = $true
}
if (-not $caught -or -not $finallyRan) {
    Write-Host "FAIL: Finally must run on rethrow"
    exit 1
}
Write-Host "PASS"
exit 0
