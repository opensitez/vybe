# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_throwing_exception_overrides_try_return
function Test-FinallyThrow {
    try {
        return "TryValue"
    } finally {
        throw "FinallyException"
    }
}
$caught = $false
try {
    $x = Test-FinallyThrow
} catch {
    $caught = $_.Exception.Message.Contains("FinallyException")
}
if (-not $caught) {
    Write-Host "FAIL: Exception thrown in finally should override try return"
    exit 1
}
Write-Host "PASS"
exit 0
