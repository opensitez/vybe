# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_throwing_exception_overrides_catch_exception
function Test-FinallyOverrideCatchEx {
    try {
        throw "TryException"
    } catch {
        throw "CatchException"
    } finally {
        throw "FinallyException"
    }
}
$caughtMsg = ""
try {
    Test-FinallyOverrideCatchEx
} catch {
    $caughtMsg = $_.Exception.Message
}
if (-not $caughtMsg.Contains("FinallyException")) {
    Write-Host "FAIL: Exception in finally should override catch exception, got '$caughtMsg'"
    exit 1
}
Write-Host "PASS"
exit 0
