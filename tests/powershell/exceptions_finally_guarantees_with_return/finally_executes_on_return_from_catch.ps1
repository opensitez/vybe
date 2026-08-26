# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_executes_on_return_from_catch
$finallyRan = $false
function Test-FinallyCatchReturn {
    try {
        throw "Err"
    } catch {
        return "Recovered"
    } finally {
        $script:finallyRan = $true
    }
}
$res = Test-FinallyCatchReturn
if ($res -ne "Recovered" -or -not $finallyRan) {
    Write-Host "FAIL: Finally on return from catch failed"
    exit 1
}
Write-Host "PASS"
exit 0
