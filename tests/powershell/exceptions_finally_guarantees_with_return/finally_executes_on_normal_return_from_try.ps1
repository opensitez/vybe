# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_executes_on_normal_return_from_try
$finallyRan = $false
function Test-FinallyReturn {
    try {
        return "EarlyReturn"
    } finally {
        $script:finallyRan = $true
    }
}
$res = Test-FinallyReturn
if ($res -ne "EarlyReturn" -or -not $finallyRan) {
    Write-Host "FAIL: Finally on normal return from try failed, res='$res', finallyRan=$finallyRan"
    exit 1
}
Write-Host "PASS"
exit 0
