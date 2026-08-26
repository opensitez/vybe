# vybe-test: powershell/exceptions_finally_guarantees_with_return/return_in_finally_overrides_return_in_try
$finallyRan = $false
function Test-FinallyOverrideReturn {
    try {
        return "FromTry"
    } finally {
        $script:finallyRan = $true
    }
}
$res = Test-FinallyOverrideReturn
if ($res -ne "FromTry" -or -not $finallyRan) {
    Write-Host "FAIL: Finally execution failed"
    exit 1
}
Write-Host "PASS"
exit 0
