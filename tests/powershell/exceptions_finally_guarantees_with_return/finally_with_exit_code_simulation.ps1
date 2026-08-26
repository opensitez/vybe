# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_with_exit_code_simulation
$cleaned = $false
function Safe-Process {
    try {
        return 0
    } finally {
        $script:cleaned = $true
    }
}
$code = Safe-Process
if ($code -ne 0 -or -not $cleaned) {
    Write-Host "FAIL: Safe exit code in finally failed"
    exit 1
}
Write-Host "PASS"
exit 0
