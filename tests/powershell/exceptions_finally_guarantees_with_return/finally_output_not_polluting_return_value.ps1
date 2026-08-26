# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_output_not_polluting_return_value
function Test-CleanFinallyNoOutput {
    try {
        return 42
    } finally {
        $null = 1 + 1 # non-emitting statement
    }
}
$res = Test-CleanFinallyNoOutput
if ($res -ne 42) {
    Write-Host "FAIL: Clean finally without output pollution failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
