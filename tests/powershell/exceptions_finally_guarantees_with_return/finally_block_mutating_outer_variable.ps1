# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_block_mutating_outer_variable
$state = "BEFORE"
function Test-FinallyMutate {
    try {
        return "DONE"
    } finally {
        $script:state = "AFTER"
    }
}
$res = Test-FinallyMutate
if ($res -ne "DONE" -or $state -ne "AFTER") {
    Write-Host "FAIL: Finally mutating outer variable failed"
    exit 1
}
Write-Host "PASS"
exit 0
