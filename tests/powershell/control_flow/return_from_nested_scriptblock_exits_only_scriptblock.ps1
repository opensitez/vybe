# vybe-test: powershell/control_flow/return_from_nested_scriptblock_exits_only_scriptblock
# A return statement inside an invoked scriptblock terminates the scriptblock but not the caller function
function InvokeNestedBlock {
    $blockResult = & {
        $x = 10
        return ($x * 2)
        $x = 999  # Unreachable
    }

    $callerContinuation = "caller-resumed"
    return @($blockResult, $callerContinuation)
}

$results = InvokeNestedBlock

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 returned values, got $($results.Count)"
    exit 1
}

if ($results[0] -ne 20) {
    Write-Host "FAIL: expected block return 20, got $($results[0])"
    exit 1
}

if ($results[1] -ne "caller-resumed") {
    Write-Host "FAIL: caller was prematurely terminated, got $($results[1])"
    exit 1
}

Write-Host "PASS"
exit 0
