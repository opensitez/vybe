# vybe-test: powershell/exceptions_custom_net_exception_classes/custom_exception_in_finally_block_execution_guarantee
class FinallyTriggerEx : System.Exception {}
$finallyRan = $false
try {
    throw [FinallyTriggerEx]::new()
} catch [FinallyTriggerEx] {
    # handled
} finally {
    $finallyRan = $true
}
if (-not $finallyRan) {
    Write-Host "FAIL: Finally block must run after catching custom exception"
    exit 1
}
Write-Host "PASS"
exit 0
