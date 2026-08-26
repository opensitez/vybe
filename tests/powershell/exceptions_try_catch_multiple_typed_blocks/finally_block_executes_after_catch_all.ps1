# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/finally_block_executes_after_catch_all
$finallyRan = $false
try {
    throw "StringError"
} catch {
    # handled
} finally {
    $finallyRan = $true
}
if (-not $finallyRan) {
    Write-Host "FAIL: Finally block after catch-all failed"
    exit 1
}
Write-Host "PASS"
exit 0
