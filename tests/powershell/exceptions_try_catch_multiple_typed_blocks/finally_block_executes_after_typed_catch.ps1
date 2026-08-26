# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/finally_block_executes_after_typed_catch
$finallyRan = $false
try {
    throw [System.ArgumentException]::new()
} catch [System.ArgumentException] {
    # handled
} finally {
    $finallyRan = $true
}
if (-not $finallyRan) {
    Write-Host "FAIL: Finally block after typed catch failed"
    exit 1
}
Write-Host "PASS"
exit 0
