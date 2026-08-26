# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_block_modifying_external_variable
$status = "INIT"
try {
    throw [System.TimeoutException]::new()
} catch [System.TimeoutException] {
    $status = "TIMED_OUT"
}
if ($status -ne "TIMED_OUT") {
    Write-Host "FAIL: Catch block modifying variable failed"
    exit 1
}
Write-Host "PASS"
exit 0
