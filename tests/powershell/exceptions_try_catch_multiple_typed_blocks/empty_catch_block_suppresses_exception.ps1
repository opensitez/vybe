# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/empty_catch_block_suppresses_exception
$val = 0
try {
    throw [System.Exception]::new("Suppressed")
    $val = 1
} catch [System.Exception] {}
if ($val -ne 0) {
    Write-Host "FAIL: Empty catch block failed"
    exit 1
}
Write-Host "PASS"
exit 0
