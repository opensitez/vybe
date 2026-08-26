# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/fallback_to_catch_all_block
$caughtType = ""
try {
    throw [System.IO.FileNotFoundException]::new()
} catch [System.DivideByZeroException] {
    $caughtType = "DivideByZero"
} catch [System.FormatException] {
    $caughtType = "Format"
} catch {
    $caughtType = "Generic"
}
if ($caughtType -ne "Generic") {
    Write-Host "FAIL: Fallback to catch-all block failed, got '$caughtType'"
    exit 1
}
Write-Host "PASS"
exit 0
