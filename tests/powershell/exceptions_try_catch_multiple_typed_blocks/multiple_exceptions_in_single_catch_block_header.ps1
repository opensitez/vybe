# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/multiple_exceptions_in_single_catch_block_header
$caughtType = ""
try {
    throw [System.FormatException]::new()
} catch [System.FormatException], [System.ArgumentNullException] {
    $caughtType = "FormatOrNull"
} catch {
    $caughtType = "Other"
}
if ($caughtType -ne "FormatOrNull") {
    Write-Host "FAIL: Comma-separated typed catch block failed, got '$caughtType'"
    exit 1
}
Write-Host "PASS"
exit 0
