# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_io_exception_hierarchy
$caughtType = ""
try {
    throw [System.IO.DirectoryNotFoundException]::new("Dir missing")
} catch [System.IO.DirectoryNotFoundException] {
    $caughtType = "DirectoryNotFound"
} catch [System.IO.IOException] {
    $caughtType = "IO"
}
if ($caughtType -ne "DirectoryNotFound") {
    Write-Host "FAIL: IO exception hierarchy catch failed, got '$caughtType'"
    exit 1
}
Write-Host "PASS"
exit 0
