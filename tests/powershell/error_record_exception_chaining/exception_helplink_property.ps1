# vybe-test: powershell/error_record_exception_chaining/exception_helplink_property
$ex = [System.Exception]::new("Test")
$ex.HelpLink = "https://docs.example.com/errors/404"
if ($ex.HelpLink -ne "https://docs.example.com/errors/404") {
    Write-Host "FAIL: Exception HelpLink property failed"
    exit 1
}
Write-Host "PASS"
exit 0
