# vybe-test: powershell/error_record_exception_chaining/exception_source_property
$ex = [System.Exception]::new("Test")
$ex.Source = "MyCustomModule"
if ($ex.Source -ne "MyCustomModule") {
    Write-Host "FAIL: Exception Source property failed"
    exit 1
}
Write-Host "PASS"
exit 0
