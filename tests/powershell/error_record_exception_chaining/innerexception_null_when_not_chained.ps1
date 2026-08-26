# vybe-test: powershell/error_record_exception_chaining/innerexception_null_when_not_chained
$ex = [System.Exception]::new("Solo")
if ($ex.InnerException -ne $null) {
    Write-Host "FAIL: Unchained exception InnerException must be null"
    exit 1
}
Write-Host "PASS"
exit 0
