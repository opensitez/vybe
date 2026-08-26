# vybe-test: powershell/error_record_exception_chaining/exception_hresult_property_preservation
$ex = [System.IO.FileNotFoundException]::new("File missing")
$hr = $ex.HResult
if ($hr -eq 0) {
    Write-Host "FAIL: Exception HResult should not be zero"
    exit 1
}
Write-Host "PASS"
exit 0
