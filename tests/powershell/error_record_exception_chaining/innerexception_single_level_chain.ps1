# vybe-test: powershell/error_record_exception_chaining/innerexception_single_level_chain
$inner = [System.IO.FileNotFoundException]::new("Missing file", "data.txt")
$outer = [System.InvalidOperationException]::new("Config load failed", $inner)
if ($outer.InnerException -ne $inner -or $outer.InnerException.Message -ne "Missing file") {
    Write-Host "FAIL: Single level InnerException chain failed"
    exit 1
}
Write-Host "PASS"
exit 0
