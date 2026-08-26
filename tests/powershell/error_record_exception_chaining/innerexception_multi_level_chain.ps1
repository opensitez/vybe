# vybe-test: powershell/error_record_exception_chaining/innerexception_multi_level_chain
$e1 = [System.FormatException]::new("Invalid port number")
$e2 = [System.ArgumentException]::new("Config invalid", $e1)
$e3 = [System.InvalidOperationException]::new("App start failed", $e2)
if ($e3.InnerException -ne $e2 -or $e3.InnerException.InnerException -ne $e1) {
    Write-Host "FAIL: Multi-level InnerException chain failed"
    exit 1
}
Write-Host "PASS"
exit 0
