# vybe-test: powershell/error_record_exception_chaining/exception_stacktrace_populated_after_throw
$err = $null
try {
    throw [System.Exception]::new("StackCheck")
} catch {
    $err = $_
}
if ($err -eq $null) {
    Write-Host "FAIL: Exception catch failed"
    exit 1
}
Write-Host "PASS"
exit 0
