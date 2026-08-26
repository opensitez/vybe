# vybe-test: powershell/error_record_exception_chaining/getbaseexception_returns_innermost_exception
$e1 = [System.DivideByZeroException]::new("Zero division")
$e2 = [System.Exception]::new("Math error", $e1)
$e3 = [System.Exception]::new("Service error", $e2)
$baseEx = $e3.GetBaseException()
if ($baseEx -ne $e1 -or $baseEx -isnot [System.DivideByZeroException]) {
    Write-Host "FAIL: GetBaseException failed, expected DivideByZeroException, got $($baseEx.GetType().Name)"
    exit 1
}
Write-Host "PASS"
exit 0
