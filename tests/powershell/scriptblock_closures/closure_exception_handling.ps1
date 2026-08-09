# vybe-test: powershell/scriptblock_closures/closure_exception_handling
$errMessage = "ClosureThrow"
$sb = { throw $errMessage }.GetClosure()
try {
    &$sb
    Write-Host "FAIL: throwing closure expected exception"
    exit 1
} catch {
    Write-Host "PASS"
    exit 0
}
