# vybe-test: powershell/argument_completers/argument_completer_exception_isolation
$completer = {
    throw "CompleterError"
}
try {
    &$completer
    Write-Host "FAIL: throwing completer scriptblock expected exception"
    exit 1
} catch {
    Write-Host "PASS"
    exit 0
}
