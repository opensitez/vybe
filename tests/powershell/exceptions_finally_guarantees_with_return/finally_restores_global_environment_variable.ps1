# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_restores_global_environment_variable
$orig = $env:TEMP_FLAG
try {
    $env:TEMP_FLAG = "ACTIVE"
    return "Done"
} finally {
    $env:TEMP_FLAG = $orig
}
if ($env:TEMP_FLAG -ne $orig) {
    Write-Host "FAIL: Environment variable restore in finally failed"
    exit 1
}
Write-Host "PASS"
exit 0
