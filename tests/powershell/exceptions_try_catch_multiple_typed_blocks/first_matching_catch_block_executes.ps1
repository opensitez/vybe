# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/first_matching_catch_block_executes
$caughtType = ""
try {
    throw [System.DivideByZeroException]::new()
} catch [System.DivideByZeroException] {
    $caughtType = "DivideByZero"
} catch [System.ArithmeticException] {
    $caughtType = "Arithmetic"
} catch {
    $caughtType = "Generic"
}
if ($caughtType -ne "DivideByZero") {
    Write-Host "FAIL: First matching catch block failed, got '$caughtType'"
    exit 1
}
Write-Host "PASS"
exit 0
