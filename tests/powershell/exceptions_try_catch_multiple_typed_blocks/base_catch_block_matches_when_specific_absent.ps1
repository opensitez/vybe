# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/base_catch_block_matches_when_specific_absent
$caughtType = ""
try {
    throw [System.OverflowException]::new()
} catch [System.DivideByZeroException] {
    $caughtType = "DivideByZero"
} catch [System.ArithmeticException] {
    $caughtType = "Arithmetic"
} catch {
    $caughtType = "Generic"
}
if ($caughtType -ne "Arithmetic") {
    Write-Host "FAIL: Base catch block match failed, got '$caughtType'"
    exit 1
}
Write-Host "PASS"
exit 0
