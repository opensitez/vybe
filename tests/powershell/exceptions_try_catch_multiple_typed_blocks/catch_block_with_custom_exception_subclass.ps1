# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_block_with_custom_exception_subclass
class MySpecificException : System.InvalidOperationException {
    MySpecificException([string]$m) : base($m) {}
}
$caught = ""
try {
    throw [MySpecificException]::new("CustomCrash")
} catch [MySpecificException] {
    $caught = "Custom"
} catch [System.InvalidOperationException] {
    $caught = "InvalidOp"
}
if ($caught -ne "Custom") {
    Write-Host "FAIL: Custom exception typed catch failed"
    exit 1
}
Write-Host "PASS"
exit 0
