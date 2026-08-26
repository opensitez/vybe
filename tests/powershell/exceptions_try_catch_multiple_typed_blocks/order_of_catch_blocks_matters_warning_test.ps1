# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/order_of_catch_blocks_matters_warning_test
$caught = ""
try {
    throw [System.ArgumentNullException]::new()
} catch [System.ArgumentNullException] {
    $caught = "Specific"
} catch [System.ArgumentException] {
    $caught = "Base"
}
if ($caught -ne "Specific") {
    Write-Host "FAIL: Specific catch block before base catch block failed"
    exit 1
}
Write-Host "PASS"
exit 0
