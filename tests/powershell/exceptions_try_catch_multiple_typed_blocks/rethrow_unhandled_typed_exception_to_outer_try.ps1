# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/rethrow_unhandled_typed_exception_to_outer_try
$outerCaught = ""
try {
    try {
        throw [System.TimeoutException]::new()
    } catch [System.ArgumentException] {
        $outerCaught = "Inner"
    }
} catch [System.TimeoutException] {
    $outerCaught = "OuterTimeout"
}
if ($outerCaught -ne "OuterTimeout") {
    Write-Host "FAIL: Unhandled typed exception propagation to outer try failed, got '$outerCaught'"
    exit 1
}
Write-Host "PASS"
exit 0
