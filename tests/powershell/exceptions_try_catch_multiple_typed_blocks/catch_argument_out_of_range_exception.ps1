# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_argument_out_of_range_exception
$caught = $false
try {
    $str = "abc"
    $x = $str.Substring(10, 2)
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Catching ArgumentOutOfRangeException failed"
    exit 1
}
Write-Host "PASS"
exit 0
