# vybe-test: powershell/data_conversion/implicit_conversion
[int]$n = "123"
if ($n -ne 123) {
    Write-Host "FAIL: Implicit conversion failed"
    exit 1
}
Write-Host "PASS"
exit 0
