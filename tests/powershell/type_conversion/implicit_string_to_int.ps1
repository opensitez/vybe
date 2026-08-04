# vybe-test: powershell/type_conversion/implicit_string_to_int
[int]$n = "123"
if ($n -ne 123) {
    Write-Host "FAIL: expected 123, got $n"
    exit 1
}
if ($n -isnot [int]) {
    Write-Host "FAIL: type should be int"
    exit 1
}
Write-Host "PASS"
exit 0
