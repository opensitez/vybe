# vybe-test: powershell/operators/split_operator
$text = "apple,banana,cherry"
$parts = $text -split ","
if ($parts.Count -ne 3) {
    Write-Host "FAIL: expected 3 parts, got $($parts.Count)"
    exit 1
}
if ($parts[1] -ne "banana") {
    Write-Host "FAIL: expected 'banana', got '$($parts[1])'"
    exit 1
}
Write-Host "PASS"
exit 0
