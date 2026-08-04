# vybe-test: powershell/strings/string_split
$str = "a,b,c"
$parts = $str.Split(",")
if ($parts.Count -ne 3) {
    Write-Host "FAIL: expected 3 parts, got $($parts.Count)"
    exit 1
}
if ($parts[1] -ne "b") {
    Write-Host "FAIL: expected 'b', got '$($parts[1])'"
    exit 1
}
Write-Host "PASS"
exit 0
