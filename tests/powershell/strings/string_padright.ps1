# vybe-test: powershell/strings/string_padright
$str = "5"
$result = $str.PadRight(3, "0")
if ($result -ne "500") {
    Write-Host "FAIL: expected '500', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
