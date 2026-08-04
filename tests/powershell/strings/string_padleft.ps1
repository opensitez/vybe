# vybe-test: powershell/strings/string_padleft
$str = "5"
$result = $str.PadLeft(3, "0")
if ($result -ne "005") {
    Write-Host "FAIL: expected '005', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
