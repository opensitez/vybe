# vybe-test: powershell/strings/string_tolower
$str = "WORLD"
$result = $str.ToLower()
if ($result -ne "world") {
    Write-Host "FAIL: expected 'world', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
