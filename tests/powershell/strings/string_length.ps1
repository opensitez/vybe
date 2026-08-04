# vybe-test: powershell/strings/string_length
$str = "PowerShell"
$length = $str.Length
if ($length -ne 10) {
    Write-Host "FAIL: expected 10, got $length"
    exit 1
}
Write-Host "PASS"
exit 0
