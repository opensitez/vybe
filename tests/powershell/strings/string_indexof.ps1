# vybe-test: powershell/strings/string_indexof
$str = "PowerShell"
$index = $str.IndexOf("Shell")
if ($index -ne 5) {
    Write-Host "FAIL: expected 5, got $index"
    exit 1
}
Write-Host "PASS"
exit 0
