# vybe-test: powershell/control_flow/switch_wildcard
$value = "PowerShell"
$result = ""
switch -wildcard ($value) {
    "Power*" { $result = "starts with Power" }
    "*Shell" { $result = "ends with Shell" }
    default { $result = "no match" }
}
if ($result -ne "starts with Power") {
    Write-Host "FAIL: expected 'starts with Power', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
