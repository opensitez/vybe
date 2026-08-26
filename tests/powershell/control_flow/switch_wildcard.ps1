# vybe-test: powershell/control_flow/switch_wildcard
$val = "PowerShell"
$matched = ""
switch -Wildcard ($val) {
    "Power*" { $matched = "starts with Power"; break }
    "*Shell" { $matched = "ends with Shell"; break }
}
if ($matched -ne "starts with Power") {
    Write-Host "FAIL: Switch wildcard match failed"
    exit 1
}
Write-Host "PASS"
exit 0
