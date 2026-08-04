# vybe-test: powershell/control_flow/switch_regex
$value = "test123"
$result = ""
switch -regex ($value) {
    "^\d+$" { $result = "numeric" }
    "^[a-z]+\d+$" { $result = "alphanumeric" }
    default { $result = "unknown" }
}
if ($result -ne "alphanumeric") {
    Write-Host "FAIL: expected 'alphanumeric', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
