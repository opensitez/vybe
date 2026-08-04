# vybe-test: powershell/control_flow/switch_default
$value = 99
$result = ""
switch ($value) {
    1 { $result = "one" }
    2 { $result = "two" }
    default { $result = "default" }
}
if ($result -ne "default") {
    Write-Host "FAIL: expected 'default', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
