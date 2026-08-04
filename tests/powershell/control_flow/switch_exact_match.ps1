# vybe-test: powershell/control_flow/switch_exact_match
$day = 3
$result = ""
switch ($day) {
    1 { $result = "Monday" }
    2 { $result = "Tuesday" }
    3 { $result = "Wednesday" }
    default { $result = "Unknown" }
}
if ($result -ne "Wednesday") {
    Write-Host "FAIL: expected 'Wednesday', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
