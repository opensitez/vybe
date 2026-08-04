# vybe-test: powershell/control_flow/ternary_operator
$x = 10
$result = $x -gt 5 ? "greater" : "less"
if ($result -ne "greater") {
    Write-Host "FAIL: expected 'greater', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
