# vybe-test: powershell/control_flow/nested_if
$x = 10
$y = 5
$result = ""
if ($x -gt 5) {
    if ($y -gt 3) {
        $result = "both"
    }
}
if ($result -ne "both") {
    Write-Host "FAIL: expected 'both', got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
