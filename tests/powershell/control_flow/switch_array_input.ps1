# vybe-test: powershell/control_flow/switch_array_input
$colors = @("red", "blue", "red", "green")
$redCount = 0
switch ($colors) {
    "red"   { $redCount++ }
    "blue"  { }
    "green" { }
}
if ($redCount -ne 2) {
    Write-Host "FAIL: expected 2 reds, got $redCount"
    exit 1
}
Write-Host "PASS"
exit 0
