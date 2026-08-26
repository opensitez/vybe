# vybe-test: powershell/psalias_properties/psalias_property_null_target_val
$x = 10
$x += 5
$x *= 2
if ($x -eq 30) {
    Write-Host "PASS"
    exit 0
}
Write-Host "FAIL"
exit 1
