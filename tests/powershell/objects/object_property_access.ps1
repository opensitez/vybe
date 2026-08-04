# vybe-test: powershell/objects/object_property_access
$obj = [PSCustomObject]@{ X = 10; Y = 20 }
$sum = $obj.X + $obj.Y
if ($sum -ne 30) {
    Write-Host "FAIL: expected 30, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
