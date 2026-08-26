# vybe-test: powershell/ets_psnoteproperty_dynamic_manipulation/psnoteproperty_iteration_over_all_properties
$obj = [pscustomobject]@{ A = 10; B = 20; C = 30 }
$sum = 0
foreach ($prop in $obj.PSObject.Properties) {
    $sum += $prop.Value
}
if ($sum -ne 60) {
    Write-Host "FAIL: Iteration over PSObject.Properties failed, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
