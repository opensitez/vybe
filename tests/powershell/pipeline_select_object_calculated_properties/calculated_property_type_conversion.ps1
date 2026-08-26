# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_type_conversion
$item = [pscustomobject]@{ NumberStr = "42" }
$res = $item | Select-Object @{ N = "ParsedInt"; E = { [int]$_.NumberStr } }
if ($res.ParsedInt -ne 42 -or $res.ParsedInt -isnot [int]) {
    Write-Host "FAIL: Calculated property type conversion failed"
    exit 1
}
Write-Host "PASS"
exit 0
