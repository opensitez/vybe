# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_with_first_parameter
$res = @(1..10 | Select-Object -First 3 @{ N = "Double"; E = { $_ * 2 } })
if ($res.Length -ne 3 -or $res[0].Double -ne 2 -or $res[2].Double -ne 6) {
    Write-Host "FAIL: Calculated property combined with -First failed"
    exit 1
}
Write-Host "PASS"
exit 0
