# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_with_skip_parameter
$res = @(1..5 | Select-Object -Skip 3 @{ N = "Triple"; E = { $_ * 3 } })
if ($res.Length -ne 2 -or $res[0].Triple -ne 12 -or $res[1].Triple -ne 15) {
    Write-Host "FAIL: Calculated property combined with -Skip failed"
    exit 1
}
Write-Host "PASS"
exit 0
