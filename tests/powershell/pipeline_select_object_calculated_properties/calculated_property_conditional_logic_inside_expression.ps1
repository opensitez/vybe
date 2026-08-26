# vybe-test: powershell/pipeline_select_object_calculated_properties/calculated_property_conditional_logic_inside_expression
$items = @(
    [pscustomobject]@{ Score = 90 },
    [pscustomobject]@{ Score = 55 }
)
$res = @($items | Select-Object @{ N = "Result"; E = { if ($_.Score -ge 70) { "Pass" } else { "Fail" } } })
if ($res[0].Result -ne "Pass" -or $res[1].Result -ne "Fail") {
    Write-Host "FAIL: Calculated property conditional logic failed"
    exit 1
}
Write-Host "PASS"
exit 0
